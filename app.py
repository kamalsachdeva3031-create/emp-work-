"""Rack Inventory: local-network multi-employee stock app, Python standard library + openpyxl."""
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote
from pathlib import Path
from io import BytesIO
import sqlite3, json, threading, datetime, socket, os, sys, traceback, re, secrets, hashlib, hmac, time, http.cookies
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import qrcode

BASE=Path(__file__).resolve().parent
DATA=Path(os.environ.get('DATA_DIR',str(BASE))).resolve()
DATA.mkdir(parents=True,exist_ok=True)
DB=DATA/'inventory.db'
LOCK=threading.RLock()

def connect():
    c=sqlite3.connect(DB, timeout=30)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA journal_mode=WAL')
    c.execute('PRAGMA busy_timeout=30000')
    return c

def setup():
    with connect() as db:
        db.executescript('''
        CREATE TABLE IF NOT EXISTS racks (rack TEXT PRIMARY KEY, capacity INTEGER NOT NULL DEFAULT 20);
        CREATE TABLE IF NOT EXISTS stock (rack TEXT NOT NULL, model TEXT NOT NULL COLLATE NOCASE, qty INTEGER NOT NULL DEFAULT 0, PRIMARY KEY (rack, model), FOREIGN KEY(rack) REFERENCES racks(rack));
        CREATE TABLE IF NOT EXISTS movements (id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, employee TEXT NOT NULL, action TEXT NOT NULL, rack TEXT NOT NULL, model TEXT NOT NULL, qty INTEGER NOT NULL, before_qty INTEGER, after_qty INTEGER, note TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY COLLATE NOCASE, pass_hash TEXT NOT NULL, salt TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1);
        CREATE TABLE IF NOT EXISTS sessions (token_hash TEXT PRIMARY KEY, username TEXT NOT NULL, role TEXT NOT NULL, until INTEGER NOT NULL);
        ''')

# Admin password is set as an environment variable, not stored in files/repository.
ADMIN_PASSWORD=os.environ.get('ADMIN_PASSWORD','')
if not ADMIN_PASSWORD:
    raise RuntimeError('Set ADMIN_PASSWORD before starting. Use a unique strong password.')

# Server-side sessions. A cookie is never trusted without looking up its hash.
def hash_password(password, salt):
    return hashlib.pbkdf2_hmac('sha256',password.encode(),bytes.fromhex(salt),260_000).hex()

def user_for(h):
    cookie=http.cookies.SimpleCookie()
    try: cookie.load(h.headers.get('Cookie',''))
    except: return None
    token=cookie.get('rackflow_session')
    if not token:return None
    hashed=hashlib.sha256(token.value.encode()).hexdigest()
    with connect() as db:
        row=db.execute('SELECT username,role FROM sessions WHERE token_hash=? AND until>?',(hashed,int(time.time()))).fetchone()
        if row and row['role']=='employee':
            active=db.execute('SELECT active FROM users WHERE username=?',(row['username'],)).fetchone()
            if not active or not active['active']:return None
    return dict(row) if row else None

def login_user(h,data):
    username=cleanstr(data.get('username')).lower()
    password=str(data.get('password') or '')
    if len(username)>50 or len(password)>256 or not username or not password:raise ValueError('Invalid login')
    if username=='admin':
        valid=hmac.compare_digest(password,ADMIN_PASSWORD)
        role='admin'
    else:
        role='employee'
        with connect() as db: u=db.execute('SELECT salt,pass_hash,active FROM users WHERE username=?',(username,)).fetchone()
        valid=bool(u and u['active'] and hmac.compare_digest(hash_password(password,u['salt']),u['pass_hash']))
    if not valid:
        time.sleep(0.3)
        raise ValueError('Incorrect username or password')
    token=secrets.token_urlsafe(40)
    with connect() as db:
        db.execute('DELETE FROM sessions WHERE until<?',(int(time.time()),))
        db.execute('INSERT INTO sessions(token_hash,username,role,until) VALUES(?,?,?,?)',(hashlib.sha256(token.encode()).hexdigest(),username,role,int(time.time())+60*60*24*14))
        db.commit()
    payload=json.dumps({'ok':True,'username':username,'role':role}).encode()
    h.send_response(200)
    h.send_header('Content-Type','application/json')
    h.send_header('Cache-Control','no-store')
    h.send_header('Set-Cookie',f'rackflow_session={token}; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=1209600')
    h.send_header('Content-Length',str(len(payload)))
    h.end_headers();h.wfile.write(payload)

def admin_users():
    with connect() as db:
        return [dict(row) for row in db.execute('SELECT username,active FROM users ORDER BY username')]

def create_user(data):
    name=cleanstr(data.get('username')).lower()
    password=str(data.get('password') or '')
    if not re.fullmatch(r'[a-z0-9][a-z0-9._-]{1,31}',name) or name=='admin':raise ValueError('Employee username must be 2–32 letters, numbers, dot, underscore or dash.')
    if len(password)<8 or len(password)>128:raise ValueError('Employee password must have 8–128 characters')
    salt=secrets.token_hex(16)
    with connect() as db:
        db.execute('INSERT INTO users(username,pass_hash,salt,active) VALUES(?,?,?,1) ON CONFLICT(username) DO UPDATE SET pass_hash=excluded.pass_hash,salt=excluded.salt,active=1',(name,hash_password(password,salt),salt))
        db.commit()
    return {'ok':True,'username':name}

def set_user_active(data):
    name=cleanstr(data.get('username')).lower()
    if name=='admin':raise ValueError('Cannot change admin here')
    active=1 if data.get('active') else 0
    with connect() as db:
        r=db.execute('UPDATE users SET active=? WHERE username=?',(active,name))
        if not r.rowcount:raise ValueError('Employee not found')
        if not active:db.execute('DELETE FROM sessions WHERE username=?',(name,))
        db.commit()
    return {'ok':True}

def stamp(): return datetime.datetime.now().astimezone().isoformat(timespec='seconds')
def respond_json(h,data,status=200):
    payload=json.dumps(data,ensure_ascii=False).encode()
    h.send_response(status);h.send_header('Content-Type','application/json; charset=utf-8');h.send_header('Content-Length',str(len(payload)));h.send_header('Cache-Control','no-store');h.end_headers();h.wfile.write(payload)
def error(h,message,status=400):respond_json(h,{'error':message},status)
def cleanstr(v):return str(v or '').strip()
def positive(v,name='Quantity'):
    try:
        n=int(v)
        if str(v).strip() not in (str(n),):
            if not isinstance(v,int): raise ValueError()
        if n<=0: raise ValueError()
        return n
    except:raise ValueError(f'{name} must be a positive whole number')

def make_state(query=None):
    with connect() as db:
        rack_rows=db.execute('SELECT rack,capacity FROM racks ORDER BY rack COLLATE NOCASE').fetchall()
        all_stock=db.execute('SELECT rack,model,qty FROM stock ORDER BY model COLLATE NOCASE').fetchall()
        moves=db.execute('SELECT * FROM movements ORDER BY id DESC LIMIT 100').fetchall()
    by_rack={r['rack']:{'rack':r['rack'],'capacity':r['capacity'],'used':0,'models':[]} for r in rack_rows}
    for a in all_stock:
        if a['rack'] in by_rack:
            by_rack[a['rack']]['used']+=a['qty']
            by_rack[a['rack']]['models'].append({'model':a['model'],'qty':a['qty']})
    racks=list(by_rack.values())
    for r in racks:
        r['available']=max(0,r['capacity']-r['used'])
        r['status']='Empty' if r['used']==0 else ('Full' if r['used']>=r['capacity'] else 'Occupied')
    return {'racks':racks,'movements':[dict(m) for m in moves],'summary':{'racks':len(racks),'empty':sum(r['used']==0 for r in racks),'occupied':sum(r['used']>0 for r in racks),'units':sum(r['used'] for r in racks)}}

def mutate(data, direction):
    rack=cleanstr(data.get('rack')).upper(); model=cleanstr(data.get('model')); employee=cleanstr(data.get('employee')) or 'Employee'
    qty=positive(data.get('qty'));note=cleanstr(data.get('note'))
    if not rack or not model:raise ValueError('Select rack and model')
    with LOCK,connect() as db:
        db.execute('BEGIN IMMEDIATE')
        cap=db.execute('SELECT capacity FROM racks WHERE rack=?',(rack,)).fetchone()
        if cap is None: raise ValueError('Unknown rack. Import the rack in Excel first.')
        row=db.execute('SELECT qty FROM stock WHERE rack=? AND model=?',(rack,model)).fetchone()
        if not row and direction==-1: raise ValueError('This model is not assigned to this rack.')
        before=row['qty'] if row else 0
        used=db.execute('SELECT COALESCE(SUM(qty),0) FROM stock WHERE rack=?',(rack,)).fetchone()[0]
        after=before+direction*qty
        if after<0:raise ValueError(f'Only {before} pieces available for {model} in rack {rack}')
        if direction==1 and used+qty>cap['capacity']:raise ValueError(f'Rack {rack} has room for only {max(0,cap["capacity"]-used)} pieces')
        db.execute('INSERT INTO stock(rack,model,qty) VALUES(?,?,?) ON CONFLICT(rack,model) DO UPDATE SET qty=excluded.qty',(rack,model,after))
        db.execute('INSERT INTO movements (at,employee,action,rack,model,qty,before_qty,after_qty,note) VALUES (?,?,?,?,?,?,?,?,?)',(stamp(),employee,'IN' if direction==1 else 'OUT',rack,model,qty,before,after,note))
        db.commit()
        return {'ok':True,'rack':rack,'model':model,'before':before,'after':after}

def make_excel(template=False):
    wb=Workbook();ws=wb.active;ws.title='Stock Master'
    ws.append(['Rack Number','Model','Quantity','Capacity'])
    if template:
        for row in [('A-001','iPhone 13 D-15',15,20),('A-002','iPhone 14 CBK-1',0,20),('B-001','Samsung A55 C-1',10,20)]:ws.append(row)
    else:
        state=make_state()
        for rack in state['racks']:
            if rack['models']:
                for s in rack['models']:ws.append([rack['rack'],s['model'],s['qty'],rack['capacity']])
            else:ws.append([rack['rack'],'',0,rack['capacity']])
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    for cell in ws[1]:cell.font=Font(color='FFFFFF',bold=True);cell.fill=PatternFill('solid',fgColor='17364B');cell.alignment=Alignment(horizontal='center')
    for col,width in enumerate([22,43,18,18],1):ws.column_dimensions[get_column_letter(col)].width=width
    info=wb.create_sheet('Instructions')
    for x in [
        ['RACK STOCK MASTER — INSTRUCTIONS'],
        ['1. Each row assigns one phone-cover model to a rack.'],
        ['2. Edit Rack Number, Model, Quantity, Capacity and upload the .xlsx in the app.'],
        ['3. Repeat rack numbers for multiple models in one rack (capacity must be identical).'],
        ['4. To reserve an empty rack, enter rack number, blank model, quantity 0 and capacity.'],
        ['5. Upload REPLACES current stock numbers; download latest first to keep employee changes.'],
        ['6. Employee IN and OUT changes are logged separately.'],
        ['7. Racks must have stock <= capacity; quantities cannot be negative.'],
    ]: info.append(x)
    info.column_dimensions['A'].width=110
    buf=BytesIO();wb.save(buf);return buf.getvalue()

def import_excel(raw):
    if len(raw)>12*1024*1024:raise ValueError('Excel file must be under 12 MB')
    try: wb=load_workbook(BytesIO(raw),read_only=True,data_only=True)
    except Exception:raise ValueError('Please upload a valid .xlsx file')
    ws=wb['Stock Master'] if 'Stock Master' in wb.sheetnames else wb.active
    rows=ws.iter_rows(values_only=True)
    try: headers=[re.sub(r'[^a-z0-9]','',str(c or '').lower()) for c in next(rows)]
    except StopIteration: raise ValueError('Excel has no rows')
    def index(options):
        for v in options:
            if v in headers:return headers.index(v)
        return None
    rix=index(('racknumber','rack','rackid','location'));mix=index(('model','modelname','product','item'));qix=index(('quantity','qty','stock','stockquantity'));cix=index(('capacity','rackcapacity','maxqty'))
    if None in (rix,mix,qix): raise ValueError('Need columns Rack Number, Model, Quantity. Capacity is optional.')
    racks={};records={};issues=[]
    for i,row in enumerate(rows,2):
        if not any(v is not None and str(v).strip() for v in row):continue
        def at(j):return row[j] if j is not None and j<len(row) else None
        rack=cleanstr(at(rix)).upper();model=cleanstr(at(mix))
        if not rack:issues.append(f'Row {i}: rack is blank');continue
        try:
            rawqty=at(qix); qty=int(rawqty if rawqty is not None and str(rawqty).strip() else 0)
            rawcap=at(cix);cap=int(rawcap if rawcap is not None and str(rawcap).strip() else 20)
            if qty<0 or cap<=0 or qty!=float(rawqty if rawqty is not None and str(rawqty).strip() else 0):raise ValueError()
        except:issues.append(f'Row {i}: invalid quantity or capacity');continue
        if rack in racks and racks[rack]!=cap:issues.append(f'Row {i}: different capacity for rack {rack}')
        racks[rack]=cap
        if qty and not model:issues.append(f'Row {i}: model missing for positive quantity')
        if model:
            key=(rack,model.lower()); rec=records.setdefault(key,{'rack':rack,'model':model,'qty':0});rec['qty']+=qty
    usage={r:0 for r in racks}
    for rec in records.values():usage[rec['rack']]+=rec['qty']
    for r,used in usage.items():
        if used>racks[r]:issues.append(f'Rack {r}: {used} units exceeds capacity {racks[r]}')
    if issues:raise ValueError('Excel validation failed: '+'; '.join(issues[:12]))
    if not racks:raise ValueError('No racks found in Excel')
    with LOCK,connect() as db:
        db.execute('BEGIN IMMEDIATE');db.execute('DELETE FROM stock');db.execute('DELETE FROM racks')
        db.executemany('INSERT INTO racks(rack,capacity) VALUES(?,?)',racks.items())
        db.executemany('INSERT INTO stock(rack,model,qty) VALUES(?,?,?)',[(r['rack'],r['model'],r['qty']) for r in records.values()])
        db.execute('INSERT INTO movements(at,employee,action,rack,model,qty,before_qty,after_qty,note) VALUES(?,?,?,?,?,?,?,?,?)',(stamp(),'Admin','EXCEL IMPORT','ALL','Master stock replaced',sum(usage.values()),None,None,f'{len(racks)} racks; {len(records)} model-rack records'))
        db.commit()
    return {'ok':True,'racks':len(racks),'items':len(records),'units':sum(usage.values())}

class Handler(BaseHTTPRequestHandler):
    def log_message(self,fmt,*args): print('%s %s'%(self.address_string(),fmt%args))
    def do_GET(self):
        path=urlparse(self.path).path
        identity=user_for(self)
        if path=='/':
            raw=(BASE/'static'/'index.html').read_bytes();self.binary(raw,'text/html; charset=utf-8');return
        if path=='/api/me':respond_json(self,identity or {'role':'guest'});return
        if path=='/health':respond_json(self,{'ok':True});return
        if not identity:return error(self,'Please log in',401)
        if path=='/jsQR.js':
            p=BASE/'static'/'jsQR.js'
            if p.exists(): self.binary(p.read_bytes(),'application/javascript');return
            error(self,'Scanner library not installed',404);return
        if path=='/api/state':
            s=make_state()
            if identity['role']!='admin':s.pop('movements',None)
            respond_json(self,s);return
        if path=='/api/users':
            if identity['role']!='admin':return error(self,'Admin only',403)
            respond_json(self,admin_users());return
        if path in ('/api/export','/api/template'):
            if identity['role']!='admin':return error(self,'Admin only',403)
            self.binary(make_excel(path.endswith('template')),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',attachment='rack_stock_template.xlsx' if path.endswith('template') else 'rack_stock_current.xlsx');return
        if path.startswith('/api/qr/'):
            if identity['role']!='admin':return error(self,'Admin only',403)
            rack=path[len('/api/qr/'):].strip().upper()
            if not rack or len(rack)>80:return error(self,'Invalid rack')
            img=qrcode.make(rack);buf=BytesIO();img.save(buf,format='PNG');self.binary(buf.getvalue(),'image/png');return
        error(self,'Not found',404)
    def binary(self,payload,ctype,attachment=None):
        self.send_response(200);self.send_header('Content-Type',ctype);self.send_header('Content-Length',str(len(payload)))
        if attachment:self.send_header('Content-Disposition',f'attachment; filename="{attachment}"')
        self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(payload)
    def do_POST(self):
        path=urlparse(self.path).path
        try:
            origin=self.headers.get('Origin')
            if origin:
                from urllib.parse import urlsplit
                target=urlsplit(origin)
                host=self.headers.get('X-Forwarded-Host',self.headers.get('Host','')).split(',')[0].strip()
                if target.netloc.lower()!=host.lower():return error(self,'Invalid request origin',403)
            size=int(self.headers.get('Content-Length','0'))
            if size<=0 or size>12*1024*1024:raise ValueError('Request is empty or too large')
            raw=self.rfile.read(size)
            identity=user_for(self)
            if path=='/api/login':
                try:data=json.loads(raw)
                except:raise ValueError('Invalid request')
                login_user(self,data);return
            if not identity:return error(self,'Please log in',401)
            if path=='/api/logout':
                cookie=http.cookies.SimpleCookie()
                try:cookie.load(self.headers.get('Cookie',''))
                except:pass
                token=cookie.get('rackflow_session')
                if token:
                    with connect() as db:db.execute('DELETE FROM sessions WHERE token_hash=?',(hashlib.sha256(token.value.encode()).hexdigest(),));db.commit()
                payload=b'{"ok":true}'
                self.send_response(200);self.send_header('Set-Cookie','rackflow_session=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Strict');self.send_header('Content-Type','application/json');self.send_header('Content-Length',str(len(payload)));self.end_headers();self.wfile.write(payload);return
            if path=='/api/import':
                if identity['role']!='admin':return error(self,'Admin only',403)
                respond_json(self,import_excel(raw));return
            try:data=json.loads(raw)
            except:raise ValueError('Invalid request')
            if path=='/api/users':
                if identity['role']!='admin':return error(self,'Admin only',403)
                respond_json(self,create_user(data));return
            if path=='/api/user-status':
                if identity['role']!='admin':return error(self,'Admin only',403)
                respond_json(self,set_user_active(data));return
            if path=='/api/out':
                data['employee']=identity['username']
                respond_json(self,mutate(data,-1));return
            if path=='/api/in':
                data['employee']=identity['username']
                respond_json(self,mutate(data,1));return
            error(self,'Not found',404)
        except ValueError as ex:error(self,str(ex))
        except Exception as ex:traceback.print_exc();error(self,'Server error. Please retry.',500)

def local_ip():
    try:
        s=socket.socket(socket.AF_INET,socket.SOCK_DGRAM);s.connect(('8.8.8.8',80));ip=s.getsockname()[0];s.close();return ip
    except:return 'YOUR-PC-IP'
if __name__=='__main__':
    setup();port=int(os.environ.get('PORT','8765'))
    print('\nRACK INVENTORY APP\nPC: http://localhost:%s\nEmployees on same Wi-Fi: http://%s:%s\n'%(port,local_ip(),port),flush=True)
    ThreadingHTTPServer(('0.0.0.0',port),Handler).serve_forever()
