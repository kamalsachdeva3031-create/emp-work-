"""Rack Inventory: local-network multi-employee stock app, Python standard library + openpyxl."""
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote
from pathlib import Path
from io import BytesIO
import sqlite3, json, threading, datetime, socket, os, sys, traceback, re, secrets, hashlib, hmac, time, http.cookies, uuid, shutil, tempfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import qrcode

BASE=Path(__file__).resolve().parent
IS_RAILWAY=any(os.getenv(k) for k in ('RAILWAY_ENVIRONMENT','RAILWAY_ENVIRONMENT_ID','RAILWAY_PROJECT_ID','RAILWAY_DEPLOYMENT_ID'))
# IMPORTANT: Never fall back to /app on Railway; its filesystem is disposable.
if IS_RAILWAY:
    if os.getenv('DATA_DIR') != '/data':
        raise RuntimeError('STORAGE SAFETY STOP: Railway DATA_DIR must be /data. No database was changed.')
    if os.getenv('RAILWAY_VOLUME_MOUNT_PATH') != '/data':
        raise RuntimeError('STORAGE SAFETY STOP: no Railway volume mounted at /data. Attach original volume; do not create a new one.')
DATA=Path(os.environ.get('DATA_DIR',str(BASE))).resolve()
if not IS_RAILWAY:
    DATA.mkdir(parents=True,exist_ok=True)
if not DATA.is_dir():
    raise RuntimeError(f'STORAGE SAFETY STOP: persistent data folder {DATA} does not exist.')
DB=DATA/'inventory.db'
BACKUPS=DATA/'backups'
GUARD=DATA/'.rackflow_volume_identity'
LOCK=threading.RLock()
GUARD_ACTIVE=False
EXPECTED_ID=None
REQUIRED_TABLES={'racks','stock','movements','users','sessions'}

def verify_existing_db():
    # Read-only verification: sqlite3.connect(DB) without mode=ro would CREATE an empty DB.
    if not DB.is_file() or DB.stat().st_size < 4096:
        raise RuntimeError(f'STORAGE SAFETY STOP: {DB} is missing/empty. The app will NOT create a new database. Attach the original volume or restore a backup.')
    with sqlite3.connect(f'file:{DB}?mode=ro',uri=True) as c:
        tables={r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if not REQUIRED_TABLES.issubset(tables):
            raise RuntimeError(f'STORAGE SAFETY STOP: {DB} has missing inventory tables: {sorted(REQUIRED_TABLES-tables)}')
        result=c.execute('PRAGMA quick_check').fetchone()[0]
        if result!='ok':
            raise RuntimeError(f'STORAGE SAFETY STOP: database integrity issue: {result}')

def verify_mount():
    if not IS_RAILWAY:return
    # Railway injects mount information only at runtime, not at build time.
    try:
        mounts=Path('/proc/self/mountinfo').read_text().splitlines()
        mounted=any(line.split(' - ',1)[0].split()[4]=='/data' for line in mounts if ' - ' in line)
    except OSError as e:
        raise RuntimeError('STORAGE SAFETY STOP: unable to verify Railway volume mount') from e
    if not mounted:
        raise RuntimeError('STORAGE SAFETY STOP: /data is not an actual mounted volume. No database was changed.')

def connect():
    if GUARD_ACTIVE:
        if not DB.is_file() or not GUARD.is_file() or GUARD.read_text().strip()!=EXPECTED_ID:
            raise RuntimeError('STORAGE SAFETY STOP: database/volume identity changed. Refusing to write.')
    c=sqlite3.connect(DB,timeout=30)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA journal_mode=WAL')
    c.execute('PRAGMA busy_timeout=30000')
    return c

def backup_db(category='latest'):
    """Consistent SQLite backup that includes uncheckpointed WAL changes. Stored on /data."""
    BACKUPS.mkdir(parents=True,exist_ok=True)
    fd, tmp=tempfile.mkstemp(prefix='.rackflow_tmp_',suffix='.db',dir=BACKUPS)
    os.close(fd)
    try:
        src=sqlite3.connect(f'file:{DB}?mode=ro',uri=True)
        dst=sqlite3.connect(tmp)
        try:src.backup(dst)
        finally:dst.close();src.close()
        with sqlite3.connect(f'file:{tmp}?mode=ro',uri=True) as chk:
            if chk.execute('PRAGMA quick_check').fetchone()[0]!='ok':
                raise RuntimeError('SQLite backup verification failed')
        now=datetime.datetime.now(datetime.timezone.utc)
        if category=='pre-import':
            name=f'pre-import_{now.strftime("%Y%m%d_%H%M%S_%f")}.db'
        elif category=='manual':
            name=f'manual_{now.strftime("%Y%m%d_%H%M%S_%f")}.db'
        elif category=='pre-upgrade':
            name=f'pre-upgrade_{now.strftime("%Y%m%d_%H%M%S_%f")}.db'
        else:
            name='latest.db'
        out=BACKUPS/name
        os.replace(tmp,out)
        # Rotating point-in-time recovery copies, distinct from latest.
        if category=='latest':
            hour=BACKUPS/f'hourly_{now.strftime("%Y%m%d_%H")}.db'
            day=BACKUPS/f'daily_{now.strftime("%Y%m%d")}.db'
            if not hour.exists():shutil.copyfile(out,hour)
            if not day.exists():shutil.copyfile(out,day)
        for prefix,retain in [('hourly_',24),('daily_',14),('pre-import_',15),('manual_',10),('pre-upgrade_',5)]:
            old=sorted(BACKUPS.glob(prefix+'*.db'))
            for stale in old[:-retain]:stale.unlink()
        return out
    finally:
        if os.path.exists(tmp):os.unlink(tmp)

def backup_after_change():
    try:backup_db('latest')
    except Exception as e:
        print('WARNING: Stock/account saved, but automatic backup failed:',repr(e),flush=True)

def setup():
    verify_mount()
    verify_existing_db() # stop instead of initializing a fresh/incorrect database
    global GUARD_ACTIVE,EXPECTED_ID
    # Look for the volume's identity without modifying its database.
    with sqlite3.connect(f'file:{DB}?mode=ro',uri=True) as ro:
        exists=ro.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='rackflow_meta'").fetchone()
        row=ro.execute('SELECT value FROM rackflow_meta WHERE key=?',('volume_id',)).fetchone() if exists else None
        db_id=row[0] if row else None
    marker_id=GUARD.read_text().strip() if GUARD.is_file() else None
    pinned=os.getenv("RACKFLOW_EXPECTED_VOLUME_ID","").strip()
    if pinned and (not db_id or not marker_id):
        raise RuntimeError("STORAGE SAFETY STOP: pinned volume identity not found. Attach original Railway volume.")
    if not db_id and not marker_id:
        # One-time migration of an EXISTING valid DB; never create a new DB.
        backup_db('pre-upgrade')
        db_id=secrets.token_hex(24)
        with sqlite3.connect(DB) as c:
            c.execute('CREATE TABLE IF NOT EXISTS rackflow_meta (key TEXT PRIMARY KEY,value TEXT NOT NULL)')
            c.execute('INSERT INTO rackflow_meta(key,value) VALUES(?,?)',('volume_id',db_id))
            c.commit()
        GUARD.write_text(db_id)
    elif not db_id or not marker_id or db_id!=marker_id:
        raise RuntimeError('STORAGE SAFETY STOP: volume identity missing/mismatch. Do not recreate users or stock.')
    if pinned and pinned!=db_id:
        raise RuntimeError("STORAGE SAFETY STOP: different Railway volume detected. Attach original volume; no inventory modified.")
    EXPECTED_ID=db_id
    GUARD_ACTIVE=True
    # Additive migration: preserve existing rack stock, movements and employee logins.
    with LOCK,connect() as db:
        db.execute("CREATE TABLE IF NOT EXISTS varieties (rack TEXT NOT NULL, model TEXT NOT NULL, variety TEXT NOT NULL DEFAULT '', PRIMARY KEY(rack,model))")
        db.execute('CREATE TABLE IF NOT EXISTS shop_orders (id INTEGER PRIMARY KEY AUTOINCREMENT, shop TEXT NOT NULL, created_by TEXT NOT NULL, created_at TEXT NOT NULL)')
        db.execute('CREATE TABLE IF NOT EXISTS shop_order_items (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL, variety TEXT NOT NULL, model TEXT NOT NULL, rack TEXT NOT NULL, qty INTEGER NOT NULL, picked INTEGER NOT NULL DEFAULT 0, FOREIGN KEY(order_id) REFERENCES shop_orders(id))')
        # Additive order-cancellation fields; retain old orders and all inventory.
        order_columns={r[1] for r in db.execute('PRAGMA table_info(shop_orders)')}
        for column,definition in [('status',"TEXT NOT NULL DEFAULT 'active'"),('cancelled_by','TEXT'),('cancelled_at','TEXT')]:
            if column not in order_columns:db.execute(f'ALTER TABLE shop_orders ADD COLUMN {column} {definition}')
        db.execute('CREATE TABLE IF NOT EXISTS stock_imports (export_id TEXT PRIMARY KEY, imported_at TEXT NOT NULL)')
        db.execute('CREATE TABLE IF NOT EXISTS stock_add_uploads (file_sha256 TEXT PRIMARY KEY, uploaded_at TEXT NOT NULL, uploaded_by TEXT NOT NULL)')
        db.commit()
    backup_db('latest')
    print('RackFlow persistent SQLite verified:',DB,'volume identity',EXPECTED_ID[:8],flush=True)

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
    backup_after_change()
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
    backup_after_change()
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
        varieties={(v['rack'],v['model']):v['variety'] for v in db.execute('SELECT rack,model,variety FROM varieties')}
    by_rack={r['rack']:{'rack':r['rack'],'capacity':r['capacity'],'used':0,'models':[]} for r in rack_rows}
    for a in all_stock:
        if a['rack'] in by_rack:
            by_rack[a['rack']]['used']+=a['qty']
            by_rack[a['rack']]['models'].append({'model':a['model'],'qty':a['qty'],'variety':varieties.get((a['rack'],a['model']),'')})
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
        backup_after_change()
        return {'ok':True,'rack':rack,'model':model,'before':before,'after':after}

def make_excel(template=False):
    # A snapshot travels with each downloaded workbook so live employee movements are
    # preserved when the admin later uploads their edited copy.
    with LOCK,connect() as db:
        baseline=[(r['rack'],r['model'],r['qty']) for r in db.execute('SELECT rack,model,qty FROM stock')]
        snapshot_racks=[(r['rack'],r['capacity']) for r in db.execute('SELECT rack,capacity FROM racks ORDER BY rack COLLATE NOCASE')]
        snapshot_varieties={(r['rack'],r['model']):r['variety'] for r in db.execute('SELECT rack,model,variety FROM varieties')}
        export_id=uuid.uuid4().hex
    wb=Workbook();ws=wb.active;ws.title='Stock Master' 
    ws.append(['Rack Number','Model','Quantity','Capacity','Variety'])
    if template:
        for row in [('A-001','iPhone 13 D-15',15,20,'D-15'),('A-002','iPhone 14 CBK-1',0,20,'CBK-1'),('B-001','Samsung A55 C-1',10,20,'C-1')]:ws.append(row)
    else:
        for rack,capacity in snapshot_racks:
            models=[(model,qty) for rr,model,qty in baseline if rr==rack]
            if models:
                for model,qty in models:ws.append([rack,model,qty,capacity,snapshot_varieties.get((rack,model),'')])
            else:ws.append([rack,'',0,capacity,''])
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    for cell in ws[1]:cell.font=Font(color='FFFFFF',bold=True);cell.fill=PatternFill('solid',fgColor='17364B');cell.alignment=Alignment(horizontal='center')
    for col,width in enumerate([22,43,18,18,22],1):ws.column_dimensions[get_column_letter(col)].width=width
    info=wb.create_sheet('Instructions')
    for x in [
        ['RACK STOCK MASTER — INSTRUCTIONS'],
        ['1. Each row assigns one phone-cover model to a rack.'],
        ['2. Edit Rack Number, Model, Quantity, Capacity and Variety in the .xlsx file.'],
        ['2a. Variety is used for the Shop Orders variety-to-model dropdown.'],
        ['3. Repeat rack numbers for multiple models in one rack (capacity must be identical).'],
        ['4. To reserve an empty rack, enter rack number, blank model, quantity 0 and capacity.'],
        ['5. Upload REPLACES current stock numbers; download latest first to keep employee changes.'],
        ['6. Employee IN and OUT changes are logged separately.'],
        ['7. Racks must have stock <= capacity; quantities cannot be negative.'],
    ]: info.append(x)
    info.column_dimensions['A'].width=110
    info['A10']='SAFE IMPORT: Do not delete the hidden _RackFlow_Snapshot sheet. Keep it with the Stock Master.'
    info['A11']='Live employee IN/OUT changes after download are automatically retained when you import the edited Excel.'
    info['A12']='To set an absolute final count, download a fresh stock file and edit it again.'
    info['A13']='An edited workbook can only be imported once. Download a fresh copy for the next update.'
    snapshot=wb.create_sheet('_RackFlow_Snapshot')
    snapshot.append(['RACKFLOW_SAFE_IMPORT_V1',export_id,EXPECTED_ID or ''])
    snapshot.append(['Rack Number','Model','Original Quantity'])
    for row in baseline:snapshot.append(row)
    snapshot.sheet_state='veryHidden'
    buf=BytesIO();wb.save(buf);return buf.getvalue()

def make_add_only_template():
    """Independent add-only workbook; never exports or snapshots existing stock."""
    wb=Workbook();ws=wb.active;ws.title='Add Rack Stock'
    ws.append(['Rack Number','Model','Quantity','Capacity','Variety'])
    ws.freeze_panes='A2';ws.auto_filter.ref='A1:E1'
    for cell in ws[1]:
        cell.font=Font(color='FFFFFF',bold=True)
        cell.fill=PatternFill('solid',fgColor='17364B')
        cell.alignment=Alignment(horizontal='center')
    for col,width in enumerate((24,44,18,18,22),1):ws.column_dimensions[get_column_letter(col)].width=width
    info=wb.create_sheet('Instructions')
    for line in (
        'ADD RACK & STOCK — INCREMENTAL IMPORT',
        'Fill only the NEW inventory being added. Do not include existing stock quantities.',
        'Rack Number: enter existing rack or new rack name (e.g. Rack B1, Rack 11).',
        'Model: product model and colour (e.g. Vivo V30 Pro - Black).',
        'Quantity: positive number of physical pieces being ADDED, not final stock count.',
        'Capacity: physical total rack capacity, NOT space remaining. For existing rack, leave blank to keep its current capacity.',
        'Variety: e.g. D10; blank only for existing model with a variety already assigned.',
        'For a NEW rack, provide capacity on every row for that rack.',
        'The app ADDS quantities to live stock. Employee take-outs are never overwritten.',
        'Uploading the identical file a second time is blocked. Use a new template for a new batch.',
        'No existing racks or models are deleted. Admin only.',
    ):info.append([line])
    info.column_dimensions['A'].width=115
    buf=BytesIO();wb.save(buf);return buf.getvalue()

def add_only_stock_excel(raw, username):
    """Atomic, admin-only additive import; validates full batch before any changes."""
    if len(raw)>12*1024*1024:raise ValueError('Excel file must be under 12 MB')
    file_hash=hashlib.sha256(raw).hexdigest()
    try:wb=load_workbook(BytesIO(raw),read_only=True,data_only=True)
    except Exception:raise ValueError('Upload an .xlsx Add Rack & Stock template')
    try:
        if 'Add Rack Stock' not in wb.sheetnames:
            raise ValueError('Use the Add Rack & Stock template, not the Current Stock Excel')
        ws=wb['Add Rack Stock']
        header=[re.sub(r'[^a-z0-9]','',str(c.value or '').lower()) for c in next(ws.iter_rows(min_row=1,max_row=1))]
        expected=['racknumber','model','quantity','capacity','variety']
        if header[:5]!=expected:raise ValueError('Template columns must be Rack Number, Model, Quantity, Capacity, Variety')
        rows=[]
        for i,values in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
            if all(v is None or not str(v).strip() for v in values[:5]):continue
            if i>20002:raise ValueError('Too many rows (maximum 20,000)')
            rack=cleanstr(values[0]).upper();model=cleanstr(values[1]);variety=cleanstr(values[4] if len(values)>4 else '')
            if not rack or len(rack)>80 or not model or len(model)>200:raise ValueError(f'Row {i}: rack and model are required')
            try:
                val=values[2]
                if isinstance(val,bool) or not str(val).strip() or int(val)!=float(val):raise ValueError()
                qty=int(val)
            except (ValueError,TypeError,OverflowError):raise ValueError(f'Row {i}: quantity must be a whole number')
            if qty<=0:raise ValueError(f'Row {i}: quantity must be greater than zero')
            cap=None
            if len(values)>3 and values[3] is not None and str(values[3]).strip():
                try:
                    val=values[3]
                    if isinstance(val,bool) or int(val)!=float(val):raise ValueError()
                    cap=int(val)
                except (ValueError,TypeError,OverflowError):raise ValueError(f'Row {i}: capacity must be a whole number')
                if cap<=0:raise ValueError(f'Row {i}: capacity must be greater than zero')
            rows.append((rack,model,qty,cap,variety,i))
        if not rows:raise ValueError('No stock rows found. Enter at least one model and quantity')
        # Same rack/model can appear multiple times within a file; sum amounts.
        caps={};adds={};requested_variety={};display={}
        for rack,model,qty,cap,variety,i in rows:
            key=(rack,model.casefold())
            adds[key]=adds.get(key,0)+qty;display.setdefault(key,model)
            if cap is not None:
                if rack in caps and caps[rack]!=cap:raise ValueError(f'Row {i}: conflicting capacities for {rack}')
                caps[rack]=cap
            if variety:
                if key in requested_variety and requested_variety[key].casefold()!=variety.casefold():raise ValueError(f'Row {i}: conflicting varieties for {model}')
                requested_variety[key]=variety
    finally:wb.close()
    with LOCK,connect() as db:
        db.execute('BEGIN IMMEDIATE')
        if db.execute('SELECT 1 FROM stock_add_uploads WHERE file_sha256=?',(file_hash,)).fetchone():
            raise ValueError('This exact stock-addition file was already uploaded; no stock was changed')
        existing_caps={r['rack']:r['capacity'] for r in db.execute('SELECT rack,capacity FROM racks')}
        existing_stocks={(r['rack'],r['model'].casefold()):(r['model'],r['qty']) for r in db.execute('SELECT rack,model,qty FROM stock')}
        old_variety={(r['rack'],r['model'].casefold()):r['variety'] for r in db.execute('SELECT rack,model,variety FROM varieties')}
        used={r['rack']:r['used'] for r in db.execute('SELECT rack,COALESCE(SUM(qty),0) used FROM stock GROUP BY rack')}
        added_per_rack={}
        for (rack,_),qty in adds.items():added_per_rack[rack]=added_per_rack.get(rack,0)+qty
        new_caps={}
        for rack,additional in added_per_rack.items():
            if rack not in existing_caps and rack not in caps:raise ValueError(f'New rack {rack} needs a capacity in the Excel file')
            final_capacity=caps.get(rack,existing_caps.get(rack))
            if used.get(rack,0)+additional>final_capacity:
                raise ValueError(f'Rack {rack}: after adding {additional}, stock would be {used.get(rack,0)+additional}, over capacity {final_capacity}')
            new_caps[rack]=final_capacity
        for key in adds:
            if key not in existing_stocks and not requested_variety.get(key):
                raise ValueError(f'New model {display[key]} in {key[0]} requires Variety (e.g. D10)')
            if key in old_variety and requested_variety.get(key) and old_variety[key] and old_variety[key].casefold()!=requested_variety[key].casefold():
                raise ValueError(f'Variety mismatch for {key[0]} / {display[key]}; stock not changed')
        # Make recovery snapshot before transaction; database is unchanged at this point.
        backup_db('pre-import')
        for rack,cap in new_caps.items():
            db.execute('INSERT INTO racks(rack,capacity) VALUES(?,?) ON CONFLICT(rack) DO UPDATE SET capacity=excluded.capacity',(rack,cap))
        for key,amount in adds.items():
            rack,low=key; existing=existing_stocks.get(key);model=existing[0] if existing else display[key]
            before=existing[1] if existing else 0
            if existing:
                db.execute('UPDATE stock SET qty=? WHERE rack=? AND model=?',(before+amount,rack,model))
            else:
                db.execute('INSERT INTO stock(rack,model,qty) VALUES(?,?,?)',(rack,model,amount))
            variety=old_variety.get(key) or requested_variety.get(key,'')
            db.execute('INSERT INTO varieties(rack,model,variety) VALUES(?,?,?) ON CONFLICT(rack,model) DO UPDATE SET variety=excluded.variety',(rack,model,variety))
            db.execute('INSERT INTO movements(at,employee,action,rack,model,qty,before_qty,after_qty,note) VALUES(?,?,?,?,?,?,?,?,?)',(stamp(),username,'IN',rack,model,amount,before,before+amount,'Add Rack & Stock Excel'))
        db.execute('INSERT INTO stock_add_uploads(file_sha256,uploaded_at,uploaded_by) VALUES(?,?,?)',(file_hash,stamp(),username))
        db.commit()
    backup_after_change()
    return {'ok':True,'racks':len(new_caps),'new_racks':sum(r not in existing_caps for r in new_caps),'items':len(adds),'added_units':sum(adds.values())}

def import_excel(raw):
    if len(raw)>12*1024*1024:raise ValueError('Excel file must be under 12 MB')
    try: wb=load_workbook(BytesIO(raw),read_only=True,data_only=True)
    except Exception:raise ValueError('Please upload a valid .xlsx file')
    ws=wb['Stock Master'] if 'Stock Master' in wb.sheetnames else wb.active
    if '_RackFlow_Snapshot' not in wb.sheetnames:
        raise ValueError('Unsafe older Excel file: download a new Current Stock Excel from this updated RackFlow app, copy your edits into it, then upload. This protects employee take-outs.')
    snap=wb['_RackFlow_Snapshot']
    marker,export_id,volume_id=[snap.cell(1,c).value for c in (1,2,3)]
    if marker!='RACKFLOW_SAFE_IMPORT_V1' or not isinstance(export_id,str) or not re.fullmatch('[a-f0-9]{32}',export_id):
        raise ValueError('Invalid Excel stock snapshot. Download a fresh Current Stock file.')
    if volume_id!=EXPECTED_ID:
        raise ValueError('This stock Excel belongs to a different RackFlow database. No stock changed.')
    baseline={}
    for rr,mm,qq in snap.iter_rows(min_row=3,max_col=3,values_only=True):
        if rr is None or mm is None:continue
        key=(cleanstr(rr).upper(),cleanstr(mm).lower())
        if key in baseline:raise ValueError('Duplicate baseline row in Excel snapshot')
        if not isinstance(qq,int) or qq<0:raise ValueError('Invalid baseline quantity in Excel snapshot')
        baseline[key]=qq
    rows=ws.iter_rows(values_only=True)
    try: headers=[re.sub(r'[^a-z0-9]','',str(c or '').lower()) for c in next(rows)]
    except StopIteration: raise ValueError('Excel has no rows')
    def index(options):
        for v in options:
            if v in headers:return headers.index(v)
        return None
    rix=index(('racknumber','rack','rackid','location'));mix=index(('model','modelname','product','item'));qix=index(('quantity','qty','stock','stockquantity'));cix=index(('capacity','rackcapacity','maxqty'));vix=index(('variety','varietyname','covertype','type','producttype'))
    if None in (rix,mix,qix): raise ValueError('Need columns Rack Number, Model, Quantity. Capacity is optional.')
    racks={};records={};issues=[]
    for i,row in enumerate(rows,2):
        if not any(v is not None and str(v).strip() for v in row):continue
        def at(j):return row[j] if j is not None and j<len(row) else None
        rack=cleanstr(at(rix)).upper();model=cleanstr(at(mix));variety=cleanstr(at(vix))
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
            key=(rack,model.lower()); rec=records.setdefault(key,{'rack':rack,'model':model,'qty':0,'variety':variety});rec['qty']+=qty
            if rec['variety']!=variety:issues.append(f'Row {i}: conflicting variety for {rack}/{model}')
    usage={r:0 for r in racks}
    for rec in records.values():usage[rec['rack']]+=rec['qty']
    for r,used in usage.items():
        if used>racks[r]:issues.append(f'Rack {r}: {used} units exceeds capacity {racks[r]}')
    if issues:raise ValueError('Excel validation failed: '+'; '.join(issues[:12]))
    if not racks:raise ValueError('No racks found in Excel')
    with LOCK,connect() as db:
        db.execute('BEGIN IMMEDIATE')
        if db.execute('SELECT 1 FROM stock_imports WHERE export_id=?',(export_id,)).fetchone():
            raise ValueError('This Excel has already been uploaded. Download a fresh Current Stock Excel to make another update.')
        current={(r['rack'],r['model'].lower()):(r['model'],r['qty']) for r in db.execute('SELECT rack,model,qty FROM stock')}
        existing_varieties={(r['rack'],r['model'].lower()):r['variety'] for r in db.execute('SELECT rack,model,variety FROM varieties')}
        proposed={}
        target_keys=set(records)|set(baseline)
        for key in target_keys:
            initial=baseline.get(key,0)
            new_excel=records[key]['qty'] if key in records else 0
            existing=current.get(key,(key[1],0))[1]
            final=existing+(new_excel-initial)
            if final<0:
                raise ValueError(f'Upload stopped: {key[0]} / {key[1]} has only {existing} live pieces after employee take-outs. Your Excel change would make stock negative. Download a fresh file.')
            proposed[key]=(records[key]['model'] if key in records else current.get(key,(key[1],0))[0],final)
        for key,(model,qty) in current.items():
            if key not in proposed:proposed[key]=(model,qty)
        totals={rack:0 for rack in racks}
        for (rack,_),(model,qty) in proposed.items():
            if rack not in racks:
                if qty>0:raise ValueError(f'Upload stopped: rack {rack} still has live stock; add it to Excel or move its stock first.')
                continue
            totals[rack]+=qty
        for rack,used in totals.items():
            if used>racks[rack]:
                raise ValueError(f'Upload stopped: {rack} has {used} live pieces after employee movements, exceeding Excel capacity {racks[rack]}. Increase its capacity or download fresh stock.')
        # Prepare a restorable snapshot before applying the checked changes.
        db.rollback()
        backup_db('pre-import')
        db.execute('BEGIN IMMEDIATE')
        if db.execute('SELECT 1 FROM stock_imports WHERE export_id=?',(export_id,)).fetchone():
            raise ValueError('This file was already imported.')
        # LOCK prevents employee IN/OUT during validation and commit.
        db.execute('DELETE FROM stock');db.execute('DELETE FROM racks');db.execute('DELETE FROM varieties')
        db.executemany('INSERT INTO racks(rack,capacity) VALUES(?,?)',racks.items())
        db.executemany('INSERT INTO stock(rack,model,qty) VALUES(?,?,?)',[(rack,model,qty) for (rack,_),(model,qty) in proposed.items() if rack in racks])
        db.executemany('INSERT INTO varieties(rack,model,variety) VALUES(?,?,?)',[(rack,model,records.get((rack,low),{}).get('variety',existing_varieties.get((rack,low),''))) for (rack,low),(model,qty) in proposed.items() if rack in racks])
        db.execute('INSERT INTO stock_imports(export_id,imported_at) VALUES(?,?)',(export_id,stamp()))
        db.execute('INSERT INTO movements(at,employee,action,rack,model,qty,before_qty,after_qty,note) VALUES(?,?,?,?,?,?,?,?,?)',(stamp(),'Admin','EXCEL SAFE IMPORT','ALL','Admin Excel changes merged with live employee IN/OUT',sum(totals.values()),None,None,f'{len(racks)} racks; {len(proposed)} model-rack records; export {export_id[:8]}'))
        db.commit()
        backup_after_change()
    return {'ok':True,'racks':len(racks),'items':len(proposed),'units':sum(totals.values()),'safe_merge':True}

def shop_catalog():
    with connect() as db:
        rows=db.execute("SELECT s.rack,s.model,s.qty,COALESCE(v.variety,'') AS variety FROM stock s LEFT JOIN varieties v ON v.rack=s.rack AND v.model=s.model WHERE s.model<>'' ORDER BY v.variety,s.model,s.rack").fetchall()
        return [dict(r) for r in rows]

def list_shop_orders():
    with connect() as db:
        orders=[dict(r) for r in db.execute('SELECT * FROM shop_orders ORDER BY id DESC LIMIT 150')]
        for o in orders:
            o['items']=[dict(r) for r in db.execute('SELECT id,variety,model,rack,qty,picked FROM shop_order_items WHERE order_id=? ORDER BY rack,model',(o['id'],))]
    return orders

def create_shop_order(data,username):
    shop=cleanstr(data.get('shop'))
    lines=data.get('items')
    if not shop or len(shop)>100:raise ValueError('Enter a valid shop name')
    if not isinstance(lines,list) or not 1<=len(lines)<=150:raise ValueError('Add 1 to 150 order items')
    prepared=[]
    with LOCK,connect() as db:
        db.execute('BEGIN IMMEDIATE')
        for line in lines:
            if not isinstance(line,dict):raise ValueError('Invalid order line')
            variety=cleanstr(line.get('variety'));model=cleanstr(line.get('model'))
            qty=positive(line.get('qty'))
            if qty>100000:raise ValueError('Order quantity too large')
            if not variety or not model:raise ValueError('Select variety and model')
            matches=db.execute('SELECT s.rack,s.qty FROM stock s INNER JOIN varieties v ON v.rack=s.rack AND v.model=s.model WHERE v.variety=? AND s.model=? AND s.qty>0 ORDER BY s.qty DESC,s.rack',(variety,model)).fetchall()
            if not matches:raise ValueError(f'No available rack for {variety} / {model}')
            remaining=qty
            for r in matches:
                if remaining<=0:break
                # Reserve nothing until picking; assignment uses current stock.
                amount=min(remaining,r['qty'])
                prepared.append((variety,model,r['rack'],amount))
                remaining-=amount
            if remaining:raise ValueError(f'Only {qty-remaining} pieces available for {variety} / {model}')
        cur=db.execute('INSERT INTO shop_orders(shop,created_by,created_at) VALUES(?,?,?)',(shop,username,stamp()))
        db.executemany('INSERT INTO shop_order_items(order_id,variety,model,rack,qty,picked) VALUES(?,?,?,?,?,0)',[(cur.lastrowid,*r) for r in prepared])
        db.commit()
    backup_after_change()
    return {'ok':True,'order_id':cur.lastrowid}

def cancel_shop_order(data,username,role):
    order_id=positive(data.get('order_id'),'Order ID')
    with LOCK,connect() as db:
        db.execute('BEGIN IMMEDIATE')
        order=db.execute('SELECT id,status,created_by FROM shop_orders WHERE id=?',(order_id,)).fetchone()
        if not order:raise ValueError('Order not found')
        if role!='admin' and order['created_by']!=username:raise PermissionError('You can cancel only orders you placed')
        if order['status']=='cancelled':raise ValueError('Order is already cancelled')
        picked=db.execute('SELECT COALESCE(SUM(picked),0) FROM shop_order_items WHERE order_id=?',(order_id,)).fetchone()[0]
        if picked:raise ValueError('Picking has already started. This order cannot be cancelled. Return items through the regular stock IN process, with a separate record.')
        db.execute("UPDATE shop_orders SET status='cancelled',cancelled_by=?,cancelled_at=? WHERE id=? AND status='active'",(username,stamp(),order_id))
        db.commit()
    # No stock adjustment: orders reserve nothing and stock is deducted only during Take Out.
    backup_after_change()
    return {'ok':True,'order_id':order_id,'status':'cancelled'}

def pick_shop_item(data,username):
    item_id=positive(data.get('item_id'),'Item ID');qty=positive(data.get('qty'),'Take Out quantity')
    with LOCK,connect() as db:
        db.execute('BEGIN IMMEDIATE')
        item=db.execute('SELECT * FROM shop_order_items WHERE id=?',(item_id,)).fetchone()
        if item is None:raise ValueError('Order item not found')
        order_status=db.execute('SELECT status FROM shop_orders WHERE id=?',(item['order_id'],)).fetchone()
        if not order_status or order_status['status']!='active':raise ValueError('Cancelled orders cannot be picked')
        if qty>item['qty']-item['picked']:raise ValueError('Quantity exceeds remaining order quantity')
        stock=db.execute('SELECT qty FROM stock WHERE rack=? AND model=?',(item['rack'],item['model'])).fetchone()
        if not stock or stock['qty']<qty:raise ValueError('Not enough stock in assigned rack. Check physical stock and update inventory.')
        before=stock['qty'];after=before-qty
        db.execute('UPDATE stock SET qty=? WHERE rack=? AND model=?',(after,item['rack'],item['model']))
        db.execute('UPDATE shop_order_items SET picked=picked+? WHERE id=?',(qty,item_id))
        order=db.execute('SELECT shop FROM shop_orders WHERE id=?',(item['order_id'],)).fetchone()
        db.execute('INSERT INTO movements(at,employee,action,rack,model,qty,before_qty,after_qty,note) VALUES(?,?,?,?,?,?,?,?,?)',(stamp(),username,'OUT',item['rack'],item['model'],qty,before,after,f'Shop Order #{item["order_id"]} — {order["shop"]}'))
        db.commit()
    backup_after_change()
    return {'ok':True,'remaining':item['qty']-item['picked']-qty,'stock_after':after}

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
        if path=='/api/shop-catalog':respond_json(self,shop_catalog());return
        if path=='/api/shop-orders':respond_json(self,list_shop_orders());return
        if path=='/api/users':
            if identity['role']!='admin':return error(self,'Admin only',403)
            respond_json(self,admin_users());return
        if path=='/api/backup':
            if identity['role']!='admin':return error(self,'Admin only',403)
            with LOCK:
                snapshot=backup_db('manual')
                payload=snapshot.read_bytes()
            self.binary(payload,'application/octet-stream',attachment='rackflow-full-backup.db');return
        if path=='/api/storage-status':
            if identity['role']!='admin':return error(self,'Admin only',403)
            with connect() as db:
                counts={t:db.execute(f'SELECT COUNT(*) FROM {t}').fetchone()[0] for t in ('racks','stock','users','movements')}
            respond_json(self,{'persistent':True,'db_path':str(DB),'backup_files':len(list(BACKUPS.glob('*.db'))),'volume_id':EXPECTED_ID,'external_guard_set':bool(os.getenv('RACKFLOW_EXPECTED_VOLUME_ID')),'counts':counts});return
        if path=='/api/add-stock-template':
            if identity['role']!='admin':return error(self,'Admin only',403)
            self.binary(make_add_only_template(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',attachment='rackflow_add_rack_stock_template.xlsx');return
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
        self.send_header('Cache-Control','no-store');self.send_header('X-Content-Type-Options','nosniff');self.end_headers();self.wfile.write(payload)
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
            if path=='/api/add-stock-import':
                if identity['role']!='admin':return error(self,'Admin only',403)
                respond_json(self,add_only_stock_excel(raw,identity['username']));return
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
            if path=='/api/shop-orders':respond_json(self,create_shop_order(data,identity['username']));return
            if path=='/api/shop-pick':respond_json(self,pick_shop_item(data,identity['username']));return
            if path=='/api/shop-cancel':
                try:respond_json(self,cancel_shop_order(data,identity['username'],identity['role']))
                except PermissionError as ex:return error(self,str(ex),403)
                return
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
