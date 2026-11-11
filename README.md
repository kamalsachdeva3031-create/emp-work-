# RackFlow — Excel-controlled warehouse rack scanner

## Start
1. Install Python 3.10+ on your Windows hosting PC (enable **Add Python to PATH**).
2. Extract this ZIP and double-click `START_WINDOWS.bat`.
3. On PC open `http://localhost:8765`. On phones the same Wi-Fi network can open `http://PC-IP:8765` for manual/USB barcode scanning. **For phone CAMERA scanning, you must access the app over HTTPS** because phone browsers generally block camera access on ordinary `http://PC-IP` websites.
4. On the home page select **Admin / Excel**, upload the included `Rack_Master.xlsx` filled with your stock, and print QR stickers to attach to racks.
5. An employee enters their name, scans rack sticker, chooses IN or OUT, enters model and quantity, then saves.

## Excel master

`Rack_Master.xlsx` has `Rack Number | Model | Quantity | Rack Capacity | Capacity Left | Rack Status | QR Code Link`. Edit only A-D. Columns E-G are formula-driven. One rack may have multiple models; repeat rack number and same rack capacity. QR Code Link creates a link that displays a printable barcode via quickchart.io (internet needed); the app itself generates the same QR codes offline via **Admin / Excel > Print QR Labels**. A QR sticker always encodes only the rack ID.

The app imports columns `Rack Number`, `Model`, `Quantity`, `Rack Capacity` and derives remaining capacity. Upload **replaces** master stock. Download current Excel immediately before editing to avoid overwriting IN/OUT changes. There is no live file link to Excel on your computer; uploads and downloads are the synchronization method.

**IN:** can enter an existing model or a new model; app ensures total rack stock doesn't exceed capacity. **OUT:** shows models in scanned rack and validates sufficient stock. Changes are stored in SQLite (`inventory.db`) with employee and timestamp.

## CAMERA setup — IMPORTANT

The phone scanner uses its camera via `getUserMedia`, then scans QR labels via native `BarcodeDetector` when supported, with `jsQR` fallback. The fallback loads from CDN when the library is not locally available. **A camera opened over ordinary HTTP on another LAN computer does not meet browsers' secure-origin requirement.** To make scanning work on employees' phones, deploy the app on a trusted HTTPS URL using a secure reverse proxy/tunnel (or host with HTTPS and a valid trusted certificate). Camera permission must be granted, rear camera unobstructed, and no other app occupying the camera. Check on Android Chrome or iOS Safari. There is a manual rack field as a backup. Camera behavior must be tested on the actual device and HTTPS URL; it cannot be guaranteed solely by code.

This is a local admin-without-login starter app. Do not expose the app directly to the public Internet without adding user authentication and access controls; an HTTPS tunnel alone is **not** sufficient security. Keep backups of `inventory.db`. Restrict who can import Excel because import replaces all existing stock.
