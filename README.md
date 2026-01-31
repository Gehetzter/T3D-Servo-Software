# T3D Servo Software

Overview
--------
T3D Servo Software is a desktop application written in Python that provides a graphical interface to communicate with T3D servo drives over an RS-485 (Modbus-like) link. The application is intended for parameter inspection, parameter editing, and reading runtime status information from multiple drives.

HLTNC provides ServoTuning_V1.00_English.exe for this purpose, but you can only connect to one drive at a time and baud rate settings are limited.

Key features
------------
- Configure serial connection: choose COM port, baudrate and parity.
- Add and manage multiple drives (one tab per drive).
- Load parameter definitions from `config/parameter_SpindleHS1.xml` and present them as editable rows.
- Read/write individual parameters using function codes (reads: `0x03`, writes: `0x06`).
- Read status registers via function code `0x04` (address range `0x0000..0x0063`) or `0x03` (address range `0x1000..0x1063`) with dedicated Status tabs.
- Enable/disable a drive by writing to parameter address `0x62` using function `0x06`.
- Save parameters to EEPROM (write to `0x1001` and wait the required time before allowing further operations).

Repository layout
-----------------
- `src/` — application source code
  - `src/main.py` — application entry point
  - `src/gui.py` — Tkinter GUI and UI logic
  - `src/transport.py` — serial transport layer, CRC, send/receive handling
- `config/parameter_SpindleHS1.xml` — parameter definitions used by the UI
- `config/status_04.xml` — status list for function `0x04` (addresses `0x0000..0x002D`)
- `config/status_03.xml` — status list for function `0x03` (addresses `0x1000..0x1027`)
- `config/gui_settings.json` — GUI settings saved at runtime (per-user; excluded from source control)

Requirements
------------
- Python 3.8 or newer
- `pyserial` (see `requirements.txt`)

Quick setup
-----------
1. Create and activate a virtual environment (PowerShell example):

```powershell
python -m venv .venv
& .\.venv\Scripts\Activate.ps1
```

2. Install Python dependencies:

```powershell
pip install -r requirements.txt
```

Run the application
-------------------

```powershell
python src/main.py
```

Usage notes
-----------
- Connect: select a COM port, choose the baudrate and parity, then click `Connect`.
- Add a drive: enter the drive ID (address) and click `Add Drive`. A tab for the drive will appear.
- Parameters: each parameter row shows `Description`, `Min`, `Value`, `Max` and has `Read`/`Write` buttons. Use `Read All` to refresh all values; read errors are aggregated into a single dialog to avoid many popups.
- Status tabs: each drive has two status tabs (`Status 04` and `Status 03`) showing `Addr (hex) | Description | Value | Units`. Use the `Refresh` buttons to poll the device; reads are executed in chunks up to 8 registers.
- Enable/Disable: toggles the drive enabled state by writing `1`/`0` to parameter address `0x62` via function `0x06`.
- EEPROM Save: writing the EEPROM-save value writes to address `0x1001` and the UI waits (several seconds) while the drive completes the EEPROM write.

Protocol and CRC
----------------
- The transport layer implements Modbus-RTU style framing. The CRC-16 (Modbus) is computed in RTU mode and appended to frames as `CRC Low` then `CRC High`.
- `src/transport.py` contains the CRC implementation and the logic to send a request then wait for the response while respecting half-duplex timing.

Configuration persistence
------------------------
- The app stores per-user settings (selected COM port, baudrate, parity and saved drive list) in `config/gui_settings.json`. This file is excluded from source control by `.gitignore` to avoid committing local settings.

Troubleshooting
---------------
- If you see CRC errors, verify the raw bytes being sent and received (enable transport debug in `src/main.py` to log raw frames).
- Check wiring and RS-485 adapter polarity (A/B) and that any required DE/RE direction control is handled by the adapter (the app assumes the adapter handles DE/RE automatically).
- For timing issues, verify the baudrate and that the device accepts the configured timeout. The transport computes a start-wait timeout based on the baudrate and a configurable character count.

Development notes
-----------------
- GUI: `tkinter` + `ttk` (native widgets). The app populates parameter rows from `config/parameter_SpindleHS1.xml` at startup.
- Status definitions: `config/status_04.xml` and `config/status_03.xml` define which addresses and descriptions appear in the Status tabs.
- Tests: basic import/compile smoke tests are used. Hardware tests must be done against the real servo drives.

Contributing
------------
Contributions are welcome. Please open issues or pull requests.

License
-------
MIT License

Copyright (c) 2026 gehetzter

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.