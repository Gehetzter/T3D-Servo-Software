import tkinter as tk
from tkinter import ttk, messagebox
import threading
import struct
import time
import tkinter.font as tkfont
import xml.etree.ElementTree as ET
import serial.tools.list_ports
from functools import partial
import json
import os

from transport import SerialTransport, compute_crc

# GUI-related constants
# Parameter address used to enable/disable the drive
ENABLE_PARAM_ADDR = 0x62
# EEPROM save address and value
EEPROM_SAVE_ADDR = 0x1001
EEPROM_SAVE_VALUE = 0x1234
# Wait (seconds) after sending EEPROM save command; recommended >5s
EEPROM_WAIT_SECONDS = 5.5


def load_parameters(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    params = []
    for node in root.findall('ServoParameterTable'):
        try:
            pid = int(node.find('id').text)
        except Exception:
            continue
        params.append({
            'id': pid,
            'name': node.findtext('name',''),
            'description': node.findtext('description',''),
            'value': node.findtext('value','0'),
            'min': node.findtext('valueMin',''),
            'max': node.findtext('valueMax',''),
            'default': node.findtext('defaultValue',''),
            'type': node.findtext('type',''),
            'access': node.findtext('accessType',''),
        })
    return params


def load_status(xml_path):
    """Parse `config/status.xml` and return a list of status dicts.

    Each dict contains: id (int), name, description, value, type, units.
    """
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except Exception:
        return []
    stats = []
    for node in root.findall('ServoStatusTable'):
        try:
            sid = int(node.findtext('id','0'))
        except Exception:
            continue
        stats.append({
            'id': sid,
            'name': node.findtext('name',''),
            'description': node.findtext('description',''),
            'value': node.findtext('value','0'),
            'type': node.findtext('type',''),
            'units': node.findtext('units',''),
        })
    # sort by id
    stats.sort(key=lambda x: x['id'])
    return stats


class DriveTab:
    def __init__(self, parent, drive_id, transport: SerialTransport, params, tk_parent, close_callback=None, status_entries_04=None, status_entries_03=None, desc_width_chars: int = 40):
        self.drive_id = drive_id
        self.transport = transport
        self.params = params
        self.status_entries_04 = status_entries_04 or []
        self.status_entries_03 = status_entries_03 or []
        self.desc_width_chars = desc_width_chars
        self.frame = ttk.Frame(parent)
        self.tk_parent = tk_parent
        self.enabled = False
        self._close_cb = close_callback
        # Diagnostic: print how many parameters were passed to this drive tab
        try:
            print(f"Drive {self.drive_id}: {len(self.params)} parameters")
        except Exception:
            pass
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self.frame)
        top.pack(fill='x', padx=4, pady=4)
        self.enable_btn = ttk.Button(top, text='Enable', command=self.toggle_enable)
        self.enable_btn.pack(side='left')
        self.read_all_btn = ttk.Button(top, text='Read All', command=self.read_all)
        self.read_all_btn.pack(side='left', padx=6)
        self.save_eeprom_btn = ttk.Button(top, text='Save EEPROM', command=self.save_eeprom)
        self.save_eeprom_btn.pack(side='left', padx=6)
        # Close tab button
        if self._close_cb:
            self.close_btn = ttk.Button(top, text='Close', command=self._close_cb)
            self.close_btn.pack(side='right')

        # Create a sub-notebook with parameter groups (0xx, 1xx, 2xx) and Status
        container = ttk.Frame(self.frame)
        container.pack(fill='both', expand=True)

        param_notebook = ttk.Notebook(container)
        param_notebook.pack(fill='both', expand=True, padx=4, pady=4)

        # Helper to create a scrollable frame for each parameter group
        def make_param_page(title):
            page = ttk.Frame(param_notebook)
            canvas = tk.Canvas(page)
            canvas.pack(side='left', fill='both', expand=True)
            vsb = ttk.Scrollbar(page, orient='vertical', command=canvas.yview)
            vsb.pack(side='right', fill='y')
            canvas.configure(yscrollcommand=vsb.set)
            inner = ttk.Frame(canvas)
            canvas.create_window((0, 0), window=inner, anchor='nw')
            inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
            param_notebook.add(page, text=title)
            return inner

        page0 = make_param_page('Params 0xx')
        page1 = make_param_page('Params 1xx')
        page2 = make_param_page('Params 2xx')

        self.param_widgets = {}

        # column headers for each page
        def add_headers(parent):
            header = ttk.Frame(parent)
            header.pack(fill='x', padx=4, pady=(2,6))
            ttk.Label(header, text='Name', width=14).grid(row=0, column=0, sticky='w')
            ttk.Label(header, text='Description', width=self.desc_width_chars).grid(row=0, column=1, sticky='w')
            ttk.Label(header, text='Min', width=8).grid(row=0, column=2)
            ttk.Label(header, text='Value', width=10).grid(row=0, column=3)
            ttk.Label(header, text='Max', width=8).grid(row=0, column=4)
            ttk.Label(header, text='').grid(row=0, column=5)

        add_headers(page0)
        add_headers(page1)
        add_headers(page2)

        # Distribute parameters into pages based on id ranges
        for p in self.params:
            try:
                pid = int(p.get('id', 0))
            except Exception:
                pid = 0
            if 0 <= pid <= 99:
                parent = page0
            elif 100 <= pid <= 199:
                parent = page1
            else:
                parent = page2

            row = ttk.Frame(parent)
            row.pack(fill='x', padx=4, pady=2)
            name_lbl = ttk.Label(row, text=f"{p['name']} ({p['id']})", width=14, anchor='w', font=('Segoe UI', 9, 'bold'))
            name_lbl.grid(row=0, column=0, sticky='w')
            desc_lbl = ttk.Label(row, text=p.get('description',''), width=self.desc_width_chars, anchor='w')
            desc_lbl.grid(row=0, column=1, sticky='w', padx=(6,0))
            min_lbl = ttk.Label(row, text=str(p.get('min','')), width=8)
            min_lbl.grid(row=0, column=2)
            entry = ttk.Entry(row, width=10)
            entry.insert(0, str(p.get('value','')))
            entry.grid(row=0, column=3, padx=6)
            max_lbl = ttk.Label(row, text=str(p.get('max','')), width=8)
            max_lbl.grid(row=0, column=4)
            read_btn = ttk.Button(row, text='Read', command=partial(self.read_param, p, entry))
            read_btn.grid(row=0, column=5, padx=2)
            write_btn = ttk.Button(row, text='Write', command=partial(self.write_param, p, entry, entry))
            write_btn.grid(row=0, column=6, padx=2)
            self.param_widgets[p['id']] = (entry, read_btn, write_btn)

        # Status tab as separate tabs: Status 04 and Status 03
        status_page = ttk.Frame(param_notebook.master)
        param_notebook.add(status_page, text='Status 04')
        status_page_03 = ttk.Frame(param_notebook.master)
        param_notebook.add(status_page_03, text='Status 03')

        # --- Status 04 UI ---
        s_top = ttk.Frame(status_page)
        s_top.pack(fill='x')
        self.status_refresh_btn_04 = ttk.Button(s_top, text='Refresh Status 04', command=self.refresh_status_04)
        self.status_refresh_btn_04.pack(side='left')
        # Columns: addr(hex) | Description | Value | Units
        cols = ('addr', 'desc', 'value', 'units')
        self.status_tree_04 = ttk.Treeview(status_page, columns=cols, show='headings', height=10)
        self.status_tree_04.pack(fill='both', expand=True, padx=4, pady=4)
        self.status_tree_04.heading('addr', text='Addr')
        self.status_tree_04.heading('desc', text='Description')
        self.status_tree_04.heading('value', text='Value')
        self.status_tree_04.heading('units', text='Units')
        self.status_tree_04.column('addr', width=80, anchor='center')
        self.status_tree_04.column('desc', width=360)
        self.status_tree_04.column('value', width=120, anchor='e')
        self.status_tree_04.column('units', width=80)
        for s in self.status_entries_04:
            sid = int(s.get('id', 0))
            addr_text = format(sid, '#06x')
            desc = s.get('description','')
            val = s.get('value','')
            units = s.get('units','')
            self.status_tree_04.insert('', 'end', iid=str(sid), values=(addr_text, desc, val, units))
        # autosize columns to fit content
        try:
            f = tkfont.nametofont(self.status_tree_04.cget('font'))
        except Exception:
            f = None
        if f is not None:
            for col in ('addr', 'desc', 'value', 'units'):
                # measure header
                hdr_w = f.measure(self.status_tree_04.heading(col)['text'])
                max_w = hdr_w
                for iid in self.status_tree_04.get_children():
                    txt = str(self.status_tree_04.set(iid, col))
                    w = f.measure(txt)
                    if w > max_w:
                        max_w = w
                self.status_tree_04.column(col, width=max_w + 8)
        # make units consume remaining space; keep other columns minimal
        def autosize_04(event=None):
            try:
                total = self.status_tree_04.winfo_width()
                if total <= 20:
                    return
                cols = ('addr', 'desc', 'value', 'units')
                used = 0
                padding = 8 * len(cols)
                for c in cols[:-1]:
                    used += int(self.status_tree_04.column(c)['width'])
                rem = total - used - padding
                if rem < 60:
                    rem = 60
                self.status_tree_04.column('units', width=rem)
            except Exception:
                pass

        self.status_tree_04.bind('<Configure>', autosize_04)

        # --- Status 03 UI ---
        s_top3 = ttk.Frame(status_page_03)
        s_top3.pack(fill='x')
        self.status_refresh_btn_03 = ttk.Button(s_top3, text='Refresh Status 03', command=self.refresh_status_03)
        self.status_refresh_btn_03.pack(side='left')
        # Columns: addr(hex) | Description | Value | Units for 0x03 space
        cols3 = ('addr', 'desc', 'value', 'units')
        self.status_tree_03 = ttk.Treeview(status_page_03, columns=cols3, show='headings', height=10)
        self.status_tree_03.pack(fill='both', expand=True, padx=4, pady=4)
        self.status_tree_03.heading('addr', text='Addr')
        self.status_tree_03.heading('desc', text='Description')
        self.status_tree_03.heading('value', text='Value')
        self.status_tree_03.heading('units', text='Units')
        self.status_tree_03.column('addr', width=100, anchor='center')
        self.status_tree_03.column('desc', width=360)
        self.status_tree_03.column('value', width=120, anchor='e')
        self.status_tree_03.column('units', width=80)
        for s in self.status_entries_03:
            sid = int(s.get('id', 0))
            addr_text = format(sid, '#06x')
            desc = s.get('description','')
            val = s.get('value','')
            units = s.get('units','')
            self.status_tree_03.insert('', 'end', iid=str(sid), values=(addr_text, desc, val, units))
        try:
            f3 = tkfont.nametofont(self.status_tree_03.cget('font'))
        except Exception:
            f3 = None
        if f3 is not None:
            for col in ('addr', 'desc', 'value', 'units'):
                hdr_w = f3.measure(self.status_tree_03.heading(col)['text'])
                max_w = hdr_w
                for iid in self.status_tree_03.get_children():
                    txt = str(self.status_tree_03.set(iid, col))
                    w = f3.measure(txt)
                    if w > max_w:
                        max_w = w
                self.status_tree_03.column(col, width=max_w + 8)
        def autosize_03(event=None):
            try:
                total = self.status_tree_03.winfo_width()
                if total <= 20:
                    return
                cols = ('addr', 'desc', 'value', 'units')
                used = 0
                padding = 8 * len(cols)
                for c in cols[:-1]:
                    used += int(self.status_tree_03.column(c)['width'])
                rem = total - used - padding
                if rem < 60:
                    rem = 60
                self.status_tree_03.column('units', width=rem)
            except Exception:
                pass

        self.status_tree_03.bind('<Configure>', autosize_03)

    def toggle_enable(self):
        # Perform write to parameter address ENABLE_PARAM_ADDR using function 0x06
        target_addr = ENABLE_PARAM_ADDR
        new_val = 1 if not self.enabled else 0

        def worker():
            try:
                req = struct.pack('>B B H H', self.drive_id, 0x06, target_addr, new_val)
                crc = compute_crc(req)
                req += struct.pack('<H', crc)
                resp = self.transport.send_and_receive(req)
                # if success toggle state
                self.enabled = (new_val == 1)
                self.tk_parent.after(0, lambda: self.enable_btn.config(text='Disable' if self.enabled else 'Enable'))
            except Exception as e:
                self.tk_parent.after(0, (lambda exc=e: messagebox.showerror('Enable error', str(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def read_param(self, p, label):
        def worker():
            try:
                addr = int(p['id'])
                req = struct.pack('>B B H H', self.drive_id, 0x03, addr, 1)
                crc = compute_crc(req)
                req += struct.pack('<H', crc)
                resp = self.transport.send_and_receive(req)
                byte_num = resp[2]
                data = resp[3:3+byte_num]
                if len(data) >= 2:
                    val = data[0] << 8 | data[1]
                elif len(data) == 1:
                    val = data[0]
                else:
                    val = 0
                self.tk_parent.after(0, lambda: label.delete(0, 'end') or label.insert(0, str(val)))
            except Exception as e:
                self.tk_parent.after(0, (lambda exc=e: messagebox.showerror('Read error', str(exc))))
        threading.Thread(target=worker, daemon=True).start()

    def write_param(self, p, entry, label):
        def worker():
            try:
                vtext = entry.get().strip()
                if vtext == '':
                    raise ValueError('No value')
                val = int(vtext)
                addr = int(p['id'])
                req = struct.pack('>B B H H', self.drive_id, 0x06, addr, val)
                crc = compute_crc(req)
                req += struct.pack('<H', crc)
                resp = self.transport.send_and_receive(req)
                self.tk_parent.after(0, lambda: None)
            except Exception as e:
                self.tk_parent.after(0, (lambda exc=e: messagebox.showerror('Write error', str(exc))))
        threading.Thread(target=worker, daemon=True).start()

    def read_all(self):
        def worker():
            errors = []
            for p in self.params:
                widgets = self.param_widgets.get(p['id'])
                if not widgets:
                    continue
                entry = widgets[0]
                try:
                    addr = int(p['id'])
                    req = struct.pack('>B B H H', self.drive_id, 0x03, addr, 1)
                    crc = compute_crc(req)
                    req += struct.pack('<H', crc)
                    resp = self.transport.send_and_receive(req)
                    byte_num = resp[2]
                    data = resp[3:3+byte_num]
                    if len(data) >= 2:
                        val = data[0] << 8 | data[1]
                    elif len(data) == 1:
                        val = data[0]
                    else:
                        val = 0
                    # update UI
                    self.tk_parent.after(0, lambda e=entry, v=val: (e.delete(0, 'end'), e.insert(0, str(v))))
                except Exception as exc:
                    errors.append(f"{p.get('name','id'+str(p.get('id')))}: {exc}")
            if errors:
                # aggregate errors into a single dialog (limit to first 20 to avoid huge dialogs)
                max_show = 20
                msg_lines = errors[:max_show]
                if len(errors) > max_show:
                    msg_lines.append(f"... and {len(errors)-max_show} more errors")
                msg = "\n".join(msg_lines)
                self.tk_parent.after(0, lambda: messagebox.showerror('Read-All Errors', msg))

        threading.Thread(target=worker, daemon=True).start()

    def refresh_status(self):
        """Read status registers (in chunks up to 8) and update the status tree."""
        def worker():
            try:
                entries = sorted(self.status_entries, key=lambda x: int(x.get('id', 0)))
                # build chunks of contiguous addresses up to 8 registers
                i = 0
                n = len(entries)
                while i < n:
                    start = int(entries[i]['id'])
                    count = 1
                    j = i + 1
                    while j < n and int(entries[j]['id']) == start + count and count < 8:
                        count += 1
                        j += 1
                    # perform read for this chunk
                    vals = self.transport.read_status(self.drive_id, start, count, func=0x04)
                    # update UI for each value
                    for k, v in enumerate(vals):
                        addr = start + k
                        iid = str(addr)
                        try:
                            item = self.status_tree.item(iid)
                            if item:
                                vals_now = list(item.get('values', []))
                                # values tuple: (id, name, value, units, desc)
                                if len(vals_now) >= 3:
                                    vals_now[2] = str(v)
                                    self.tk_parent.after(0, lambda iid=iid, vlist=vals_now: self.status_tree.item(iid, values=vlist))
                        except Exception:
                            pass
                    i = j
            except Exception as e:
                self.tk_parent.after(0, (lambda exc=e: messagebox.showerror('Status read error', str(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_status_generic(self, entries, tree, func):
        def worker():
            try:
                entries_sorted = sorted(entries, key=lambda x: int(x.get('id', 0)))
                i = 0
                n = len(entries_sorted)
                while i < n:
                    start = int(entries_sorted[i]['id'])
                    count = 1
                    j = i + 1
                    while j < n and int(entries_sorted[j]['id']) == start + count and count < 8:
                        count += 1
                        j += 1
                    vals = self.transport.read_status(self.drive_id, start, count, func=func)
                    for k, v in enumerate(vals):
                        addr = start + k
                        iid = str(addr)
                        try:
                            item = tree.item(iid)
                            if item:
                                vals_now = list(item.get('values', []))
                                if len(vals_now) >= 3:
                                    vals_now[2] = str(v)
                                    # update the tree item
                                    self.tk_parent.after(0, lambda iid=iid, vlist=vals_now, t=tree: t.item(iid, values=vlist))
                                    # autosize the column for 'value' if needed
                                    try:
                                        f = tkfont.nametofont(tree.cget('font'))
                                    except Exception:
                                        f = None
                                    if f is not None:
                                        # column id of value is the third column in tree['columns']
                                        col_id = tree['columns'][2]
                                        text_w = f.measure(str(vals_now[2]))
                                        hdr_w = f.measure(tree.heading(col_id)['text'])
                                        curr_w = tree.column(col_id)['width']
                                        desired = max(text_w, hdr_w) + 20
                                        if desired > curr_w:
                                            self.tk_parent.after(0, lambda c=col_id, w=desired, t=tree: t.column(c, width=w))
                        except Exception:
                            pass
                    i = j
            except Exception as e:
                self.tk_parent.after(0, (lambda exc=e: messagebox.showerror('Status read error', str(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def refresh_status_04(self):
        self._refresh_status_generic(self.status_entries_04, self.status_tree_04, 0x04)

    def refresh_status_03(self):
        self._refresh_status_generic(self.status_entries_03, self.status_tree_03, 0x03)

    def _set_tab_enabled(self, enabled: bool):
        state = 'normal' if enabled else 'disabled'
        try:
            self.enable_btn.config(state=state)
            self.read_all_btn.config(state=state)
            self.save_eeprom_btn.config(state=state)
            if hasattr(self, 'close_btn'):
                self.close_btn.config(state=state)
            for wid in self.param_widgets.values():
                entry = wid[0]
                entry.config(state=state)
                # read and write buttons
                if len(wid) > 1:
                    wid[1].config(state=state)
                if len(wid) > 2:
                    wid[2].config(state=state)
        except Exception:
            pass

    def save_eeprom(self):
        # Write value 0x1234 to address 0x1001 using function 0x06
        addr = EEPROM_SAVE_ADDR
        val = EEPROM_SAVE_VALUE

        def worker():
            try:
                # disable tab controls while saving
                self.tk_parent.after(0, lambda: self._set_tab_enabled(False))
                req = struct.pack('>B B H H', self.drive_id, 0x06, addr, val)
                crc = compute_crc(req)
                req += struct.pack('<H', crc)
                # send and don't expect immediate long response (drive will take time to save)
                self.transport.send_and_receive(req)
                # inform user and wait EEPROM_WAIT_SECONDS before re-enabling
                self.tk_parent.after(0, lambda: messagebox.showinfo('EEPROM', f'Save command sent. Waiting ~{EEPROM_WAIT_SECONDS}s for EEPROM write.'))
                time.sleep(EEPROM_WAIT_SECONDS)
                self.tk_parent.after(0, lambda: messagebox.showinfo('EEPROM', 'EEPROM save should be complete.'))
            except Exception as e:
                self.tk_parent.after(0, (lambda exc=e: messagebox.showerror('EEPROM save error', str(exc))))
            finally:
                # re-enable tab controls
                self.tk_parent.after(0, lambda: self._set_tab_enabled(True))

        threading.Thread(target=worker, daemon=True).start()


class App:
    def __init__(self, root, xml_path, transport_debug: bool = False):
        self.root = root
        self.root.title('RS485 Servo GUI')
        self.transport = SerialTransport(debug=bool(transport_debug))
        self.params = load_parameters(xml_path)
        # determine description column width (characters) based on longest description
        try:
            f = tkfont.nametofont('TkDefaultFont')
            if self.params:
                max_px = max((f.measure(p.get('description','')) for p in self.params))
            else:
                max_px = f.measure('Description') * 6
            avg_char = max(1, f.measure('0'))
            # add small padding
            desc_chars = max(20, int(max_px / avg_char) + 2)
        except Exception:
            desc_chars = 40
        self.desc_width_chars = desc_chars
        # load status definitions for func 0x04 and 0x03
        self.status_entries_04 = load_status(os.path.join('config', 'status_04.xml'))
        self.status_entries_03 = load_status(os.path.join('config', 'status_03.xml'))
        self.config_path = os.path.join('config', 'gui_settings.json')
        self.config = self.load_config()
        self._build_ui()
        # don't show saved drives until connected
        self._saved_drives = list(self.config.get('drives', []))

    def _build_ui(self):
        conn = ttk.LabelFrame(self.root, text='Connection')
        conn.pack(fill='x', padx=6, pady=6)

        ttk.Label(conn, text='COM Port').grid(row=0, column=0, padx=4, pady=4)
        self.port_cb = ttk.Combobox(conn, values=self.list_com_ports(), width=12)
        self.port_cb.grid(row=0, column=1, padx=4, pady=4)
        if self.config.get('port'):
            self.port_cb.set(self.config.get('port'))

        ttk.Label(conn, text='Baudrate').grid(row=0, column=2, padx=4, pady=4)
        self.baud_cb = ttk.Combobox(conn, values=[9600,19200,38400,57600,115200], width=10)
        self.baud_cb.set(115200)
        self.baud_cb.grid(row=0, column=3, padx=4, pady=4)
        if self.config.get('baud'):
            try:
                self.baud_cb.set(int(self.config.get('baud')))
            except Exception:
                pass

        ttk.Label(conn, text='Parity').grid(row=0, column=4, padx=4, pady=4)
        self.par_cb = ttk.Combobox(conn, values=['N','E','O'], width=4)
        self.par_cb.set('N')
        self.par_cb.grid(row=0, column=5, padx=4, pady=4)
        if self.config.get('parity'):
            self.par_cb.set(self.config.get('parity'))

        self.connect_btn = ttk.Button(conn, text='Connect', command=self.toggle_connect)
        self.connect_btn.grid(row=0, column=6, padx=6)

        drives_frame = ttk.LabelFrame(self.root, text='Drives')
        drives_frame.pack(fill='both', expand=True, padx=6, pady=6)

        add_frame = ttk.Frame(drives_frame)
        add_frame.pack(fill='x', padx=4, pady=4)
        ttk.Label(add_frame, text='Drive ID').pack(side='left')
        self.drive_entry = ttk.Entry(add_frame, width=6)
        self.drive_entry.pack(side='left', padx=4)
        self.add_drive_btn = ttk.Button(add_frame, text='Add Drive', command=self.add_drive)
        self.add_drive_btn.pack(side='left', padx=6)

        # disable drive controls until connected
        self.enable_drive_controls(False)

        self.notebook = ttk.Notebook(drives_frame)
        self.notebook.pack(fill='both', expand=True)

        self.drive_tabs = {}

    def list_com_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        return ports

    def toggle_connect(self):
        if self.transport.ser and self.transport.ser.is_open:
            self.transport.close()
            self.connect_btn.config(text='Connect')
            messagebox.showinfo('Disconnected','Serial port closed')
            self.save_config()
            # hide drives when disconnected
            self.enable_drive_controls(False)
            self.hide_all_drives()
        else:
            port = self.port_cb.get()
            if not port:
                messagebox.showwarning('No port','Choose a COM port')
                return
            baud = int(self.baud_cb.get())
            parity = self.par_cb.get()
            try:
                self.transport.open(port, baud, parity)
                self.connect_btn.config(text='Disconnect')
                messagebox.showinfo('Connected',f'Opened {port} {baud} {parity}')
                self.config['port'] = port
                self.config['baud'] = baud
                self.config['parity'] = parity
                self.save_config()
                # enable drive UI and show saved drives
                self.enable_drive_controls(True)
                self.show_saved_drives()
            except Exception as e:
                messagebox.showerror('Open error', str(e))

    def add_drive(self, did_str: str = None, save: bool = True):
        txt = did_str if did_str is not None else self.drive_entry.get().strip()
        if not txt:
            return
        try:
            did = int(txt)
        except ValueError:
            messagebox.showwarning('Invalid', 'Drive ID must be integer')
            return
        if did in self.drive_tabs:
            messagebox.showinfo('Exists','Drive already added')
            return
        # only allow adding drives when connected
        if not (self.transport.ser and self.transport.ser.is_open):
            messagebox.showwarning('Not connected', 'Connect to serial port first')
            return
        tab = DriveTab(self.notebook, did, self.transport, self.params, self.root, close_callback=lambda d=did: self.remove_drive(d), status_entries_04=self.status_entries_04, status_entries_03=self.status_entries_03, desc_width_chars=self.desc_width_chars)
        self.drive_tabs[did] = tab
        self.notebook.add(tab.frame, text=f'Drive {did}')
        if save:
            self.config.setdefault('drives', [])
            if did not in self.config['drives']:
                self.config['drives'].append(did)
            self.save_config()

    def enable_drive_controls(self, enable: bool):
        state = 'normal' if enable else 'disabled'
        try:
            self.drive_entry.config(state=state)
            self.add_drive_btn.config(state=state)
        except Exception:
            pass

    def hide_all_drives(self):
        # remove all drive tabs from notebook but keep config
        for did in list(self.drive_tabs.keys()):
            try:
                tab = self.drive_tabs[did]
                self.notebook.forget(tab.frame)
            except Exception:
                pass
            del self.drive_tabs[did]

    def show_saved_drives(self):
        # create tabs for drives in config if not already present
        for did in self._saved_drives:
            if did in self.drive_tabs:
                continue
            try:
                self.add_drive(did_str=str(did), save=False)
            except Exception:
                pass

    def remove_drive(self, did: int):
        if did not in self.drive_tabs:
            return
        tab = self.drive_tabs[did]
        # remove from notebook
        try:
            self.notebook.forget(tab.frame)
        except Exception:
            pass
        # delete internal
        del self.drive_tabs[did]
        # remove from config and save
        try:
            if 'drives' in self.config and did in self.config['drives']:
                self.config['drives'].remove(did)
            self.save_config()
        except Exception:
            pass

    def load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def save_config(self):
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            messagebox.showwarning('Config save failed', str(e))
