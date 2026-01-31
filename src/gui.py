import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import struct
import time
import tkinter.font as tkfont
import xml.etree.ElementTree as ET
import serial.tools.list_ports
from functools import partial
import json
import os
import datetime
from tkinter.scrolledtext import ScrolledText

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
    # simple on-disk cache to avoid reparsing XML repeatedly
    try:
        global _PARAM_CACHE
    except NameError:
        _PARAM_CACHE = {}
    if xml_path in globals().get('_PARAM_CACHE', {}):
        return globals()['_PARAM_CACHE'][xml_path]
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
    # store in cache
    try:
        _PARAM_CACHE[xml_path] = params
    except Exception:
        pass
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
    def __init__(self, parent, drive_id, transport: SerialTransport, params, tk_parent, app, close_callback=None, status_entries_04=None, status_entries_03=None, desc_width_chars: int = 40):
        self.drive_id = drive_id
        self.transport = transport
        self.params = params
        self.status_entries_04 = status_entries_04 or []
        self.status_entries_03 = status_entries_03 or []
        self.desc_width_chars = desc_width_chars
        self.frame = ttk.Frame(parent)
        self.tk_parent = tk_parent
        self.app = app
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
            # Mouse wheel support: bind while cursor is over this canvas
            def _on_mousewheel(event):
                # Normalize for platform
                if getattr(event, 'num', None) == 4:
                    delta = -1
                elif getattr(event, 'num', None) == 5:
                    delta = 1
                else:
                    try:
                        delta = int(-1 * (event.delta / 120))
                    except Exception:
                        delta = 0
                # scroll fewer units for smoother motion
                canvas.yview_scroll(delta * 3, 'units')

            def _on_enter(e):
                # bind mouse wheel globally while over this canvas
                canvas.bind_all('<MouseWheel>', _on_mousewheel)
                canvas.bind_all('<Button-4>', _on_mousewheel)
                canvas.bind_all('<Button-5>', _on_mousewheel)

            def _on_leave(e):
                try:
                    canvas.unbind_all('<MouseWheel>')
                except Exception:
                    pass
                try:
                    canvas.unbind_all('<Button-4>')
                    canvas.unbind_all('<Button-5>')
                except Exception:
                    pass

            canvas.bind('<Enter>', _on_enter)
            canvas.bind('<Leave>', _on_leave)

            return inner

        page0 = make_param_page('Params 0xx')
        page1 = make_param_page('Params 1xx')
        page2 = make_param_page('Params 2xx')
        page3 = make_param_page('Params 3xx')

        self.param_widgets = {}

        # column headers for each page
        def add_headers(parent):
            header = ttk.Frame(parent)
            header.pack(fill='x', padx=4, pady=(2,6))
            ttk.Label(header, text='', width=2).grid(row=0, column=0)
            ttk.Label(header, text='Name', width=14).grid(row=0, column=1, sticky='w')
            ttk.Label(header, text='Description', width=self.desc_width_chars).grid(row=0, column=2, sticky='w')
            ttk.Label(header, text='Min', width=8).grid(row=0, column=3)
            ttk.Label(header, text='Value', width=10).grid(row=0, column=4)
            ttk.Label(header, text='Max', width=8).grid(row=0, column=5)
            ttk.Label(header, text='').grid(row=0, column=6)

        add_headers(page0)
        add_headers(page1)
        add_headers(page2)
        add_headers(page3)

        # Distribute parameters into page lists based on id ranges
        params_page0 = []
        params_page1 = []
        params_page2 = []
        params_page3 = []
        for p in self.params:
            try:
                pid = int(p.get('id', 0))
            except Exception:
                pid = 0
            if 0 <= pid <= 99:
                params_page0.append(p)
            elif 100 <= pid <= 199:
                params_page1.append(p)
            elif 200 <= pid <= 299:
                params_page2.append(p)
            else:
                params_page3.append(p)

        # Populate rows in small batches to keep UI responsive
        def populate_page(parent, params_list, drive_tab=None):
            batch = 30
            def add_batch(start=0):
                end = min(start + batch, len(params_list))
                for idx in range(start, end):
                    p = params_list[idx]
                    row = ttk.Frame(parent)
                    row.pack(fill='x', padx=4, pady=2)
                    # star button for favorites
                    pid_str = str(p.get('id',''))
                    is_fav = False
                    try:
                        fav_list = self.app.config.get('favorites', {}).get(str(self.drive_id), [])
                        is_fav = ('p:' + pid_str) in fav_list
                    except Exception:
                        is_fav = False
                    star_txt = '★' if is_fav else '☆'
                    star_btn = ttk.Button(row, text=star_txt, width=2)
                    star_btn.grid(row=0, column=0)

                    name_lbl = ttk.Label(row, text=f"{p.get('name','')} ({p.get('id','')})", width=14, anchor='w', font=('Segoe UI', 9, 'bold'))
                    name_lbl.grid(row=0, column=1, sticky='w')
                    desc_lbl = ttk.Label(row, text=p.get('description',''), width=self.desc_width_chars, anchor='w')
                    desc_lbl.grid(row=0, column=2, sticky='w', padx=(6,0))
                    min_lbl = ttk.Label(row, text=str(p.get('min','')), width=8)
                    min_lbl.grid(row=0, column=3)
                    entry = ttk.Entry(row, width=10)
                    entry.insert(0, str(p.get('value','')))
                    entry.grid(row=0, column=4, padx=6)
                    max_lbl = ttk.Label(row, text=str(p.get('max','')), width=8)
                    max_lbl.grid(row=0, column=5)
                    read_btn = ttk.Button(row, text='Read', command=partial(self.read_param, p, entry))
                    read_btn.grid(row=0, column=6, padx=2)
                    write_btn = ttk.Button(row, text='Write', command=partial(self.write_param, p, entry, entry))
                    write_btn.grid(row=0, column=7, padx=2)
                    # store widgets keyed by string id for consistency
                    self.param_widgets[str(p['id'])] = {
                        'entry': entry,
                        'read': read_btn,
                        'write': write_btn,
                        'star': star_btn,
                    }

                    # local toggle handler
                    def _toggle_local(pid=pid_str, btn=star_btn):
                        new_state = self.app.toggle_favorite(self.drive_id, 'p:' + pid)
                        try:
                            btn.config(text='★' if new_state else '☆')
                        except Exception:
                            pass
                        # refresh local and global favorite views
                        try:
                            self.refresh_local_favorites()
                        except Exception:
                            pass
                        try:
                            self.app.refresh_global_favorites()
                        except Exception:
                            pass

                    star_btn.config(command=_toggle_local)
                if end < len(params_list):
                    # schedule next batch, small delay to keep UI responsive
                    parent.after(20, add_batch, end)
            # start populating
            parent.after(10, add_batch, 0)

        populate_page(page0, params_page0)
        populate_page(page1, params_page1)
        populate_page(page2, params_page2)
        populate_page(page3, params_page3)

        # After parameter rows are populated in batches, refresh the local
        # favorites view so it can show live values. Delay allows the
        # batches to create widgets first.
        try:
            self.tk_parent.after(500, self.refresh_local_favorites)
        except Exception:
            pass

        # Status tab as separate tabs: Status 04 and Status 03
        status_page = ttk.Frame(param_notebook.master)
        param_notebook.add(status_page, text='Status 04')
        status_page_03 = ttk.Frame(param_notebook.master)
        param_notebook.add(status_page_03, text='Status 03')

        # Favorites tab (local to this drive)
        fav_page = ttk.Frame(param_notebook.master)
        param_notebook.add(fav_page, text='Favorites')
        self.fav_tree = ttk.Treeview(fav_page, columns=('addr','desc','min','max','value','units'), show='headings', height=10)
        self.fav_tree.pack(fill='both', expand=True, padx=4, pady=4)
        for c,h in (('addr','Addr'),('desc','Description'),('min','Min'),('max','Max'),('value','Value'),('units','Units')):
            self.fav_tree.heading(c, text=h)
            if c == 'desc':
                self.fav_tree.column(c, width=360)
            elif c in ('min','max'):
                self.fav_tree.column(c, width=80, anchor='e')
            else:
                self.fav_tree.column(c, width=120)
        # initial populate will be done after rows are created
        # add Read/Write buttons for this drive's favorites
        try:
            fav_ops = ttk.Frame(fav_page)
            fav_ops.pack(fill='x', padx=4, pady=(0,4))
            self.fav_read_btn = ttk.Button(fav_ops, text='Read Selected', command=self.read_selected_local_favorite)
            self.fav_read_btn.pack(side='left')
            self.fav_write_btn = ttk.Button(fav_ops, text='Write Selected', command=self.write_selected_local_favorite)
            self.fav_write_btn.pack(side='left', padx=6)
            try:
                self.fav_write_btn.config(state='disabled')
            except Exception:
                pass
            # enable/disable write button based on selection
            try:
                self.fav_tree.bind('<<TreeviewSelect>>', lambda e: self._on_local_fav_select())
            except Exception:
                pass
        except Exception:
            pass
        # initial populate will be done after rows are created

        # --- Status 04 UI ---
        s_top = ttk.Frame(status_page)
        s_top.pack(fill='x')
        self.status_refresh_btn_04 = ttk.Button(s_top, text='Refresh Status 04', command=self.refresh_status_04)
        self.status_refresh_btn_04.pack(side='left')
        # Columns: star | addr(hex) | Description | Value | Units
        cols = ('star','addr', 'desc', 'value', 'units')
        self.status_tree_04 = ttk.Treeview(status_page, columns=cols, show='headings', height=10)
        self.status_tree_04.pack(fill='both', expand=True, padx=4, pady=4)
        self.status_tree_04.heading('star', text='')
        self.status_tree_04.column('star', width=28, anchor='center', stretch=False)
        self.status_tree_04.heading('addr', text='Addr')
        self.status_tree_04.heading('desc', text='Description')
        self.status_tree_04.heading('value', text='Value')
        self.status_tree_04.heading('units', text='Units')
        self.status_tree_04.column('addr', width=80, anchor='center')
        self.status_tree_04.column('desc', width=360)
        self.status_tree_04.column('value', width=120, anchor='e')
        self.status_tree_04.column('units', width=80)
        # configure tags for coloring favorite rows
        try:
            self.status_tree_04.tag_configure('fav', foreground='orange')
            self.status_tree_04.tag_configure('nofav', foreground='black')
        except Exception:
            pass
        for s in self.status_entries_04:
            sid = int(s.get('id', 0))
            addr_text = format(sid, '#06x')
            desc = s.get('description','')
            val = s.get('value','')
            units = s.get('units','')
            iid = f's:04:{sid}'
            isfav = (f's:04:{sid}') in self.app.config.get('favorites', {}).get(str(self.drive_id), [])
            star = '★' if isfav else '☆'
            tag = 'fav' if isfav else 'nofav'
            self.status_tree_04.insert('', 'end', iid=iid, values=(star, addr_text, desc, val, units), tags=(tag,))
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
        try:
            # single-click star column to toggle favorite
            self.status_tree_04.bind('<Button-1>', lambda e: self._on_status_tree_click(e, '04', self.status_tree_04))
            # mouse wheel while over tree
            def _enter04(e):
                self.status_tree_04.bind_all('<MouseWheel>', lambda ev: self.status_tree_04.yview_scroll(int(-1*(ev.delta/120)), 'units'))
            def _leave04(e):
                try:
                    self.status_tree_04.unbind_all('<MouseWheel>')
                except Exception:
                    pass
            self.status_tree_04.bind('<Enter>', _enter04)
            self.status_tree_04.bind('<Leave>', _leave04)
        except Exception:
            pass

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
        try:
            self.status_tree_03.tag_configure('fav', foreground='orange')
            self.status_tree_03.tag_configure('nofav', foreground='black')
        except Exception:
            pass
        for s in self.status_entries_03:
            sid = int(s.get('id', 0))
            addr_text = format(sid, '#06x')
            desc = s.get('description','')
            val = s.get('value','')
            units = s.get('units','')
            iid = f's:03:{sid}'
            isfav = (f's:03:{sid}') in self.app.config.get('favorites', {}).get(str(self.drive_id), [])
            star = '★' if isfav else '☆'
            tag = 'fav' if isfav else 'nofav'
            self.status_tree_03.insert('', 'end', iid=iid, values=(star, addr_text, desc, val, units), tags=(tag,))
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
        try:
            self.status_tree_03.bind('<Button-1>', lambda e: self._on_status_tree_click(e, '03', self.status_tree_03))
            def _enter03(e):
                self.status_tree_03.bind_all('<MouseWheel>', lambda ev: self.status_tree_03.yview_scroll(int(-1*(ev.delta/120)), 'units'))
            def _leave03(e):
                try:
                    self.status_tree_03.unbind_all('<MouseWheel>')
                except Exception:
                    pass
            self.status_tree_03.bind('<Enter>', _enter03)
            self.status_tree_03.bind('<Leave>', _leave03)
        except Exception:
            pass

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

    def refresh_local_favorites(self):
        """Populate the local Favorites tree for this drive."""
        try:
            if not hasattr(self, 'fav_tree') or not self.fav_tree:
                return
            # clear
            for iid in self.fav_tree.get_children():
                self.fav_tree.delete(iid)
            favs = self.app.config.get('favorites', {}).get(str(self.drive_id), [])
            for fav in favs:
                try:
                    if isinstance(fav, str) and fav.startswith('p:'):
                        pid = fav.split(':', 1)[1]
                        pinfo = next((p for p in self.params if str(p.get('id','')) == str(pid)), None)
                        desc = pinfo.get('description','') if pinfo else ''
                        min_val = pinfo.get('min','') if pinfo else ''
                        max_val = pinfo.get('max','') if pinfo else ''
                        val = ''
                        try:
                            w = self.param_widgets.get(str(pid))
                            if w:
                                if isinstance(w, dict):
                                    val = w.get('entry').get()
                                else:
                                    # legacy tuple
                                    val = w[0].get()
                        except Exception:
                            val = ''
                        addr_text = str(pid)
                        units = pinfo.get('units','') if pinfo else ''
                        self.fav_tree.insert('', 'end', iid=fav, values=(addr_text, desc, min_val, max_val, val, units))
                    elif isinstance(fav, str) and fav.startswith('s:'):
                        # format: s:<func>:<addr>
                        parts = fav.split(':')
                        func = parts[1] if len(parts) > 1 else '04'
                        addr = parts[2] if len(parts) > 2 else parts[-1]
                        desc = ''
                        units = ''
                        val = ''
                        # lookup description from status entries
                        try:
                            sid = int(addr)
                        except Exception:
                            sid = None
                        if func == '03':
                            entry = next((s for s in self.status_entries_03 if int(s.get('id',0)) == sid), None)
                            tree = getattr(self, 'status_tree_03', None)
                        else:
                            entry = next((s for s in self.status_entries_04 if int(s.get('id',0)) == sid), None)
                            tree = getattr(self, 'status_tree_04', None)
                        if entry:
                            desc = entry.get('description','')
                            units = entry.get('units','')
                        # try to read live value from tree if present
                        try:
                            iid = f's:{func}:{sid}' if sid is not None else None
                            if tree is not None and iid is not None and tree.exists(iid):
                                val = tree.set(iid, 'value')
                        except Exception:
                            val = ''
                        addr_text = f"{func}:{addr}"
                        # status entries don't have min/max
                        self.fav_tree.insert('', 'end', iid=fav, values=(addr_text, desc, '', '', val, units))
                except Exception:
                    pass
        except Exception:
            pass

    def _on_local_fav_select(self, event=None):
        """Enable or disable the drive-local Write button depending on selection.

        If a status favorite (iid starting with 's:') is selected, disable write.
        """
        try:
            sel = self.fav_tree.selection()
            if not sel:
                try:
                    self.fav_write_btn.config(state='disabled')
                except Exception:
                    pass
                return
            iid = sel[0]
            try:
                if isinstance(iid, str) and iid.startswith('s:'):
                    self.fav_write_btn.config(state='disabled')
                else:
                    self.fav_write_btn.config(state='normal')
            except Exception:
                pass
        except Exception:
            pass

    def read_selected_local_favorite(self):
        """Read the currently selected local favorite row and update its value."""
        try:
            sel = self.fav_tree.selection()
            if not sel:
                return
            iid = sel[0]
            fav = iid
            # reuse existing read_fav
            try:
                self.read_fav(fav)
                # also append to global log
                try:
                    self.app.append_log(f"Drive {self.drive_id} read favorite {fav}")
                except Exception:
                    pass
            except Exception as e:
                try:
                    messagebox.showerror('Read favorite', str(e))
                except Exception:
                    pass
        except Exception:
            pass

    def write_selected_local_favorite(self):
        """Prompt for a value and write it to the selected favorite if it's a parameter."""
        try:
            sel = self.fav_tree.selection()
            if not sel:
                return
            iid = sel[0]
            fav = iid
            if isinstance(fav, str) and fav.startswith('p:'):
                pid = fav.split(':',1)[1]
            else:
                # legacy numeric or invalid
                if isinstance(fav, str) and fav.startswith('s:'):
                    messagebox.showwarning('Write', 'Cannot write to status favorites')
                    return
                pid = str(fav)
            # ask for value
            v = simpledialog.askstring('Write Parameter', f'Value for parameter {pid}:')
            if v is None:
                return
            try:
                val = int(v)
            except Exception:
                messagebox.showwarning('Invalid', 'Value must be integer')
                return
            # perform write
            addr = int(pid)
            try:
                req = struct.pack('>B B H H', self.drive_id, 0x06, addr, val)
                crc = compute_crc(req)
                req += struct.pack('<H', crc)
                self.transport.send_and_receive(req)
                # update UI cell if present
                try:
                    if self.fav_tree.exists(fav):
                        self.fav_tree.set(fav, 'value', str(val))
                except Exception:
                    pass
                try:
                    self.app.append_log(f"Drive {self.drive_id} wrote param {pid} = {val}")
                except Exception:
                    pass
            except Exception as e:
                messagebox.showerror('Write error', str(e))
        except Exception:
            pass

    def _toggle_status_fav(self, func_str, tree):
        """Toggle favorite for a status row (func_str '04' or '03')."""
        try:
            sel = tree.selection()
            if not sel:
                return
            for iid in sel:
                fav_key = f's:{func_str}:{iid}'
                self.app.toggle_favorite(self.drive_id, fav_key)
            # refresh views
            try:
                self.refresh_local_favorites()
            except Exception:
                pass
            try:
                self.app.refresh_global_favorites()
            except Exception:
                pass
        except Exception:
            pass

    def _on_status_tree_click(self, event, func_str, tree):
        """Handle single-clicks in the status tree. Toggle favorite when star column clicked."""
        try:
            col = tree.identify_column(event.x)
            row = tree.identify_row(event.y)
            if not row:
                return
            # star column is '#1'
            if col == '#1':
                iid = row
                # iid may already be 's:04:16'
                if not iid.startswith('s:'):
                    fav_key = f's:{func_str}:{iid}'
                else:
                    fav_key = iid
                self.app.toggle_favorite(self.drive_id, fav_key)
                # update visual immediately for this drive
                try:
                    current = set(self.app.config.get('favorites', {}).get(str(self.drive_id), []))
                    self.apply_favorite_states(current)
                except Exception:
                    pass
                try:
                    self.app.refresh_global_favorites()
                except Exception:
                    pass
                # update row tag color
                try:
                    isfav = fav_key in set(self.app.config.get('favorites', {}).get(str(self.drive_id), []))
                    tag = 'fav' if isfav else 'nofav'
                    tree.item(iid, tags=(tag,))
                except Exception:
                    pass
        except Exception:
            pass

    def apply_favorite_states(self, favs):
        """Update star visuals for parameters and status rows using favs (iterable of fav keys)."""
        try:
            fav_set = set(favs) if not isinstance(favs, set) else favs
        except Exception:
            fav_set = set()
        # update parameter stars
        try:
            for pid, widgets in list(self.param_widgets.items()):
                try:
                    fav_key = f'p:{pid}'
                    star = '★' if fav_key in fav_set else '☆'
                    if isinstance(widgets, dict):
                        btn = widgets.get('star')
                        if btn:
                            btn.config(text=star)
                    else:
                        # legacy tuple: star is last element
                        try:
                            widgets[-1].config(text=star)
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass
        # update status tree stars
        try:
            for tree_attr in ('status_tree_04', 'status_tree_03'):
                tree = getattr(self, tree_attr, None)
                if tree is None:
                    continue
                for iid in tree.get_children():
                    try:
                        # iid expected format 's:04:16' or 's:03:10'
                        isfav = iid in fav_set
                        star = '★' if isfav else '☆'
                        tree.set(iid, 'star', star)
                    except Exception:
                        pass
        except Exception:
            pass

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

    def read_fav(self, fav):
        """Read a single favorite entry (p:<id> or s:<func>:<addr>)."""
        try:
            if isinstance(fav, str) and fav.startswith('p:'):
                pid = fav.split(':',1)[1]
                p = next((x for x in self.params if str(x.get('id','')) == str(pid)), None)
                if not p:
                    return
                widgets = self.param_widgets.get(str(pid))
                if not widgets:
                    return
                entry = widgets.get('entry') if isinstance(widgets, dict) else widgets[0]
                # reuse existing read_param logic
                self.read_param(p, entry)
            elif isinstance(fav, str) and fav.startswith('s:'):
                parts = fav.split(':')
                func = int(parts[1]) if len(parts) > 1 else 4
                addr = int(parts[2]) if len(parts) > 2 else None
                if addr is None:
                    return
                self.read_status_item(func, addr)
            else:
                # legacy numeric
                pid = fav
                p = next((x for x in self.params if str(x.get('id','')) == str(pid)), None)
                if not p:
                    return
                widgets = self.param_widgets.get(str(pid))
                if not widgets:
                    return
                entry = widgets.get('entry') if isinstance(widgets, dict) else widgets[0]
                self.read_param(p, entry)
        except Exception:
            pass

    def read_status_item(self, func, addr):
        """Read a single status register and update the tree value."""
        def worker():
            try:
                vals = self.transport.read_status(self.drive_id, addr, 1, func=func)
                if not vals:
                    return
                v = vals[0]
                iid = f's:{func:02d}:{addr}' if isinstance(func, int) else f's:{func}:{addr}'
                # try update both trees if exists
                try:
                    if hasattr(self, 'status_tree_04') and self.status_tree_04.exists(iid):
                        self.tk_parent.after(0, lambda: self.status_tree_04.set(iid, 'value', str(v)))
                    if hasattr(self, 'status_tree_03') and self.status_tree_03.exists(iid):
                        self.tk_parent.after(0, lambda: self.status_tree_03.set(iid, 'value', str(v)))
                except Exception:
                    pass
            except Exception:
                pass
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
                widgets = self.param_widgets.get(str(p['id']))
                if not widgets:
                    continue
                # support new dict or legacy tuple
                if isinstance(widgets, dict):
                    entry = widgets.get('entry')
                else:
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

    def refresh_local_favorites_for_drive(self, drive_id):
        # helper to call drive tab refresh when App wants to refresh
        dt = self.drive_tabs.get(drive_id)
        if dt:
            try:
                dt.refresh_local_favorites()
            except Exception:
                pass

    def _set_tab_enabled(self, enabled: bool):
        state = 'normal' if enabled else 'disabled'
        try:
            self.enable_btn.config(state=state)
            self.read_all_btn.config(state=state)
            self.save_eeprom_btn.config(state=state)
            if hasattr(self, 'close_btn'):
                self.close_btn.config(state=state)
            for wid in self.param_widgets.values():
                # support dict or tuple
                if isinstance(wid, dict):
                    entry = wid.get('entry')
                    if entry:
                        entry.config(state=state)
                    if wid.get('read'):
                        wid.get('read').config(state=state)
                    if wid.get('write'):
                        wid.get('write').config(state=state)
                    if wid.get('star'):
                        wid.get('star').config(state=state)
                else:
                    try:
                        entry = wid[0]
                        entry.config(state=state)
                        if len(wid) > 1:
                            wid[1].config(state=state)
                        if len(wid) > 2:
                            wid[2].config(state=state)
                    except Exception:
                        pass
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
        # ensure favorites structure exists
        if 'favorites' not in self.config:
            self.config['favorites'] = {}
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

        # NOTE: Global favorites are shown as a permanent, left-most tab
        # in the drives notebook. The actual Notebook is created below; we
        # will add the favorites tab after creating the notebook so it
        # appears first.

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

        # create a permanent left-most tab for global favorites
        self.global_fav_tab = ttk.Frame(self.notebook)
        # Insert at 0 so it's the left-most tab; adding then moving is fine
        self.notebook.add(self.global_fav_tab, text='Favorites (Global)')
        fav_top = ttk.Frame(self.global_fav_tab)
        fav_top.pack(fill='x')
        ttk.Label(fav_top, text='Favorites').pack(side='left')
        # Controls: Refresh, Read All, Auto-Read (interval seconds)
        refresh_fav_btn = ttk.Button(fav_top, text='Refresh Favorites', command=self.refresh_global_favorites)
        refresh_fav_btn.pack(side='right')
        read_fav_btn = ttk.Button(fav_top, text='Read Favorites', command=self.read_all_favorites)
        read_fav_btn.pack(side='right', padx=6)
        ttk.Label(fav_top, text='Auto-Read (s)').pack(side='right', padx=(8,2))
        self.autoread_interval = ttk.Entry(fav_top, width=6)
        self.autoread_interval.insert(0, '5')
        self.autoread_interval.pack(side='right')
        self.autoread_btn = ttk.Button(fav_top, text='Auto-Read: Off', command=self._toggle_autoread)
        self.autoread_btn.pack(side='right', padx=6)
        # Replace the Treeview with a scrollable frame of inline rows (entry + read/write like parameters)
        self.global_fav_rows = {}
        fav_list_container = ttk.Frame(self.global_fav_tab)
        fav_list_container.pack(fill='both', expand=True, padx=4, pady=4)
        canvas = tk.Canvas(fav_list_container)
        vsb = ttk.Scrollbar(fav_list_container, orient='vertical', command=canvas.yview)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        canvas.configure(yscrollcommand=vsb.set)
        inner = ttk.Frame(canvas)
        self._global_fav_inner = inner
        canvas.create_window((0,0), window=inner, anchor='nw')
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        # header row for inline favorites
        hdr = ttk.Frame(inner)
        hdr.pack(fill='x', pady=(0,4))
        ttk.Label(hdr, text='').grid(row=0, column=0, padx=2)
        ttk.Label(hdr, text='Drive', width=8).grid(row=0, column=1)
        ttk.Label(hdr, text='Param', width=12).grid(row=0, column=2)
        ttk.Label(hdr, text='Description', width=40, anchor='w').grid(row=0, column=3)
        ttk.Label(hdr, text='Min', width=8).grid(row=0, column=4)
        ttk.Label(hdr, text='Max', width=8).grid(row=0, column=5)
        ttk.Label(hdr, text='Value', width=12).grid(row=0, column=6)
        ttk.Label(hdr, text='').grid(row=0, column=7)
        # Buttons for operating on favorites are per-row now; keep no global read/write buttons

        # Log area for read results and errors
        try:
            ttk.Label(self.global_fav_tab, text='Log').pack(fill='x', padx=4)
            self.fav_log = ScrolledText(self.global_fav_tab, height=8, state='disabled')
            self.fav_log.pack(fill='both', padx=4, pady=(0,6), expand=False)
        except Exception:
            self.fav_log = None

        self.drive_tabs = {}
        # populate global favorites from saved config at startup
        try:
            self.refresh_global_favorites()
        except Exception:
            pass
        # hide global favorites until connected
        try:
            if not (self.transport.ser and self.transport.ser.is_open):
                try:
                    self.notebook.forget(self.global_fav_tab)
                except Exception:
                    pass
        except Exception:
            pass
        # ensure any existing drive tabs show correct star states
        try:
            favs_map = self.config.get('favorites', {})
            for did, dt in self.drive_tabs.items():
                try:
                    dt.apply_favorite_states(set(favs_map.get(str(did), [])))
                except Exception:
                    pass
        except Exception:
            pass

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
            # hide global favorites when disconnected
            try:
                self.hide_global_favorites()
            except Exception:
                try:
                    self.notebook.forget(self.global_fav_tab)
                except Exception:
                    pass
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
                # show global favorites when connected
                try:
                    self.show_global_favorites()
                except Exception:
                    try:
                        self.notebook.insert(0, self.global_fav_tab, text='Favorites (Global)')
                    except Exception:
                        pass
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
        tab = DriveTab(self.notebook, did, self.transport, self.params, self.root, self, close_callback=lambda d=did: self.remove_drive(d), status_entries_04=self.status_entries_04, status_entries_03=self.status_entries_03, desc_width_chars=self.desc_width_chars)
        self.drive_tabs[did] = tab
        self.notebook.add(tab.frame, text=f'Drive {did}')
        # apply saved favorites visuals for this drive
        try:
            favs = set(self.config.get('favorites', {}).get(str(did), []))
            tab.apply_favorite_states(favs)
        except Exception:
            pass
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
        for did in self.config.get('drives', []):
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

    def show_global_favorites(self):
        """Ensure the global favorites tab is visible (inserted at left-most position)."""
        try:
            # if already present, nothing to do
            try:
                self.notebook.index(self.global_fav_tab)
                return
            except Exception:
                pass
            self.notebook.insert(0, self.global_fav_tab, text='Favorites (Global)')
            try:
                self.refresh_global_favorites()
            except Exception:
                pass
        except Exception:
            pass

    def hide_global_favorites(self):
        """Remove the global favorites tab from the notebook (keeps widget alive)."""
        try:
            try:
                self.notebook.forget(self.global_fav_tab)
            except Exception:
                pass
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

    def toggle_favorite(self, drive_id, param_id):
        """Toggle favorite state for drive_id and param_id. Returns True if now favorite."""
        try:
            favs = self.config.setdefault('favorites', {})
            key = str(drive_id)
            lst = set(map(str, favs.get(key, [])))
            pid = str(param_id)
            # backward compatibility: treat bare numeric as parameter
            if not (pid.startswith('p:') or pid.startswith('s:')):
                pid = 'p:' + pid
            if pid in lst:
                lst.remove(pid)
                favs[key] = list(lst)
                self.save_config()
                return False
            else:
                lst.add(pid)
                favs[key] = list(lst)
                self.save_config()
                return True
        except Exception:
            return False

    def refresh_global_favorites(self):
        # Build or refresh the inline global favorites list (rows with entries + Read/Write)
        try:
            if not getattr(self, '_global_fav_inner', None):
                return
            # clear existing rows (keep header)
            children = list(self._global_fav_inner.winfo_children())
            for child in children[1:]:
                try:
                    child.destroy()
                except Exception:
                    pass
            self.global_fav_rows = {}
            favs = self.config.get('favorites', {})
            for did_str, plist in favs.items():
                try:
                    did = int(did_str)
                except Exception:
                    continue
                for fav in plist:
                    try:
                        desc = ''
                        val = ''
                        label = fav
                        min_val = ''
                        max_val = ''
                        # parameter favorite: 'p:<id>' or legacy numeric
                        if isinstance(fav, str) and fav.startswith('p:'):
                            pid = fav.split(':',1)[1]
                            pinfo = next((p for p in self.params if str(p.get('id','')) == str(pid)), None)
                            desc = pinfo.get('description','') if pinfo else ''
                            if pinfo:
                                min_val = pinfo.get('min','')
                                max_val = pinfo.get('max','')
                            dt = self.drive_tabs.get(did)
                            if dt:
                                try:
                                    w = dt.param_widgets.get(str(pid))
                                    if w:
                                        if isinstance(w, dict):
                                            val = w.get('entry').get()
                                        else:
                                            val = w[0].get()
                                except Exception:
                                    val = ''
                            label = pid
                        elif isinstance(fav, str) and fav.startswith('s:'):
                            parts = fav.split(':')
                            func = parts[1] if len(parts) > 1 else '04'
                            addr = parts[2] if len(parts) > 2 else parts[-1]
                            try:
                                sid = int(addr)
                            except Exception:
                                sid = None
                            if func == '03':
                                entry = next((s for s in self.status_entries_03 if int(s.get('id',0)) == sid), None)
                                tree = getattr(self.drive_tabs.get(did), 'status_tree_03', None) if self.drive_tabs.get(did) else None
                            else:
                                entry = next((s for s in self.status_entries_04 if int(s.get('id',0)) == sid), None)
                                tree = getattr(self.drive_tabs.get(did), 'status_tree_04', None) if self.drive_tabs.get(did) else None
                            if entry:
                                desc = entry.get('description','')
                                val = entry.get('value','')
                            try:
                                iid = f's:{func}:{sid}' if sid is not None else None
                                if tree is not None and iid is not None and tree.exists(iid):
                                    val = tree.set(iid, 'value')
                            except Exception:
                                pass
                            label = f"s{func}:{addr}"
                        else:
                            # legacy numeric pid
                            pid = fav
                            pinfo = next((p for p in self.params if str(p.get('id','')) == str(pid)), None)
                            desc = pinfo.get('description','') if pinfo else ''
                            if pinfo:
                                min_val = pinfo.get('min','')
                                max_val = pinfo.get('max','')
                            dt = self.drive_tabs.get(did)
                            if dt:
                                try:
                                    w = dt.param_widgets.get(str(pid))
                                    if w:
                                        if isinstance(w, dict):
                                            val = w.get('entry').get()
                                        else:
                                            val = w[0].get()
                                except Exception:
                                    val = ''

                        # create inline row
                        row = ttk.Frame(self._global_fav_inner)
                        row.pack(fill='x', pady=2)
                        star_txt = '★'
                        star_btn = ttk.Button(row, text=star_txt, width=2)
                        star_btn.grid(row=0, column=0, padx=2)
                        drv_lbl = ttk.Label(row, text=str(did), width=8)
                        drv_lbl.grid(row=0, column=1)
                        param_lbl = ttk.Label(row, text=label, width=12)
                        param_lbl.grid(row=0, column=2)
                        desc_lbl = ttk.Label(row, text=desc, width=40, anchor='w')
                        desc_lbl.grid(row=0, column=3, sticky='w')
                        min_lbl = ttk.Label(row, text=str(min_val), width=8, anchor='e')
                        min_lbl.grid(row=0, column=4)
                        max_lbl = ttk.Label(row, text=str(max_val), width=8, anchor='e')
                        max_lbl.grid(row=0, column=5)
                        entry = ttk.Entry(row, width=12)
                        try:
                            entry.delete(0, 'end')
                        except Exception:
                            pass
                        entry.insert(0, str(val))
                        entry.grid(row=0, column=6, padx=6)
                        read_btn = ttk.Button(row, text='Read', width=6, command=partial(self._global_read, did, fav, entry))
                        read_btn.grid(row=0, column=7, padx=(4,2))
                        # Do not create a write button for status favorites
                        if isinstance(fav, str) and fav.startswith('s:'):
                            write_btn = ttk.Label(row, text='', width=6)
                            write_btn.grid(row=0, column=8, padx=(2,4))
                        else:
                            write_btn = ttk.Button(row, text='Write', width=6, command=partial(self._global_write, did, fav, entry))
                            write_btn.grid(row=0, column=8, padx=(2,4))

                        def _toggle(d=did, f=fav, btn=star_btn):
                            try:
                                new = self.toggle_favorite(d, f)
                                btn.config(text='★' if new else '☆')
                                # refresh both local and global views
                                try:
                                    self.refresh_global_favorites()
                                except Exception:
                                    pass
                                try:
                                    dt = self.drive_tabs.get(d)
                                    if dt:
                                        dt.refresh_local_favorites()
                                        dt.apply_favorite_states(set(self.config.get('favorites', {}).get(str(d), [])))
                                except Exception:
                                    pass
                            except Exception:
                                pass

                        star_btn.config(command=_toggle)
                        # store widgets for potential updates
                        row_iid = f"{did_str}|{fav}"
                        self.global_fav_rows[row_iid] = {
                            'frame': row,
                            'star': star_btn,
                            'entry': entry,
                            'read': read_btn,
                            'write': write_btn,
                        }
                    except Exception:
                        pass
        except Exception:
            pass

    def read_selected_global_favorite(self):
        """Read the selected row in the global favorites tree."""
        try:
            sel = self.global_fav_tree.selection()
            if not sel:
                return
            iid = sel[0]
            # iid format: '{did}|{fav}'
            try:
                did_str, fav = iid.split('|', 1)
            except Exception:
                vals = self.global_fav_tree.item(iid, 'values')
                if not vals:
                    return
                did_str = vals[1]
                fav = vals[2]
            try:
                did = int(did_str)
            except Exception:
                return
            # if drive tab present, delegate to it
            dt = self.drive_tabs.get(did)
            if dt:
                dt.read_fav(fav)
                try:
                    self.append_log(f"Drive {did} read favorite {fav}")
                except Exception:
                    pass
                return
            # otherwise perform direct transport read
            try:
                if isinstance(fav, str) and fav.startswith('p:'):
                    pid = fav.split(':',1)[1]
                    addr = int(pid)
                    vals = self.transport.read_status(did, addr, 1, func=0x03)
                    if vals:
                        v = vals[0]
                        # update tree value column
                        self.global_fav_tree.set(iid, 'value', str(v))
                        self.append_log(f"Drive {did} Param {pid}: OK -> {v}")
                    else:
                        self.append_log(f"Drive {did} Param {pid}: no response")
                elif isinstance(fav, str) and fav.startswith('s:'):
                    parts = fav.split(':')
                    if len(parts) >= 3:
                        func = int(parts[1])
                        addr = int(parts[2])
                        vals = self.transport.read_status(did, addr, 1, func=func)
                        if vals:
                            v = vals[0]
                            self.global_fav_tree.set(iid, 'value', str(v))
                            self.append_log(f"Drive {did} Status {addr}: OK -> {v}")
                        else:
                            self.append_log(f"Drive {did} Status {addr}: no response")
                else:
                    # legacy numeric
                    pid = str(fav)
                    addr = int(pid)
                    vals = self.transport.read_status(did, addr, 1, func=0x03)
                    if vals:
                        v = vals[0]
                        self.global_fav_tree.set(iid, 'value', str(v))
                        self.append_log(f"Drive {did} Param {pid}: OK -> {v}")
                    else:
                        self.append_log(f"Drive {did} Param {pid}: no response")
            except Exception as e:
                try:
                    self.append_log(f"Read selected failed: {e}")
                except Exception:
                    pass
        except Exception:
            pass

    def _global_read(self, did, fav, entry_widget):
        """Read favorite and place result into provided entry widget."""
        try:
            # delegate to drive tab if exists
            dt = self.drive_tabs.get(did)
            if dt:
                try:
                    dt.read_fav(fav)
                except Exception:
                    pass
                return
            # otherwise direct transport
            if isinstance(fav, str) and fav.startswith('p:'):
                pid = fav.split(':',1)[1]
                addr = int(pid)
                vals = self.transport.read_status(did, addr, 1, func=0x03)
                if vals:
                    v = vals[0]
                    try:
                        entry_widget.delete(0, 'end')
                        entry_widget.insert(0, str(v))
                    except Exception:
                        pass
                    try:
                        self.append_log(f"Drive {did} Param {pid}: OK -> {v}")
                    except Exception:
                        pass
                else:
                    try:
                        self.append_log(f"Drive {did} Param {pid}: no response")
                    except Exception:
                        pass
            elif isinstance(fav, str) and fav.startswith('s:'):
                parts = fav.split(':')
                if len(parts) >= 3:
                    func = int(parts[1])
                    addr = int(parts[2])
                    vals = self.transport.read_status(did, addr, 1, func=func)
                    if vals:
                        v = vals[0]
                        try:
                            entry_widget.delete(0, 'end')
                            entry_widget.insert(0, str(v))
                        except Exception:
                            pass
                        try:
                            self.append_log(f"Drive {did} Status {addr}: OK -> {v}")
                        except Exception:
                            pass
                    else:
                        try:
                            self.append_log(f"Drive {did} Status {addr}: no response")
                        except Exception:
                            pass
        except Exception:
            pass

    def _global_write(self, did, fav, entry_widget):
        """Write the integer value currently in entry_widget to the favorite (if param)."""
        try:
            if isinstance(fav, str) and fav.startswith('s:'):
                messagebox.showwarning('Write', 'Cannot write to status favorites')
                return
            # resolve pid
            if isinstance(fav, str) and fav.startswith('p:'):
                pid = fav.split(':',1)[1]
            else:
                pid = str(fav)
            try:
                vtext = entry_widget.get().strip()
                val = int(vtext)
            except Exception:
                messagebox.showwarning('Invalid', 'Value must be integer')
                return
            addr = int(pid)
            try:
                req = struct.pack('>B B H H', did, 0x06, addr, val)
                crc = compute_crc(req)
                req += struct.pack('<H', crc)
                self.transport.send_and_receive(req)
                try:
                    self.append_log(f"Drive {did} wrote param {pid} = {val}")
                except Exception:
                    pass
            except Exception as e:
                messagebox.showerror('Write error', str(e))
        except Exception:
            pass

    def write_selected_global_favorite(self):
        """Prompt for a value and write it to the selected global favorite (parameters only)."""
        try:
            sel = self.global_fav_tree.selection()
            if not sel:
                return
            iid = sel[0]
            try:
                did_str, fav = iid.split('|', 1)
            except Exception:
                vals = self.global_fav_tree.item(iid, 'values')
                if not vals:
                    return
                did_str = vals[1]
                fav = vals[2]
            try:
                did = int(did_str)
            except Exception:
                return
            if isinstance(fav, str) and fav.startswith('s:'):
                messagebox.showwarning('Write', 'Cannot write to status favorites')
                return
            # resolve pid
            if isinstance(fav, str) and fav.startswith('p:'):
                pid = fav.split(':',1)[1]
            else:
                pid = str(fav)
            v = simpledialog.askstring('Write Parameter', f'Value for parameter {pid} (Drive {did}):')
            if v is None:
                return
            try:
                val = int(v)
            except Exception:
                messagebox.showwarning('Invalid', 'Value must be integer')
                return
            try:
                addr = int(pid)
                req = struct.pack('>B B H H', did, 0x06, addr, val)
                crc = compute_crc(req)
                req += struct.pack('<H', crc)
                self.transport.send_and_receive(req)
                # update UI
                try:
                    self.global_fav_tree.set(iid, 'value', str(val))
                except Exception:
                    pass
                try:
                    self.append_log(f"Drive {did} wrote param {pid} = {val}")
                except Exception:
                    pass
            except Exception as e:
                messagebox.showerror('Write error', str(e))
        except Exception:
            pass

    def _on_global_fav_click(self, event):
        """Handle clicks on the global favorites tree. Toggle favorite when star column clicked."""
        try:
            tree = self.global_fav_tree
            region = tree.identify_region(event.x, event.y)
            if region != 'cell' and region != 'heading':
                # still allow row selection
                return
            col = tree.identify_column(event.x)  # '#1' is star
            row = tree.identify_row(event.y)
            if not row:
                return
            # star column is '#1'
            if col == '#1':
                # iid contains '{did}|{fav}'
                try:
                    did_str, fav = row.split('|', 1)
                except Exception:
                    # try to recover from values
                    vals = tree.item(row, 'values')
                    if not vals:
                        return
                    did_str = vals[1]
                    fav = vals[2]
                try:
                    did = int(did_str)
                except Exception:
                    return
                # call toggle_favorite with drive and fav (fav may already include prefix)
                try:
                    self.toggle_favorite(did, fav)
                except Exception:
                    # older toggle_favorite might expect raw id; try prefixing
                    try:
                        self.toggle_favorite(did, fav if (fav.startswith('p:') or fav.startswith('s:')) else ('p:' + fav))
                    except Exception:
                        pass
                # refresh views
                try:
                    self.refresh_global_favorites()
                except Exception:
                    pass
                try:
                    # refresh drive tab view if present
                    dt = self.drive_tabs.get(did)
                    if dt:
                        dt.refresh_local_favorites()
                        dt.apply_favorite_states(set(self.config.get('favorites', {}).get(str(did), [])))
                        # also update the specific status row tag if applicable
                        try:
                            # fav may be like 'p:123' or 's:04:16' or legacy '123'
                            if fav.startswith('s:'):
                                iid = fav
                                isfav = iid in set(self.config.get('favorites', {}).get(str(did), []))
                                tag = 'fav' if isfav else 'nofav'
                                if hasattr(dt, 'status_tree_04'):
                                    try:
                                        dt.status_tree_04.item(iid, tags=(tag,))
                                    except Exception:
                                        pass
                                if hasattr(dt, 'status_tree_03'):
                                    try:
                                        dt.status_tree_03.item(iid, tags=(tag,))
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

    def save_config(self):
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            messagebox.showwarning('Config save failed', str(e))

    # --- Read favorites / autoread support ---
    def read_all_favorites(self):
        """Launch a background read of all favorites and log results to the embedded log widget."""
        def bg():
            try:
                errors = self._do_read_all_favorites(log_summary=False)
                if errors:
                    self.append_log(f"Read Favorites completed with {len(errors)} errors")
                else:
                    self.append_log("Read Favorites completed: all OK")
            except Exception as e:
                try:
                    self.append_log(f"Read Favorites exception: {e}")
                except Exception:
                    pass
        threading.Thread(target=bg, daemon=True).start()

    def append_log(self, msg: str):
        """Append a timestamped line to the favorites log (if present).

        If the ScrolledText widget isn't available, falls back to printing.
        """
        try:
            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            line = f"[{ts}] {msg}\n"
            if getattr(self, 'fav_log', None):
                try:
                    self.fav_log.config(state='normal')
                    self.fav_log.insert('end', line)
                    self.fav_log.see('end')
                    self.fav_log.config(state='disabled')
                except Exception:
                    # if widget fails, fallback to stdout
                    print(line.strip())
            else:
                print(line.strip())
        except Exception:
            pass

    def _do_read_all_favorites(self, log_summary: bool = True):
        """Synchronously read all favorites and write per-item status to the log.

        Returns a list of error messages (possibly empty).
        """
        errors = []
        favs = self.config.get('favorites', {})
        for did_str, plist in favs.items():
            try:
                did = int(did_str)
            except Exception:
                continue
            for fav in list(plist):
                try:
                    if isinstance(fav, str) and fav.startswith('p:'):
                        pid = fav.split(':', 1)[1]
                        try:
                            addr = int(pid)
                        except Exception:
                            err = f"Drive {did} Param {pid}: invalid id"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        try:
                            vals = self.transport.read_status(did, addr, 1, func=0x03)
                        except Exception as ex:
                            err = f"Drive {did} Param {pid}: {ex}"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        if not vals:
                            err = f"Drive {did} Param {pid}: no response"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        v = vals[0]
                        # update UI if widget exists
                        dt = self.drive_tabs.get(did)
                        if dt:
                            try:
                                w = dt.param_widgets.get(str(pid))
                                if w:
                                    entry = w.get('entry') if isinstance(w, dict) else w[0]
                                    self.root.after(0, lambda e=entry, vv=v: (e.delete(0, 'end'), e.insert(0, str(vv))))
                            except Exception:
                                pass
                        self.append_log(f"Drive {did} Param {pid}: OK -> {v}")
                    elif isinstance(fav, str) and fav.startswith('s:'):
                        parts = fav.split(':')
                        if len(parts) < 3:
                            err = f"Drive {did} Status {fav}: malformed key"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        try:
                            func = int(parts[1])
                            addr = int(parts[2])
                        except Exception:
                            err = f"Drive {did} Status {fav}: invalid func/addr"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        try:
                            vals = self.transport.read_status(did, addr, 1, func=func)
                        except Exception as ex:
                            err = f"Drive {did} Status {addr}: {ex}"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        if not vals:
                            err = f"Drive {did} Status {addr}: no response"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        v = vals[0]
                        # update UI tree if present
                        dt = self.drive_tabs.get(did)
                        iid = f's:{parts[1]}:{addr}'
                        if dt:
                            def _upd(dt=dt, iid=iid, v=v):
                                try:
                                    if hasattr(dt, 'status_tree_04') and dt.status_tree_04.exists(iid):
                                        dt.status_tree_04.set(iid, 'value', str(v))
                                    if hasattr(dt, 'status_tree_03') and dt.status_tree_03.exists(iid):
                                        dt.status_tree_03.set(iid, 'value', str(v))
                                except Exception:
                                    pass
                            self.root.after(0, _upd)
                        self.append_log(f"Drive {did} Status {addr}: OK -> {v}")
                    else:
                        # legacy numeric parameter id
                        pid = str(fav)
                        try:
                            addr = int(pid)
                        except Exception:
                            err = f"Drive {did} Param {pid}: invalid id"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        try:
                            vals = self.transport.read_status(did, addr, 1, func=0x03)
                        except Exception as ex:
                            err = f"Drive {did} Param {pid}: {ex}"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        if not vals:
                            err = f"Drive {did} Param {pid}: no response"
                            errors.append(err)
                            self.append_log(err)
                            continue
                        v = vals[0]
                        dt = self.drive_tabs.get(did)
                        if dt:
                            try:
                                w = dt.param_widgets.get(str(pid))
                                if w:
                                    entry = w.get('entry') if isinstance(w, dict) else w[0]
                                    self.root.after(0, lambda e=entry, vv=v: (e.delete(0, 'end'), e.insert(0, str(vv))))
                            except Exception:
                                pass
                        self.append_log(f"Drive {did} Param {pid}: OK -> {v}")
                except Exception as exc:
                    err = f"Drive {did} fav {fav}: {exc}"
                    errors.append(err)
                    self.append_log(err)
        return errors

    def _toggle_autoread(self):
        try:
            if getattr(self, '_auto_read_job', None):
                # stop autoread
                try:
                    self.root.after_cancel(self._auto_read_job)
                except Exception:
                    pass
                self._auto_read_job = None
                self.autoread_btn.config(text='Auto-Read: Off')
                self.append_log('Auto-Read stopped by user')
                return

            # start autoread with simple backoff on repeated failures
            try:
                interval = float(self.autoread_interval.get())
            except Exception:
                interval = 5.0
            self.autoread_btn.config(text=f'Auto-Read: On ({interval}s)')
            # reset failure tracking
            self._auto_read_failures = 0
            self._auto_read_backoff = 0

            def run_once_and_schedule():
                # runs in a background thread and schedules the next run from main thread
                try:
                    errors = self._do_read_all_favorites(log_summary=False)
                except Exception as e:
                    errors = [str(e)]
                def after_cb():
                    # update failure/backoff counters
                    if errors:
                        self._auto_read_failures = getattr(self, '_auto_read_failures', 0) + 1
                        # after 3 consecutive failures increase backoff (exponential, capped)
                        if self._auto_read_failures >= 3:
                            self._auto_read_backoff = min(getattr(self, '_auto_read_backoff', 0) + 1, 5)
                            next_interval = interval * (2 ** self._auto_read_backoff)
                            self.append_log(f"Auto-Read: {len(errors)} errors; backing off to {next_interval}s (failures={self._auto_read_failures})")
                        else:
                            next_interval = interval
                            self.append_log(f"Auto-Read: {len(errors)} errors (failure count={self._auto_read_failures})")
                    else:
                        # success: reset counters
                        self._auto_read_failures = 0
                        self._auto_read_backoff = 0
                        next_interval = interval
                    # schedule next run
                    try:
                        # create a new thread for the next run
                        def schedule_thread():
                            threading.Thread(target=run_once_and_schedule, daemon=True).start()
                        self._auto_read_job = self.root.after(int(next_interval * 1000), schedule_thread)
                    except Exception:
                        self._auto_read_job = None
                # call after_cb on main thread
                try:
                    self.root.after(0, after_cb)
                except Exception:
                    pass

            # start first run in a thread
            threading.Thread(target=run_once_and_schedule, daemon=True).start()
        except Exception:
            pass
