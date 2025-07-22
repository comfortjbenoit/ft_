import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import json
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy
import shutil
from datetime import datetime

CONFIG_FILE = "config.json"
DATA_FILE = "data.json"

DEFAULT_STORE_COL1 = ["001", "003", "004", "005", "007", "008", "010", "011", "012", "014", "015", "017", "018", "019"]
DEFAULT_STORE_COL2 = ["201", "202", "203", "204", "205", "206", "207", "208", "209", "211", "214", "215", "216", "217"]

def colname2idx(name):
    name = name.upper()
    idx = 0
    for c in name:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1

class AreaDialog(simpledialog.Dialog):
    def __init__(self, parent, title, fields, initial_values=None):
        self.fields = fields
        self.initial_values = initial_values or {}
        self.values = {}
        super().__init__(parent, title)

    def body(self, master):
        self.entries = {}
        for i, field in enumerate(self.fields):
            tk.Label(master, text=field).grid(row=i, column=0, sticky="e")
            val = self.initial_values.get(field, "")
            entry = tk.Entry(master)
            entry.grid(row=i, column=1)
            entry.insert(0, val)
            self.entries[field] = entry
        return list(self.entries.values())[0]

    def apply(self):
        for field, entry in self.entries.items():
            self.values[field] = entry.get().strip()

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Manager")
        self.data = {}
        self.config = {
            "download_path": "",
            "inventory_export_path": "",
            "foil_export_path": "",
            "inventory_template": "",
            "foil_template": "",
            "total_export_template": "",
            "store_col1": DEFAULT_STORE_COL1,
            "store_col2": DEFAULT_STORE_COL2,
            "import_template_areas": {
                "date_cell": "C4",
                "pack_range": "A8:A44",
                "size_range": "B8:B44",
                "desc_range": "C8:C44"
            },
            "store_sheet_areas": {
                "store_cell": "G3",
                "inventory_range": "D8:D44",
                "foil_range": "G8:G11"
            },
            "export_inventory_areas": {
                "date_cell": "A2",
                "item_start_cell": "A5",
                "store_col_start": "B5"
            },
            "export_foil_areas": {
                "date_cell": "C2",
                "data_start_cell": "C5"
            }
        }
        self.template = {}
        self.load_config()
        self.load_data()
        self.build_gui()

    # --- Utility cell/range parsing ---
    def parse_cell(self, ref):
        ref = ref.strip().upper()
        for i, c in enumerate(ref):
            if c.isdigit():
                break
        col = ref[:i]
        row = ref[i:]
        return int(row)-1, colname2idx(col)

    def parse_range(self, ref):
        if ":" in ref:
            a, b = ref.split(":")
            r1, c1 = self.parse_cell(a)
            r2, c2 = self.parse_cell(b)
            return r1, c1, r2, c2
        else:
            r, c = self.parse_cell(ref)
            return r, c, r, c

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    self.config = json.load(f)
                if "store_col1" not in self.config:
                    self.config["store_col1"] = DEFAULT_STORE_COL1
                if "store_col2" not in self.config:
                    self.config["store_col2"] = DEFAULT_STORE_COL2
                if "import_template_areas" not in self.config:
                    self.config["import_template_areas"] = {
                        "date_cell": "C4",
                        "pack_range": "A8:A44",
                        "size_range": "B8:B44",
                        "desc_range": "C8:C44"
                    }
                if "store_sheet_areas" not in self.config:
                    self.config["store_sheet_areas"] = {
                        "store_cell": "G3",
                        "inventory_range": "D8:D44",
                        "foil_range": "G8:G11"
                    }
                if "export_inventory_areas" not in self.config:
                    self.config["export_inventory_areas"] = {
                        "date_cell": "A2",
                        "item_start_cell": "A5",
                        "store_col_start": "B5"
                    }
                if "export_foil_areas" not in self.config:
                    self.config["export_foil_areas"] = {
                        "date_cell": "C2",
                        "data_start_cell": "C5"
                    }
            except Exception:
                self._set_default_config()
        else:
            self._set_default_config()

    def _set_default_config(self):
        self.config = {
            "download_path": "",
            "inventory_export_path": "",
            "foil_export_path": "",
            "inventory_template": "",
            "foil_template": "",
            "total_export_template": "",
            "store_col1": DEFAULT_STORE_COL1,
            "store_col2": DEFAULT_STORE_COL2,
            "import_template_areas": {
                "date_cell": "C4",
                "pack_range": "A8:A44",
                "size_range": "B8:B44",
                "desc_range": "C8:C44"
            },
            "store_sheet_areas": {
                "store_cell": "G3",
                "inventory_range": "D8:D44",
                "foil_range": "G8:G11"
            },
            "export_inventory_areas": {
                "date_cell": "A2",
                "item_start_cell": "A5",
                "store_col_start": "B5"
            },
            "export_foil_areas": {
                "date_cell": "C2",
                "data_start_cell": "C5"
            }
        }

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(self.config, f, indent=2)

    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r') as f:
                    self.data = json.load(f)
            except Exception:
                self.data = {}

    def save_data(self):
        with open(DATA_FILE, 'w') as f:
            json.dump(self.data, f, indent=2)

    def get_all_stores(self):
        return self.config.get("store_col1", []) + self.config.get("store_col2", [])

    def build_gui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill='both', expand=True)

        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Stores menu
        store_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Stores", menu=store_menu)
        store_menu.add_command(label="Open Table Editor", command=self.open_table_editor)
        store_menu.add_separator()
        store_menu.add_command(label="Manage Store Numbers", command=self.manage_stores)

        # Areas menu
        area_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Areas", menu=area_menu)
        area_menu.add_command(label="Set Template Import Areas", command=self.set_import_template_areas)
        area_menu.add_command(label="Set Store Sheet Import Areas", command=self.set_store_sheet_areas)
        area_menu.add_separator()
        area_menu.add_command(label="Set Inventory Export Areas", command=self.set_export_inventory_areas)
        area_menu.add_command(label="Set Foil Pan Export Areas", command=self.set_export_foil_areas)

        # Settings menu for all path settings
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Set Downloads Folder", command=self.set_download_path)
        settings_menu.add_command(label="Set Final Inventory Folder", command=self.set_inventory_export_path)
        settings_menu.add_command(label="Set Foil Pan Folder", command=self.set_foil_export_path)
        settings_menu.add_separator()
        settings_menu.add_command(label="Set This Week's Inventory Template", command=self.set_inventory_template_path)
        settings_menu.add_command(label="Set Foil Pan Template", command=self.set_foil_template_path)
        settings_menu.add_command(label="Set Final Inventory Template", command=self.set_total_export_template_path)

        btn_import = ttk.Button(frame, text="Import Store Sheet", command=self.import_store_sheet)
        btn_import.grid(row=0, column=0, padx=5, pady=5)
        btn_import_template = ttk.Button(frame, text="Import Item List from Template", command=self.import_template)
        btn_import_template.grid(row=0, column=1, padx=5, pady=5)
        btn_export_combo = ttk.Button(frame, text="Export Final Totals", command=self.export_combo)
        btn_export_combo.grid(row=0, column=2, padx=5, pady=5)
        btn_export_json = ttk.Button(frame, text="Export Data", command=self.export_json_data)
        btn_export_json.grid(row=0, column=3, padx=5, pady=5)
        btn_import_json = ttk.Button(frame, text="Import Data", command=self.import_json_data)
        btn_import_json.grid(row=0, column=4, padx=5, pady=5)

        store_status_frame = ttk.LabelFrame(frame, text="Store Upload Status")
        store_status_frame.grid(row=2, column=0, columnspan=5, pady=10, sticky="we")
        for col in range(2):
            store_status_frame.grid_columnconfigure(col, weight=1)
        self.store_labels_col1 = []
        self.store_labels_col2 = []
        label_font = ("Arial", 10, "bold")
        for row, store in enumerate(self.config["store_col1"]):
            lbl = tk.Label(store_status_frame, text="", anchor='center', width=8, font=label_font, justify='center')
            lbl.grid(row=row, column=0, sticky="we", padx=2, pady=1)
            self.store_labels_col1.append(lbl)
        for row, store in enumerate(self.config["store_col2"]):
            lbl = tk.Label(store_status_frame, text="", anchor='center', width=8, font=label_font, justify='center')
            lbl.grid(row=row, column=1, sticky="we", padx=2, pady=1)
            self.store_labels_col2.append(lbl)

        imported_frame = ttk.LabelFrame(frame, text="Imported Stores Progress")
        imported_frame.grid(row=4, column=0, columnspan=5, pady=10, sticky="we")
        self.imported_progress = ttk.Progressbar(imported_frame, orient="horizontal", length=700, mode="determinate")
        self.imported_progress.pack(fill='x', padx=10, pady=8)
        self.update_store_status_display()

        self.status = ttk.Label(frame, text="Status: Ready")
        self.status.grid(row=5, column=0, columnspan=5, pady=10)

        preferred_width = 800
        preferred_height = 600
        self.root.minsize(600, 400)
        self.root.geometry(f"{preferred_width}x{preferred_height}")
        self.root.resizable(True, True)
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (preferred_width // 2)
        y = (screen_height // 2) - (preferred_height // 2)
        self.root.geometry(f"+{x}+{y}")

    def update_store_status_display(self):
        check, cross = "\u2714", "\u2716"
        stores1 = self.config.get("store_col1", [])
        stores2 = self.config.get("store_col2", [])
        for i, store in enumerate(stores1):
            uploaded = store in self.data
            if i < len(self.store_labels_col1):
                if uploaded:
                    self.store_labels_col1[i]['text'] = f"{int(store):3} {check}"
                    self.store_labels_col1[i]['fg'] = "#1ca41c"
                else:
                    self.store_labels_col1[i]['text'] = f"{int(store):3} {cross}"
                    self.store_labels_col1[i]['fg'] = "black"
        for i, store in enumerate(stores2):
            uploaded = store in self.data
            if i < len(self.store_labels_col2):
                if uploaded:
                    self.store_labels_col2[i]['text'] = f"{int(store):3} {check}"
                    self.store_labels_col2[i]['fg'] = "#1ca41c"
                else:
                    self.store_labels_col2[i]['text'] = f"{int(store):3} {cross}"
                    self.store_labels_col2[i]['fg'] = "black"
        self.update_imported_stores_progress()

    def update_imported_stores_progress(self):
        total_stores = len(self.get_all_stores())
        imported = len([s for s in self.get_all_stores() if s in self.data])
        self.imported_progress["maximum"] = total_stores
        self.imported_progress["value"] = imported

    # --- Settings/Area dialogs ---
    def set_download_path(self):
        path = filedialog.askdirectory()
        if path:
            self.config["download_path"] = path
            self.save_config()
    def set_inventory_export_path(self):
        path = filedialog.askdirectory()
        if path:
            self.config["inventory_export_path"] = path
            self.save_config()
    def set_foil_export_path(self):
        path = filedialog.askdirectory()
        if path:
            self.config["foil_export_path"] = path
            self.save_config()
    def set_inventory_template_path(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            self.config["inventory_template"] = path
            self.save_config()
    def set_foil_template_path(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            self.config["foil_template"] = path
            self.save_config()
    def set_total_export_template_path(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            self.config["total_export_template"] = path
            self.save_config()
    def set_import_template_areas(self):
        fields = ["date_cell", "pack_range", "size_range", "desc_range"]
        initial = self.config.get("import_template_areas", {})
        dlg = AreaDialog(self.root, "Set Template Import Areas", fields, initial)
        if dlg.values:
            self.config["import_template_areas"] = dlg.values
            self.save_config()
    def set_store_sheet_areas(self):
        fields = ["store_cell", "inventory_range", "foil_range"]
        initial = self.config.get("store_sheet_areas", {})
        dlg = AreaDialog(self.root, "Set Store Sheet Import Areas", fields, initial)
        if dlg.values:
            self.config["store_sheet_areas"] = dlg.values
            self.save_config()
    def set_export_inventory_areas(self):
        fields = ["date_cell", "item_start_cell", "store_col_start"]
        initial = self.config.get("export_inventory_areas", {})
        dlg = AreaDialog(self.root, "Set Inventory Export Areas", fields, initial)
        if dlg.values:
            self.config["export_inventory_areas"] = dlg.values
            self.save_config()
    def set_export_foil_areas(self):
        fields = ["date_cell", "data_start_cell"]
        initial = self.config.get("export_foil_areas", {})
        dlg = AreaDialog(self.root, "Set Foil Pan Export Areas", fields, initial)
        if dlg.values:
            self.config["export_foil_areas"] = dlg.values
            self.save_config()

    # --- Import functions (template/store sheets) ---

    def import_template(self):
        path = self.config.get("inventory_template")
        if not path or not os.path.exists(path):
            path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
            if not path:
                return
            self.config["inventory_template"] = path
            self.save_config()
        try:
            book = xlrd.open_workbook(path)
            sheet = book.sheet_by_index(0)
            areas = self.config.get("import_template_areas", {})
            date_val = ""
            try:
                row, col = self.parse_cell(areas.get("date_cell", "C4"))
                date_val = str(sheet.cell_value(row, col)).strip()
            except Exception:
                pass
            self.template["date"] = date_val
            pack_r1, pack_c1, pack_r2, pack_c2 = self.parse_range(areas.get("pack_range", "A8:A44"))
            size_r1, size_c1, size_r2, size_c2 = self.parse_range(areas.get("size_range", "B8:B44"))
            desc_r1, desc_c1, desc_r2, desc_c2 = self.parse_range(areas.get("desc_range", "C8:C44"))
            n_items = max(pack_r2-pack_r1+1, size_r2-size_r1+1, desc_r2-desc_r1+1)
            items = []
            item_names = []
            for i in range(n_items):
                try: pack = str(sheet.cell_value(pack_r1+i, pack_c1)).strip()
                except Exception: pack = ""
                try: size = str(sheet.cell_value(size_r1+i, size_c1)).strip()
                except Exception: size = ""
                try: desc = str(sheet.cell_value(desc_r1+i, desc_c1)).strip()
                except Exception: desc = ""
                items.append({'case_qty': pack, 'size': size, 'description': desc})
                display_name = f"{desc}, {size}, {pack}".strip(", ")
                item_names.append(display_name)
            self.template["items"] = items
            self.template["item_names"] = item_names
            self.template["template_path"] = path
            self.status.config(text=f"Template imported: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import template: {e}")

    def load_excel_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        areas = self.config.get("store_sheet_areas", {})
        try:
            if ext == ".xls":
                book = xlrd.open_workbook(path)
                sheet = book.sheet_by_index(0)
                store_row, store_col = self.parse_cell(areas.get("store_cell", "G3"))
                store_cell = sheet.cell_value(store_row, store_col)
                if isinstance(store_cell, float):
                    store = f"{int(store_cell):03}"
                else:
                    store = str(store_cell).zfill(3)
                ir1, ic1, ir2, ic2 = self.parse_range(areas.get("inventory_range", "D8:D44"))
                inventory = []
                for i in range(ir2-ir1+1):
                    try:
                        inventory.append(sheet.cell_value(ir1+i, ic1))
                    except Exception:
                        inventory.append("")
                fr1, fc1, fr2, fc2 = self.parse_range(areas.get("foil_range", "G8:G11"))
                foil = []
                for i in range(fr2-fr1+1):
                    try:
                        foil.append(sheet.cell_value(fr1+i, fc1))
                    except Exception:
                        foil.append("")
            else:
                raise ValueError("Unsupported file format. Only .xls files are supported.")
            return store, inventory, foil
        except Exception as e:
            raise Exception(f"Import error: {e}")

    def import_store_sheet(self):
        paths = filedialog.askopenfilenames(
            initialdir=self.config.get("download_path", ""),
            filetypes=[("Excel 97-2003", "*.xls")],
            title="Select Store Sheet(s) to Import"
        )
        if not paths:
            return

        imported_stores = []
        for path in paths:
            try:
                store, inventory, foil = self.load_excel_file(path)
                store = f"{int(float(store)):03}"
                self.data[store] = {"inventory": inventory, "foil": foil}
                imported_stores.append(store)
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to read file '{os.path.basename(path)}': {e}")

        if imported_stores:
            self.save_data()
            self.status.config(text=f"Imported stores: {', '.join(imported_stores)}")
            self.update_store_status_display()

    # --- Export functions ---

    def export_inventory_to_template(self):
        template_path = self.config.get("total_export_template", "")
        export_folder = self.config.get("inventory_export_path", "")
        areas = self.config.get("export_inventory_areas", {})
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Error", "No inventory template set or file does not exist.")
            return
        if not export_folder or not os.path.exists(export_folder):
            messagebox.showerror("Error", "No inventory export folder set or folder does not exist.")
            return

        stores = self.get_all_stores()
        item_names = self.template.get("item_names", [])
        date_str = self.template.get("date", datetime.today().strftime("%m-%d-%Y"))
        try:
            if date_str and isinstance(date_str, (int, float)):
                date_str = xlrd.xldate.xldate_as_datetime(float(date_str), 0).strftime("%m-%d-%Y")
            elif date_str and date_str.isdigit():
                date_str = xlrd.xldate.xldate_as_datetime(float(date_str), 0).strftime("%m-%d-%Y")
        except Exception:
            pass

        out_path = os.path.join(export_folder, f"Final Inventory {date_str}.xls")
        shutil.copy(template_path, out_path)
        rb = xlrd.open_workbook(out_path, formatting_info=True)
        wb = xl_copy(rb)
        ws = wb.get_sheet(0)
        rb_sheet = rb.sheet_by_index(0)

        # Write date in cell A2
        date_row, date_col = self.parse_cell(areas.get("date_cell", "A2"))
        ws.write(date_row, date_col, date_str)

        # Write store numbers in A5:A34
        item_row, item_col = self.parse_cell(areas.get("item_start_cell", "A5"))
        for i, store in enumerate(stores):
            ws.write(item_row + i, item_col, store)

        # Write inventory data for each store in b5:ae5, b6:ae6, etc.
        store_col_row, store_col_col = self.parse_cell(areas.get("store_col_start", "B5"))
        for store_idx, store in enumerate(stores):
            inv = self.data.get(store, {}).get("inventory", [""] * len(item_names))
            for i, val in enumerate(inv):
                ws.write(store_col_row + store_idx, store_col_col + i, val)

        # Format cells b5:ae34 to have thin borders
        thin_border = xlwt.Borders()
        thin_border.left = thin_border.right = thin_border.top = thin_border.bottom = xlwt.Borders.THIN
        thin_xf = xlwt.XFStyle()
        thin_xf.borders = thin_border
        for r in range(store_col_row, store_col_row + len(stores)):
            for c in range(store_col_col, store_col_col + len(item_names)):
                val = rb_sheet.cell_value(r, c) if c < rb_sheet.ncols and r < rb_sheet.nrows else ""
                ws.write(r, c, val, thin_xf)

        wb.save(out_path)
        self.status.config(text=f"Inventory exported to {out_path}")
        messagebox.showinfo("Export Complete", f"Inventory exported to {out_path}")

    def export_foil_to_template(self):
        template_path = self.config.get("foil_template", "")
        export_folder = self.config.get("foil_export_path", "")
        areas = self.config.get("export_foil_areas", {})
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Error", "No foil pan template set or file does not exist.")
            return
        if not export_folder or not os.path.exists(export_folder):
            messagebox.showerror("Error", "No foil pan export folder set or folder does not exist.")
            return

        stores = self.get_all_stores()
        date_str = self.template.get("date", datetime.today().strftime("%m-%d-%Y"))
        try:
            if date_str and isinstance(date_str, (int, float)):
                date_str = xlrd.xldate.xldate_as_datetime(float(date_str), 0).strftime("%m-%d-%Y")
            elif date_str and date_str.isdigit():
                date_str = xlrd.xldate.xldate_as_datetime(float(date_str), 0).strftime("%m-%d-%Y")
        except Exception:
            pass
        out_path = os.path.join(export_folder, f"Foil Pan Order {date_str}.xls")
        shutil.copy(template_path, out_path)
        rb = xlrd.open_workbook(out_path, formatting_info=True)
        wb = xl_copy(rb)
        ws = wb.get_sheet(0)
        rb_sheet = rb.sheet_by_index(0)

        # Write date in cell C2
        date_row, date_col = self.parse_cell(areas.get("date_cell", "C2"))
        ws.write(date_row, date_col, date_str)

        # Write foil pans: store 001 in c5:f5, 003 in c6:f6, etc.
        data_row, data_col = self.parse_cell(areas.get("data_start_cell", "C5"))
        for i, store in enumerate(stores):
            foil = self.data.get(store, {}).get("foil", [""] * 4)
            for j in range(4):
                ws.write(data_row + i, data_col + j, foil[j] if len(foil) > j else "")

        # Format c5:f32 with Arial 14pt and borders
        arial14 = xlwt.Font()
        arial14.name = 'Arial'
        arial14.height = 14 * 20
        thin_border = xlwt.Borders()
        thin_border.left = thin_border.right = thin_border.top = thin_border.bottom = xlwt.Borders.THIN
        style = xlwt.XFStyle()
        style.font = arial14
        style.borders = thin_border
        for r in range(data_row, data_row + len(stores)):
            for c in range(data_col, data_col + 4):
                val = rb_sheet.cell_value(r, c) if c < rb_sheet.ncols and r < rb_sheet.nrows else ""
                ws.write(r, c, val, style)

        wb.save(out_path)
        self.status.config(text=f"Foil pan order exported to {out_path}")
        messagebox.showinfo("Export Complete", f"Foil pan order exported to {out_path}")

    def export_combo(self):
        self.export_inventory_to_template()
        self.export_foil_to_template()

    # --- Table Editor and store management code here (unchanged from previous versions) ---

    def open_table_editor(self):
        # (table editor code unchanged from previous blocks...)
        # ... see previous full code for editor and store management implementations ...
        ...

    def manage_stores(self):
        # (store management code unchanged from previous blocks...)
        # ... see previous full code for editor and store management implementations ...
        ...

    def export_json_data(self):
        export_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            initialfile="inventory_data_export.json"
        )
        if export_path:
            try:
                with open(export_path, 'w') as f:
                    json.dump({
                        "data": self.data,
                        "config": self.config,
                        "template": self.template
                    }, f, indent=2)
                self.status.config(text=f"Data exported to {export_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export data: {e}")

    def import_json_data(self):
        import_path = filedialog.askopenfilename(
            filetypes=[("JSON Files", "*.json")]
        )
        if import_path:
            try:
                with open(import_path, 'r') as f:
                    imported = json.load(f)
                self.data = imported.get("data", {})
                self.config = imported.get("config", self.config)
                self.template = imported.get("template", {})
                self.save_data()
                self.save_config()
                self.update_store_status_display()
                self.status.config(text=f"Data imported from {os.path.basename(import_path)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import data: {e}")

def main():
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
