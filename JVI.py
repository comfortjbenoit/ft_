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
                "date_cell": "A2",
                "store_start_cell": "B5"
            }
        }
        self.template = {}
        self.load_config()
        self.load_data()
        self.build_gui()

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
                        "date_cell": "A2",
                        "store_start_cell": "B5"
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
                "date_cell": "A2",
                "store_start_cell": "B5"
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

    def fix_data_store_keys(self):
        new_data = {}
        for k, v in self.data.items():
            try:
                newk = f"{int(float(k)):03}"
            except:
                newk = str(k).zfill(3)
            new_data[newk] = v
        self.data = new_data

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

        # Data menu for clear/reset function
        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Data", menu=data_menu)
        data_menu.add_command(label="Clear All Data", command=self.clear_all_data)

        # Main buttons row
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

        # Store status display
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

        # Progress bar for imported stores
        imported_frame = ttk.LabelFrame(frame, text="Imported Stores Progress")
        imported_frame.grid(row=4, column=0, columnspan=5, pady=10, sticky="we")
        self.imported_progress = ttk.Progressbar(imported_frame, orient="horizontal", length=700, mode="determinate")
        self.imported_progress.pack(fill='x', padx=10, pady=8)

        self.update_store_status_display()  # <-- Now after imported_progress is created

        self.status = ttk.Label(frame, text="Status: Ready")
        self.status.grid(row=5, column=0, columnspan=5, pady=10)

        # Reasonable window size and centering
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

    # Settings menu methods
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

    # Area dialog menu methods
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
        fields = ["date_cell", "store_start_cell"]
        initial = self.config.get("export_foil_areas", {})
        dlg = AreaDialog(self.root, "Set Foil Pan Export Areas", fields, initial)
        if dlg.values:
            self.config["export_foil_areas"] = dlg.values
            self.save_config()

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
            self.fix_data_store_keys()
            self.update_store_status_display()
            self.status.config(text=f"Imported stores: {', '.join(imported_stores)}")

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
                self.fix_data_store_keys()
                self.config = imported.get("config", self.config)
                self.template = imported.get("template", {})
                self.save_data()
                self.save_config()
                self.update_store_status_display()
                self.status.config(text=f"Data imported from {os.path.basename(import_path)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import data: {e}")

    def _copy_only_values_to_sheet(self, ws, rb_sheet, rowcolvals, date_cells_formats={}):
        # rowcolvals: list of (row, col, value), for date cells, use date_cells_formats dict to format as string
        for (row, col, value) in rowcolvals:
            fmt = None
            if (row, col) in date_cells_formats:
                fmt = date_cells_formats[(row, col)]
            if fmt:
                # Format date string for Excel (write as string)
                if isinstance(value, (int, float)):
                    value = xlrd.xldate.xldate_as_datetime(value, 0).strftime(fmt)
                elif isinstance(value, str):
                    try:
                        # Try parsing as xldate float
                        val_float = float(value)
                        value = xlrd.xldate.xldate_as_datetime(val_float, 0).strftime(fmt)
                    except Exception:
                        pass
            ws.write(row, col, value)

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
        # Always export as MM-DD-YYYY string
        date_str = self.template.get("date", datetime.today().strftime("%m-%d-%Y"))
        # If looks like Excel xldate float, format as string
        try:
            if date_str and isinstance(date_str, (int, float)):
                date_str = xlrd.xldate.xldate_as_datetime(float(date_str), 0).strftime("%m-%d-%Y")
            elif date_str and date_str.isdigit():
                # Could be excel date as string
                date_str = xlrd.xldate.xldate_as_datetime(float(date_str), 0).strftime("%m-%d-%Y")
        except Exception:
            pass

        out_path = os.path.join(export_folder, f"Final Inventory {date_str}.xls")
        shutil.copy(template_path, out_path)
        rb = xlrd.open_workbook(out_path, formatting_info=True)
        wb = xl_copy(rb)
        ws = wb.get_sheet(0)
        rb_sheet = rb.sheet_by_index(0)

        # Prepare values to write (row, col, value)
        rowcolvals = []
        # Date cell
        date_row, date_col = self.parse_cell(areas.get("date_cell", "A2"))
        rowcolvals.append((date_row, date_col, date_str))
        date_cells_formats = {(date_row, date_col): "%m-%d-%Y"}

        # Items
        item_row, item_col = self.parse_cell(areas.get("item_start_cell", "A5"))
        for i, item_display in enumerate(item_names):
            rowcolvals.append((item_row + i, item_col, item_display))

        # Store inventories
        store_col_row, store_col_col = self.parse_cell(areas.get("store_col_start", "B5"))
        for store_idx, store in enumerate(stores):
            col = store_col_col + store_idx
            inv = self.data.get(store, {}).get("inventory", [""] * len(item_names))
            for row_idx in range(len(item_names)):
                rowcolvals.append((item_row + row_idx, col, inv[row_idx] if row_idx < len(inv) else ""))

        # Only overwrite values, not formatting
        self._copy_only_values_to_sheet(ws, rb_sheet, rowcolvals, date_cells_formats=date_cells_formats)
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

        # Prepare values to write (row, col, value)
        rowcolvals = []
        date_row, date_col = self.parse_cell(areas.get("date_cell", "A2"))
        rowcolvals.append((date_row, date_col, date_str))
        date_cells_formats = {(date_row, date_col): "%m-%d-%Y"}

        # Write store and foil data starting from user-defined cell
        store_row, store_col = self.parse_cell(areas.get("store_start_cell", "B5"))
        for i, store in enumerate(stores):
            rowcolvals.append((store_row + i, store_col, store))
            foil = self.data.get(store, {}).get("foil", [""] * 4)
            for j in range(4):
                rowcolvals.append((store_row + i, store_col + 1 + j, foil[j] if len(foil) > j else ""))

        self._copy_only_values_to_sheet(ws, rb_sheet, rowcolvals, date_cells_formats=date_cells_formats)
        wb.save(out_path)
        self.status.config(text=f"Foil pan order exported to {out_path}")
        messagebox.showinfo("Export Complete", f"Foil pan order exported to {out_path}")

    def export_combo(self):
        self.export_inventory_to_template()
        self.export_foil_to_template()
        total_template = self.config.get("total_export_template", "")
        if total_template and os.path.exists(total_template):
            self.status.config(text=self.status.cget("text") + " | Total Export Template set.")

    def open_table_editor(self):
        # Load template fresh so it always shows current data
        if "item_names" not in self.template or not self.template["item_names"]:
            # Try to load from template file if not already loaded
            self.import_template()
            if "item_names" not in self.template or not self.template["item_names"]:
                messagebox.showerror("No Template", "You must import a template before editing inventory in table view.")
                return

        item_names = self.template["item_names"]
        n_items = len(item_names)
        if not item_names:
            messagebox.showerror("No Item Names", "No item names found in template.")
            return

        stores = self.get_all_stores()
        # If there is no inventory for any store, initialize blank
        for store in stores:
            if store not in self.data:
                self.data[store] = {"inventory": [""] * n_items, "foil": [""] * 4}
            else:
                inv = self.data[store].get("inventory", [])
                if len(inv) < n_items:
                    inv = inv + [""] * (n_items - len(inv))
                self.data[store]["inventory"] = inv[:n_items]

        editor = tk.Toplevel(self.root)
        editor.title("Inventory Data Table Editor")
        preferred_width = min(1400, max(900, 300 + len(stores) * 45))
        preferred_height = min(800, max(400, 40 + n_items * 22))
        editor.geometry(f"{preferred_width}x{preferred_height}")
        editor.minsize(600, 320)
        editor.resizable(True, True)

        editor_frame = ttk.Frame(editor, padding=10)
        editor_frame.pack(fill='both', expand=True)

        font_size = 8
        table_font = ("Arial", font_size)
        heading_font = ("Arial", font_size + 1, "bold")

        style = ttk.Style(editor)
        style.configure("Treeview", font=table_font, rowheight=19)
        style.configure("Treeview.Heading", font=heading_font)
        style.layout("Treeview.Cell", [
            ('Treeitem.padding', {'sticky': 'nswe', 'children': [
                ('Treeitem.image', {'side': 'left', 'sticky': ''}),
                ('Treeitem.text', {'side': 'left', 'sticky': '', 'padding': [1, 0, 1, 0]}),
            ]})
        ])

        xscroll = tk.Scrollbar(editor_frame, orient="horizontal")
        xscroll.pack(side="bottom", fill="x")
        yscroll = tk.Scrollbar(editor_frame, orient="vertical")
        yscroll.pack(side="right", fill="y")

        columns = ["Item"] + stores
        tree = ttk.Treeview(
            editor_frame,
            columns=columns,
            show="headings",
            height=min(len(item_names), 25),
            xscrollcommand=xscroll.set,
            yscrollcommand=yscroll.set,
            style="Treeview"
        )
        xscroll.config(command=tree.xview)
        yscroll.config(command=tree.yview)

        for col in columns:
            tree.heading(col, text=col)
            if col == "Item":
                tree.column(col, width=300, anchor='w', minwidth=120, stretch=True)
            else:
                tree.column(col, width=44, anchor='center', minwidth=25, stretch=True)
        tree.pack(side="left", fill="both", expand=True)

        # Clear all rows before inserting
        for rowid in tree.get_children():
            tree.delete(rowid)

        for idx, item_name in enumerate(item_names):
            row_vals = [item_name]
            for store in stores:
                inv = self.data.get(store, {}).get("inventory", [""] * n_items)
                val = inv[idx] if idx < len(inv) else ""
                row_vals.append(str(val))
            tree.insert("", "end", values=row_vals, tags=(f"row_{idx}",))

        def on_double_click(event):
            region = tree.identify("region", event.x, event.y)
            if region != "cell":
                return
            rowid = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            col_index = int(col.replace("#", "")) - 1
            if rowid == "" or col_index == 0:
                return
            x, y, width, height = tree.bbox(rowid, col)
            value = tree.set(rowid, columns[col_index])
            entry = tk.Entry(tree, width=8, font=table_font)
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, value)
            entry.focus()

            def on_entry_confirm(event=None):
                newval = entry.get()
                tree.set(rowid, columns[col_index], newval)
                entry.destroy()

            entry.bind("<Return>", on_entry_confirm)
            entry.bind("<FocusOut>", lambda e: entry.destroy())

        tree.bind("<Double-1>", on_double_click)

        def save_table_edits():
            for idx, rowid in enumerate(tree.get_children()):
                values = tree.item(rowid)["values"]
                for col_idx, store in enumerate(stores, start=1):
                    val = values[col_idx]
                    inv = self.data.get(store, {}).get("inventory", [""] * n_items)
                    while len(inv) < n_items:
                        inv.append("")
                    inv[idx] = val
                    foil = self.data.get(store, {}).get("foil", [""]*4)
                    self.data[store] = {"inventory": inv, "foil": foil}
            self.save_data()
            self.update_store_status_display()
            messagebox.showinfo("Saved", "All table edits have been saved.")
            editor.lift()
            self.status.config(text="All table edits saved.")

        savebtn = ttk.Button(editor_frame, text="Save All Changes", command=save_table_edits)
        savebtn.pack(side="bottom", pady=5)

        def resize_columns(event=None):
            width = editor.winfo_width()
            n_store_cols = max(1, len(stores))
            item_col_width = min(400, max(150, int(width * 0.35)))
            store_col_width = max(25, int((width - item_col_width - 60) / n_store_cols))
            tree.column("Item", width=item_col_width)
            for col in stores:
                tree.column(col, width=store_col_width)

        editor.bind('<Configure>', resize_columns)
        resize_columns()

    def manage_stores(self):
        def refresh_lists():
            col1_list.delete(0, tk.END)
            for s in self.config["store_col1"]:
                col1_list.insert(tk.END, s)
            col2_list.delete(0, tk.END)
            for s in self.config["store_col2"]:
                col2_list.insert(tk.END, s)

        win = tk.Toplevel(self.root)
        win.title("Manage Store Numbers")

        frame = ttk.Frame(win, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Store Column 1").grid(row=0, column=0, padx=5)
        ttk.Label(frame, text="Store Column 2").grid(row=0, column=2, padx=5)

        col1_list = tk.Listbox(frame, selectmode=tk.SINGLE, font=("Arial", 9))
        col1_list.grid(row=1, column=0, padx=5)
        col2_list = tk.Listbox(frame, selectmode=tk.SINGLE, font=("Arial", 9))
        col2_list.grid(row=1, column=2, padx=5)

        refresh_lists()

        def add_store(col):
            s = simpledialog.askstring("Add Store Number", "Enter store number (3 digits):", parent=win)
            if s:
                s = str(s).zfill(3)
                if col == 1:
                    if s not in self.config["store_col1"]:
                        self.config["store_col1"].append(s)
                else:
                    if s not in self.config["store_col2"]:
                        self.config["store_col2"].append(s)
                self.save_config()
                refresh_lists()
                self.update_store_status_display()

        def remove_store(col):
            if col == 1:
                sel = col1_list.curselection()
                if sel:
                    idx = sel[0]
                    del self.config["store_col1"][idx]
            else:
                sel = col2_list.curselection()
                if sel:
                    idx = sel[0]
                    del self.config["store_col2"][idx]
            self.save_config()
            refresh_lists()
            self.update_store_status_display()

        ttk.Button(frame, text="Add to Col 1", command=lambda: add_store(1)).grid(row=2, column=0, pady=3)
        ttk.Button(frame, text="Add to Col 2", command=lambda: add_store(2)).grid(row=2, column=2, pady=3)
        ttk.Button(frame, text="Remove from Col 1", command=lambda: remove_store(1)).grid(row=3, column=0, pady=3)
        ttk.Button(frame, text="Remove from Col 2", command=lambda: remove_store(2)).grid(row=3, column=2, pady=3)

        ttk.Button(frame, text="Close", command=win.destroy).grid(row=4, column=0, columnspan=3, pady=8)

    def clear_all_data(self):
        if messagebox.askyesno("Confirm", "Are you sure you want to CLEAR ALL imported item data and store data? This cannot be undone."):
            self.data = {}
            self.template = {}
            self.save_data()
            self.update_store_status_display()
            self.status.config(text="All data cleared.")

def main():
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
