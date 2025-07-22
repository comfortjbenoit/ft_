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

    # --- GUI omitted here for brevity, see previous code for details ---

    def build_gui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill='both', expand=True)

        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        store_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Stores", menu=store_menu)
        store_menu.add_command(label="Open Table Editor", command=self.open_table_editor)
        store_menu.add_separator()
        store_menu.add_command(label="Manage Store Numbers", command=self.manage_stores)
        area_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Areas", menu=area_menu)
        area_menu.add_command(label="Set Template Import Areas", command=self.set_import_template_areas)
        area_menu.add_command(label="Set Store Sheet Import Areas", command=self.set_store_sheet_areas)
        area_menu.add_separator()
        area_menu.add_command(label="Set Inventory Export Areas", command=self.set_export_inventory_areas)
        area_menu.add_command(label="Set Foil Pan Export Areas", command=self.set_export_foil_areas)
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

    # (All other methods such as import_template, import_store_sheet, etc. unchanged)

    # --- Table Editor (FULL WORKING VERSION) ---
    def open_table_editor(self):
        if "item_names" not in self.template or not self.template["item_names"]:
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
