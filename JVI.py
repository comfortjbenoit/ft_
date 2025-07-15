import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import xlrd
import xlwt
from datetime import datetime

CONFIG_FILE = "config.json"
DATA_FILE = "data.json"

STORE_COL1 = ["001", "003", "004", "005", "007", "008", "010", "011", "012", "014", "015", "017", "018", "019"]
STORE_COL2 = ["201", "202", "203", "204", "205", "206", "207", "208", "209", "211", "214", "215", "216", "217"]
ALL_STORES = STORE_COL1 + STORE_COL2

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Manager")
        self.data = {}
        self.config = {"download_path": "", "export_path": ""}
        self.template = {}

        self.load_config()
        self.load_data()
        self.build_gui()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    self.config = json.load(f)
            except Exception:
                self.config = {"download_path": "", "export_path": ""}

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
        """Make sure store keys are always zero-padded strings."""
        new_data = {}
        for k, v in self.data.items():
            try:
                newk = f"{int(float(k)):03}"
            except:
                newk = str(k).zfill(3)
            new_data[newk] = v
        self.data = new_data

    def build_gui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill='both', expand=True)

        # Menu for table editor
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        store_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Stores", menu=store_menu)
        store_menu.add_command(label="Open Table Editor", command=self.open_table_editor)

        btn_template = ttk.Button(frame, text="Import Template", command=self.import_template)
        btn_template.grid(row=0, column=0, padx=5, pady=5)

        btn_import = ttk.Button(frame, text="Import Store Sheet", command=self.import_store_sheet)
        btn_import.grid(row=0, column=1, padx=5, pady=5)

        btn_export = ttk.Button(frame, text="Export Master Files", command=self.export_files)
        btn_export.grid(row=0, column=2, padx=5, pady=5)

        btn_export_foil = ttk.Button(frame, text="Export Foil Pan Order Form", command=self.export_foil_pan_order)
        btn_export_foil.grid(row=0, column=3, padx=5, pady=5)

        btn_export_json = ttk.Button(frame, text="Export Data (JSON)", command=self.export_json_data)
        btn_export_json.grid(row=1, column=0, padx=5, pady=5)

        btn_import_json = ttk.Button(frame, text="Import Data (JSON)", command=self.import_json_data)
        btn_import_json.grid(row=1, column=1, padx=5, pady=5)

        # Store upload status columns (centered text, green checkmark)
        store_status_frame = ttk.LabelFrame(frame, text="Store Upload Status")
        store_status_frame.grid(row=2, column=0, columnspan=4, pady=10, sticky="we")
        for col in range(2):
            store_status_frame.grid_columnconfigure(col, weight=1)

        self.store_labels_col1 = []
        self.store_labels_col2 = []

        label_font = ("Arial", 12, "bold")

        for row, store in enumerate(STORE_COL1):
            lbl = tk.Label(store_status_frame, text="", anchor='center', width=8, font=label_font, justify='center')
            lbl.grid(row=row, column=0, sticky="we", padx=2, pady=1)
            self.store_labels_col1.append(lbl)

        for row, store in enumerate(STORE_COL2):
            lbl = tk.Label(store_status_frame, text="", anchor='center', width=8, font=label_font, justify='center')
            lbl.grid(row=row, column=1, sticky="we", padx=2, pady=1)
            self.store_labels_col2.append(lbl)

        self.update_store_status_display()

        path_frame = ttk.LabelFrame(frame, text="Folders")
        path_frame.grid(row=3, column=0, columnspan=4, pady=10, sticky='we')

        ttk.Label(path_frame, text="Downloads Folder:").grid(row=0, column=0, sticky='e')
        self.download_entry = ttk.Entry(path_frame, width=50)
        self.download_entry.grid(row=0, column=1, padx=5)
        self.download_entry.insert(0, self.config.get("download_path", ""))
        ttk.Button(path_frame, text="Set", command=self.set_download_path).grid(row=0, column=2)

        ttk.Label(path_frame, text="Export Folder:").grid(row=1, column=0, sticky='e')
        self.export_entry = ttk.Entry(path_frame, width=50)
        self.export_entry.grid(row=1, column=1, padx=5)
        self.export_entry.insert(0, self.config.get("export_path", ""))
        ttk.Button(path_frame, text="Set", command=self.set_export_path).grid(row=1, column=2)

        imported_frame = ttk.LabelFrame(frame, text="Imported Stores (all)")
        imported_frame.grid(row=4, column=0, columnspan=4, pady=10, sticky="we")
        self.imported_stores_var = tk.StringVar()
        self.imported_stores_label = ttk.Label(imported_frame, textvariable=self.imported_stores_var, anchor="w", justify="left")
        self.imported_stores_label.pack(fill='both', expand=True)
        self.update_imported_stores_display()

        self.status = ttk.Label(frame, text="Status: Ready")
        self.status.grid(row=5, column=0, columnspan=4, pady=10)

    def update_store_status_display(self):
        check, cross = "\u2714", "\u2716"
        for i, store in enumerate(STORE_COL1):
            uploaded = store in self.data
            if uploaded:
                self.store_labels_col1[i]['text'] = f"{int(store):3} {check}"
                self.store_labels_col1[i]['fg'] = "#1ca41c"  # green
            else:
                self.store_labels_col1[i]['text'] = f"{int(store):3} {cross}"
                self.store_labels_col1[i]['fg'] = "black"
        for i, store in enumerate(STORE_COL2):
            uploaded = store in self.data
            if uploaded:
                self.store_labels_col2[i]['text'] = f"{int(store):3} {check}"
                self.store_labels_col2[i]['fg'] = "#1ca41c"  # green
            else:
                self.store_labels_col2[i]['text'] = f"{int(store):3} {cross}"
                self.store_labels_col2[i]['fg'] = "black"

    def update_imported_stores_display(self):
        if self.data:
            stores = sorted(self.data.keys(), key=lambda x: (len(x), x))
            self.imported_stores_var.set(", ".join(stores))
        else:
            self.imported_stores_var.set("None")

    def set_download_path(self):
        path = filedialog.askdirectory()
        if path:
            self.config["download_path"] = path
            self.download_entry.delete(0, tk.END)
            self.download_entry.insert(0, path)
            self.save_config()

    def set_export_path(self):
        path = filedialog.askdirectory()
        if path:
            self.config["export_path"] = path
            self.export_entry.delete(0, tk.END)
            self.export_entry.insert(0, path)
            self.save_config()

    def import_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            try:
                book = xlrd.open_workbook(path)
                sheet = book.sheet_by_index(0)
                self.template["date"] = sheet.cell_value(3, 2)  # C4 is (3,2)
                items = []
                item_names = []
                # Rows 8-44 are 1-based; xlrd is 0-based => rows 7 to 43 inclusive
                for r in range(7, 44):
                    case_qty = str(sheet.cell_value(r, 0)).strip()
                    size = str(sheet.cell_value(r, 1)).strip()
                    desc = str(sheet.cell_value(r, 2)).strip()
                    items.append({'case_qty': case_qty, 'size': size, 'description': desc})
                    # For display, concatenate as requested
                    item_names.append(f"{desc}, {size}, {case_qty}")
                self.template["items"] = items
                self.template["item_names"] = item_names
                self.template["template_path"] = path
                self.status.config(text=f"Template imported: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import template: {e}")

    def load_excel_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".xls":
            book = xlrd.open_workbook(path)
            sheet = book.sheet_by_index(0)
            if sheet.nrows < 44 or sheet.ncols < 7:
                raise ValueError(f"Sheet too small: found {sheet.nrows} rows and {sheet.ncols} columns.")
            try:
                store_cell = sheet.cell_value(2, 6)  # G3
                if isinstance(store_cell, float):
                    store = f"{int(store_cell):03}"
                else:
                    store = str(store_cell).zfill(3)
            except IndexError:
                raise ValueError("Store code cell G3 (row 3, col 7) is missing in the sheet.")
            # Inventory data from D8:D44 (col 3, rows 7..43)
            inventory = []
            for i in range(7, 44):
                try:
                    inventory.append(sheet.cell_value(i, 3))
                except IndexError:
                    inventory.append("")
            foil = []
            for i in range(7, 11):
                try:
                    foil.append(sheet.cell_value(i, 6))
                except IndexError:
                    foil.append("")
        else:
            raise ValueError("Unsupported file format. Only .xls files are supported.")
        return store, inventory, foil

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
            self.update_imported_stores_display()
            self.status.config(text=f"Imported stores: {', '.join(imported_stores)}")

    def export_files(self):
        export_path = self.config.get("export_path", ".")
        if not os.path.exists(export_path):
            messagebox.showerror("Error", "Export folder does not exist.")
            return

        wb_inv = xlwt.Workbook()
        ws_inv = wb_inv.add_sheet('Inventory')
        ws_inv.write(0, 0, "Store")
        for i in range(37):  # Up to 37 items (rows 8-44)
            ws_inv.write(0, i+1, f"Item {i+1}")
        for row_idx, (store, info) in enumerate(self.data.items(), start=1):
            ws_inv.write(row_idx, 0, store)
            for col_idx, val in enumerate(info['inventory']):
                ws_inv.write(row_idx, col_idx+1, val)
        inv_file = os.path.join(export_path, "master_inventory.xls")
        wb_inv.save(inv_file)

        wb_foil = xlwt.Workbook()
        ws_foil = wb_foil.add_sheet('Foil')
        ws_foil.write(0, 0, "Store")
        for i in range(4):
            ws_foil.write(0, i+1, f"Foil {i+1}")
        for row_idx, (store, info) in enumerate(self.data.items(), start=1):
            ws_foil.write(row_idx, 0, store)
            for col_idx, val in enumerate(info['foil']):
                ws_foil.write(row_idx, col_idx+1, val)
        foil_file = os.path.join(export_path, "foil_orders.xls")
        wb_foil.save(foil_file)

        self.status.config(text=f"Export complete: {len(self.data)} stores. Files saved to {export_path}")

    def export_foil_pan_order(self):
        template_path = self.template.get("template_path")
        if not template_path or not os.path.exists(template_path):
            template_path = filedialog.askopenfilename(
                title="Select Foil Pan Order Template (.xls)",
                filetypes=[("Excel 97-2003", "*.xls")]
            )
            if not template_path:
                messagebox.showwarning("No Template", "No template selected.")
                return

        export_folder = self.config.get("export_path", ".")
        if not os.path.exists(export_folder):
            messagebox.showerror("Error", "Export folder does not exist.")
            return

        export_date = self.template.get("date", datetime.today().strftime("%m-%d-%Y"))
        output_file = f"Foil Pan Order {export_date}.xls"
        try:
            import shutil
            shutil.copy(template_path, os.path.join(export_folder, output_file))
            self.status.config(text=f"Foil Pan Order exported: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export foil pan order: {e}")

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
                self.config = imported.get("config", {"download_path": "", "export_path": ""})
                self.template = imported.get("template", {})

                self.download_entry.delete(0, tk.END)
                self.download_entry.insert(0, self.config.get("download_path", ""))
                self.export_entry.delete(0, tk.END)
                self.export_entry.insert(0, self.config.get("export_path", ""))

                self.save_data()
                self.save_config()
                self.update_store_status_display()
                self.update_imported_stores_display()

                self.status.config(text=f"Data imported from {os.path.basename(import_path)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import data: {e}")

    def open_table_editor(self):
        # Table-based editor: first column is item name, then store columns
        if "item_names" not in self.template or not self.template["item_names"]:
            messagebox.showerror("No Template", "You must import a template before editing inventory in table view.")
            return

        item_names = self.template["item_names"]
        n_items = len(item_names)
        if not item_names:
            messagebox.showerror("No Item Names", "No item names found in template.")
            return

        editor = tk.Toplevel(self.root)
        editor.title("Inventory Data Table Editor")
        editor.geometry("1200x700")
        editor_frame = ttk.Frame(editor, padding=10)
        editor_frame.pack(fill='both', expand=True)

        # Scrollbars
        xscroll = tk.Scrollbar(editor_frame, orient="horizontal")
        xscroll.pack(side="bottom", fill="x")
        yscroll = tk.Scrollbar(editor_frame, orient="vertical")
        yscroll.pack(side="right", fill="y")

        # Table: first column is "Item", then store columns
        columns = ["Item"] + ALL_STORES
        tree = ttk.Treeview(
            editor_frame,
            columns=columns,
            show="headings",
            height=min(len(item_names), 25),
            xscrollcommand=xscroll.set,
            yscrollcommand=yscroll.set
        )
        xscroll.config(command=tree.xview)
        yscroll.config(command=tree.yview)

        for col in columns:
            tree.heading(col, text=col)
            if col == "Item":
                tree.column(col, width=260, anchor='w', minwidth=130, stretch=False)
            else:
                tree.column(col, width=60, anchor='center', minwidth=45, stretch=True)
        tree.pack(side="left", fill="both", expand=True)

        # Add rows: first cell is item name, then inventory values per store
        for idx, item_name in enumerate(item_names):
            row_vals = [item_name]
            for store in ALL_STORES:
                inv = self.data.get(store, {}).get("inventory", [""] * n_items)
                val = inv[idx] if idx < len(inv) else ""
                row_vals.append(str(val))
            tree.insert("", "end", values=row_vals, tags=(f"row_{idx}",))

        # Cell editing logic
        def on_double_click(event):
            region = tree.identify("region", event.x, event.y)
            if region != "cell":
                return
            rowid = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            col_index = int(col.replace("#", "")) - 1
            if rowid == "" or col_index == 0:  # Do not edit item names
                return
            x, y, width, height = tree.bbox(rowid, col)
            value = tree.set(rowid, columns[col_index])
            entry = tk.Entry(tree, width=8)
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

        # Save button
        def save_table_edits():
            for idx, rowid in enumerate(tree.get_children()):
                values = tree.item(rowid)["values"]
                for col_idx, store in enumerate(ALL_STORES, start=1):
                    val = values[col_idx]
                    inv = self.data.get(store, {}).get("inventory", [""] * n_items)
                    while len(inv) < n_items:
                        inv.append("")
                    inv[idx] = val
                    foil = self.data.get(store, {}).get("foil", [""]*4)
                    self.data[store] = {"inventory": inv, "foil": foil}
            self.save_data()
            self.update_store_status_display()
            self.update_imported_stores_display()
            messagebox.showinfo("Saved", "All table edits have been saved.")
            editor.lift()
            self.status.config(text="All table edits saved.")

        savebtn = ttk.Button(editor_frame, text="Save All Changes", command=save_table_edits)
        savebtn.pack(side="bottom", pady=5)

def main():
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
