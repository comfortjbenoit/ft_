import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import xlrd
import xlwt
from datetime import datetime

#

CONFIG_FILE = "config.json"
DATA_FILE = "data.json"

STORE_CODES = [
    "001", "003", "004", "005", "007", "008", "010", "011", "012", "014", "015", "017", "018", "019",
    "201", "202", "203", "204", "205", "206", "207", "208", "209", "211", "214", "215", "216", "217"
]

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

    def build_gui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill='both', expand=True)

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

        self.check_vars = {}
        check_frame = ttk.LabelFrame(frame, text="Stores")
        check_frame.grid(row=2, column=0, columnspan=4, pady=10)
        for idx, code in enumerate(STORE_CODES):
            var = tk.BooleanVar(value=code in self.data)
            cb = ttk.Checkbutton(check_frame, text=code, variable=var)
            cb.grid(row=idx//7, column=idx%7, sticky='w')
            self.check_vars[code] = var

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

        self.status = ttk.Label(frame, text="Status: Ready")
        self.status.grid(row=4, column=0, columnspan=4, pady=10)

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
                self.template["items"] = [
                    [sheet.cell_value(r, c) for c in range(3)] for r in range(7, 37)
                ]  # A8:C37 is (7,0:37,2)
                self.template["template_path"] = path
                self.status.config(text=f"Template imported: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import template: {e}")

    def load_excel_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".xls":
            book = xlrd.open_workbook(path)
            sheet = book.sheet_by_index(0)
            store = str(sheet.cell_value(2, 6)).zfill(3)  # G3 is (2,6)
            inventory = [sheet.cell_value(i, 3) for i in range(7, 37)]  # D8:D37
            foil = [sheet.cell_value(i, 6) for i in range(7, 11)]  # G8:G11
        else:
            raise ValueError("Unsupported file format. Only .xls files are supported.")
        return store, inventory, foil

    def import_store_sheet(self):
        path = filedialog.askopenfilename(
            initialdir=self.config.get("download_path", ""),
            filetypes=[("Excel 97-2003", "*.xls")]
        )
        if path:
            try:
                store, inventory, foil = self.load_excel_file(path)
                self.data[store] = {"inventory": inventory, "foil": foil}
                if store in self.check_vars:
                    self.check_vars[store].set(True)
                self.save_data()
                self.status.config(text=f"Imported store {store} from {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to read file: {e}")

    def export_files(self):
        export_path = self.config.get("export_path", ".")
        if not os.path.exists(export_path):
            messagebox.showerror("Error", "Export folder does not exist.")
            return

        # Export master inventory as .xls
        wb_inv = xlwt.Workbook()
        ws_inv = wb_inv.add_sheet('Inventory')
        ws_inv.write(0, 0, "Store")
        for i in range(30):
            ws_inv.write(0, i+1, f"Item {i+1}")
        for row_idx, (store, info) in enumerate(self.data.items(), start=1):
            ws_inv.write(row_idx, 0, store)
            for col_idx, val in enumerate(info['inventory']):
                ws_inv.write(row_idx, col_idx+1, val)
        inv_file = os.path.join(export_path, "master_inventory.xls")
        wb_inv.save(inv_file)

        # Export foil pan orders as .xls
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
        # Use template file if imported, or ask user to select a .xls template
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
            # Copy the template file as the output
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
                self.config = imported.get("config", {"download_path": "", "export_path": ""})
                self.template = imported.get("template", {})

                for code, var in self.check_vars.items():
                    var.set(code in self.data)

                self.download_entry.delete(0, tk.END)
                self.download_entry.insert(0, self.config.get("download_path", ""))
                self.export_entry.delete(0, tk.END)
                self.export_entry.insert(0, self.config.get("export_path", ""))

                self.save_data()
                self.save_config()

                self.status.config(text=f"Data imported from {os.path.basename(import_path)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import data: {e}")

def main():
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
