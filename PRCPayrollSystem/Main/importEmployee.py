import customtkinter as ctk
from tkinter import font
import tkinter.filedialog as fd
import pandas as pd
import tkinter.messagebox as messagebox
import tkinter as tk
import sys
import os
from PRCPayrollSystem.Main.resource_utils import resource_path

class ExcelLikeTable(ctk.CTkFrame):
    def __init__(self, parent, rows=10, cols=10):
        super().__init__(parent, fg_color="white", border_width=2, border_color="#4B2ED5")
        self.rows = rows
        self.cols = cols
        self.cells = {}
        self._build_table()

    def _build_table(self):
        header_font = ("Arial", 10, "bold")
        for r in range(self.rows):
            for c in range(self.cols):
                entry = ctk.CTkEntry(self, width=90, justify="center", border_width=1, corner_radius=0)
                if r == 0:
                    entry.configure(fg_color="#22C32A", text_color="white", font=header_font)
                else:
                    entry.configure(font=("Arial", 10))
                entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
                self.cells[(r, c)] = entry
        for c in range(self.cols):
            self.grid_columnconfigure(c, weight=1)
        for r in range(self.rows):
            self.grid_rowconfigure(r, weight=1)

    def add_row(self):
        r = self.rows
        for c in range(self.cols):
            entry = ctk.CTkEntry(self, width=90, justify="center", border_width=1, corner_radius=0)
            entry.configure(font=("Arial", 10))
            entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
            self.cells[(r, c)] = entry
        self.grid_rowconfigure(r, weight=1)
        self.rows += 1

    def remove_row(self):
        if self.rows <= 2:
            return
        r = self.rows - 1
        for c in range(self.cols):
            entry = self.cells.pop((r, c))
            entry.destroy()
        self.rows -= 1

    def add_col(self):
        c = self.cols
        header_font = ("Arial", 10, "bold")
        for r in range(self.rows):
            entry = ctk.CTkEntry(self, width=90, justify="center", border_width=1, corner_radius=0)
            if r == 0:
                entry.configure(fg_color="#22C32A", text_color="white", font=header_font)
            else:
                entry.configure(font=("Arial", 10))
            entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
            self.cells[(r, c)] = entry
        self.grid_columnconfigure(c, weight=1)
        self.cols += 1

    def remove_col(self):
        if self.cols <= 2:
            return
        c = self.cols - 1
        for r in range(self.rows):
            entry = self.cells.pop((r, c))
            entry.destroy()
        self.cols -= 1

class ImportEmployeePage(ctk.CTkFrame):
    def __init__(self, parent, controller=None):
        super().__init__(parent, fg_color="white")
        self.rows = 10
        self.cols = 10
        self.controller = controller
        self._build_ui()
        self._load_default_employee_csv()

    def _build_ui(self):
        # Top bar
        topbar = ctk.CTkFrame(self, fg_color="white")
        topbar.pack(fill=ctk.X, pady=(5, 0))
        def go_to_main():
            self.controller.show_frame("MainMenu")
        ctk.CTkButton(topbar, text='>', font=('Consolas', 18, 'bold'), fg_color="white", text_color="#0041C2", hover_color="#e6e6e6", width=40, height=40, corner_radius=20, command=go_to_main).pack(side=ctk.LEFT, padx=(10, 20))
        btn_style = {'fg_color': '#0041C2', 'text_color': 'white', 'hover_color': '#003399', 'font': ("Arial", 12, "bold"), 'corner_radius': 8, 'width': 160, 'height': 32}
        # Import Excel button only
        ctk.CTkButton(topbar, text='Import Excel', **btn_style, command=self.import_excel).pack(side=ctk.LEFT, padx=(0, 15))
        ctk.CTkButton(topbar, text='Save', **btn_style, command=self.save_employees).pack(side=ctk.LEFT, padx=(0, 15))

        
        # Controls
        controls = ctk.CTkFrame(topbar, fg_color="white")
        controls.pack(side=ctk.RIGHT, padx=10)
        ctk.CTkLabel(controls, text='Columns', font=('Arial', 11), text_color="#0041C2").grid(row=0, column=0, padx=(0, 2))
        ctk.CTkButton(controls, text='+', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.add_col).grid(row=0, column=1)
        ctk.CTkButton(controls, text='-', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.remove_col).grid(row=0, column=2)
        ctk.CTkLabel(controls, text='Rows', font=('Arial', 11), text_color="#0041C2").grid(row=0, column=3, padx=(15, 2))
        ctk.CTkButton(controls, text='+', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.add_row).grid(row=0, column=4)
        ctk.CTkButton(controls, text='-', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.remove_row).grid(row=0, column=5)
        # Table (Scrollable)
        table_frame = ctk.CTkFrame(self, fg_color="white", border_width=1, border_color="#4B2ED5")
        table_frame.pack(fill=ctk.BOTH, expand=True, padx=2, pady=(5, 2))
        # Scrollable canvas
        self.canvas = ctk.CTkCanvas(table_frame, bg="white", highlightthickness=0)
        v_scroll = ctk.CTkScrollbar(table_frame, orientation="vertical", command=self.canvas.yview)
        h_scroll = ctk.CTkScrollbar(table_frame, orientation="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        # Frame inside canvas for the table
        self.inner_table_frame = ctk.CTkFrame(self.canvas, fg_color="white")
        self.table = ExcelLikeTable(self.inner_table_frame, self.rows, self.cols)
        self.table.pack(fill=ctk.BOTH, expand=True)
        self.inner_table_frame.bind(
            "<Configure>", lambda e: self._update_scrollregion())
        self.canvas.create_window((0, 0), window=self.inner_table_frame, anchor="nw")
        # Mousewheel scrolling
        def _on_mousewheel(event):
            if self.canvas.bbox("all") and self.canvas.winfo_height() < self.canvas.bbox("all")[3]:
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        def _on_shift_mousewheel(event):
            if self.canvas.bbox("all") and self.canvas.winfo_width() < self.canvas.bbox("all")[2]:
                self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)
        self._canvas = self.canvas

    def _update_scrollregion(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        # Hide vertical scrollbar 
        if self.canvas.bbox("all") and self.canvas.winfo_height() >= self.canvas.bbox("all")[3]:
            self.canvas.yview_moveto(0)
            self.canvas.configure(yscrollcommand=lambda *args: None)
        else:
            self.canvas.configure(yscrollcommand=self._canvas.master.children['!ctkscrollbar'].set)
        # Hide horizontal scrollbar 
        if self.canvas.bbox("all") and self.canvas.winfo_width() >= self.canvas.bbox("all")[2]:
            self.canvas.xview_moveto(0)
            self.canvas.configure(xscrollcommand=lambda *args: None)
        else:
            self.canvas.configure(xscrollcommand=self._canvas.master.children['!ctkscrollbar2'].set)

    def destroy(self):
        if hasattr(self, '_canvas'):
            self._canvas.unbind_all("<MouseWheel>")
            self._canvas.unbind_all("<Shift-MouseWheel>")
        super().destroy()

    def add_row(self):
        self.rows += 1
        self.table.add_row()
        self.inner_table_frame.update_idletasks()

    def remove_row(self):
        if self.rows > 2:
            self.rows -= 1
            self.table.remove_row()
            self.inner_table_frame.update_idletasks()

    def add_col(self):
        self.cols += 1
        self.table.add_col()
        self.inner_table_frame.update_idletasks()

    def remove_col(self):
        if self.cols > 2:
            self.cols -= 1
            self.table.remove_col()
            self.inner_table_frame.update_idletasks()

    def import_excel(self):
        file_path = fd.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
        if not file_path:
            return
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            root = tk.Toplevel(self)
            root.title("Select Sheet")
            root.geometry("350x180")
            root.update_idletasks()
            # Center the popup
            x = root.winfo_screenwidth() // 2 - 175
            y = root.winfo_screenheight() // 2 - 90
            root.geometry(f"350x180+{x}+{y}")
            tk.Label(root, text="Select sheet to import:", font=("Arial", 13, "bold")).pack(padx=20, pady=(20, 10))
            var = tk.StringVar(value=sheet_names[0])
            dropdown = tk.OptionMenu(root, var, *sheet_names)
            dropdown.config(font=("Arial", 12), width=22)
            dropdown.pack(padx=20, pady=5)
            selected = {'value': None}
            def on_ok():
                selected['value'] = var.get()
                root.destroy()
            tk.Button(root, text="OK", font=("Arial", 14, "bold"), width=12, height=2, command=on_ok).pack(pady=18)
            root.grab_set()
            root.wait_window()
            selected_sheet = selected['value']
            if not selected_sheet or selected_sheet not in sheet_names:
                messagebox.showerror("Error", f"Sheet not selected or not found.")
                return
            df = pd.read_excel(file_path, sheet_name=selected_sheet, header=None)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")
            return
        rows, cols = df.shape
        while self.table.rows > rows:
            self.table.remove_row()
        while self.table.cols > cols:
            self.table.remove_col()
        while self.table.rows < rows:
            self.table.add_row()
        while self.table.cols < cols:
            self.table.add_col()
        header_font = ("Arial", 10, "bold")
        for r in range(self.table.rows):
            for c in range(self.table.cols):
                entry = self.table.cells.get((r, c))
                if entry:
                    entry.delete(0, ctk.END)
                    if r == 0:
                        entry.configure(fg_color="#22C32A", text_color="white", font=header_font)
                    else:
                        entry.configure(fg_color="white", text_color="black", font=("Arial", 10))
        for r in range(rows):
            for c in range(cols):
                value = df.iat[r, c]
                entry = self.table.cells.get((r, c))
                if entry and pd.notna(value):
                    # Limit numbers to 2 decimal places if value is a number
                    if isinstance(value, (int, float)):
                        entry.insert(0, f"{value:.2f}")
                    else:
                        entry.insert(0, str(value))

    def save_employees(self):
        # Save the employee table data CSV file in past playslips 
        import os
        file_path = resource_path("settingsAndFields/employee_records.csv")
        # Gather data from the table
        data = []
        for r in range(self.table.rows):
            row = []
            for c in range(self.table.cols):
                entry = self.table.cells.get((r, c))
                row.append(entry.get() if entry else "")
            data.append(row)
        # Save to CSV
        import csv
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(data)
            messagebox.showinfo('Save', f'Employee information saved to {file_path}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save file: {e}')

    def _load_default_employee_csv(self):
        import os
        import csv
        file_path = resource_path("settingsAndFields/employee_records.csv")
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = list(csv.reader(f))
            if reader:
                rows = len(reader)
                cols = max(len(row) for row in reader)
                # Resize table
                while self.table.rows > rows:
                    self.table.remove_row()
                while self.table.cols > cols:
                    self.table.remove_col()
                while self.table.rows < rows:
                    self.table.add_row()
                while self.table.cols < cols:
                    self.table.add_col()
                # Fill table
                for r in range(rows):
                    for c in range(cols):
                        entry = self.table.cells.get((r, c))
                        if entry:
                            entry.delete(0, ctk.END)
                            value = reader[r][c] if c < len(reader[r]) else ""
                            entry.insert(0, value)
