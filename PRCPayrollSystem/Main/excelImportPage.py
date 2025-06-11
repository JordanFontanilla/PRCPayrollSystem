import customtkinter as ctk
import tkinter as tk
import os
import csv
import glob
from PRCPayrollSystem.Main.resource_utils import resource_path

class ExcelImportPage(ctk.CTkFrame):
    def __init__(self, parent, controller=None):
        super().__init__(parent, fg_color="white")
        self.controller = controller
        self.rows = 8  # 7 data + 1 header (DEFAULT LOGIC)
        self.cols = 11
        # Directory to save history files (always same folder)
        self.history_dir = resource_path("pastLoadedHistory")
        self._build_ui()

    def _build_ui(self):
        # Top bar
        topbar = ctk.CTkFrame(self, fg_color="white")
        topbar.pack(fill=ctk.X, pady=(5, 0))
        def go_to_main():
            if self.controller:
                self.controller.show_frame("MainMenu")
        ctk.CTkButton(topbar, text='>', font=('Consolas', 18, 'bold'), fg_color="white", text_color="#0041C2", hover_color="#e6e6e6", width=40, height=40, corner_radius=20, command=go_to_main).pack(side=ctk.LEFT, padx=(10, 20))
        # Add Generate Reports and Generate Payslip buttons
        btn_style_top = {'fg_color': '#003399', 'text_color': 'white', 'hover_color': '#002266', 'font': ("Arial", 15, "bold"), 'corner_radius': 20, 'width': 180, 'height': 40}
        btn_frame = ctk.CTkFrame(topbar, fg_color="white")
        btn_frame.pack(side=ctk.LEFT, padx=(10, 0))
        ctk.CTkButton(btn_frame, text='Generate Reports', **btn_style_top, command=self.go_to_reports_page).pack(side=ctk.LEFT, padx=(0, 20))
        ctk.CTkButton(btn_frame, text='Generate Payslip', **btn_style_top, command=self.go_to_generate_payslip).pack(side=ctk.LEFT)
        btn_style = {'fg_color': '#0041C2', 'text_color': 'white', 'hover_color': '#003399', 'font': ("Arial", 12, "bold"), 'corner_radius': 8, 'width': 160, 'height': 32}
        # Controls
        btn_style = {'fg_color': '#0041C2', 'text_color': 'white', 'hover_color': '#003399', 'font': ("Arial", 12, "bold"), 'corner_radius': 8, 'width': 120, 'height': 32}
        controls = ctk.CTkFrame(topbar, fg_color="white")
        controls.pack(side=ctk.RIGHT, padx=10)
        ctk.CTkLabel(controls, text='Columns', font=('Arial', 11), text_color="#0041C2").grid(row=0, column=0, padx=(0, 2))
        ctk.CTkButton(controls, text='+', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.add_col).grid(row=0, column=1)
        ctk.CTkButton(controls, text='-', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.remove_col).grid(row=0, column=2)
        ctk.CTkLabel(controls, text='Rows', font=('Arial', 11), text_color="#0041C2").grid(row=0, column=3, padx=(15, 2))
        ctk.CTkButton(controls, text='+', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.add_row).grid(row=0, column=4)
        ctk.CTkButton(controls, text='-', font=('Arial', 13, 'bold'), text_color="#0041C2", fg_color="white", hover_color="#e6e6e6", width=32, height=32, command=self.remove_row).grid(row=0, column=5)
        # Add Import Excel button next to controls
        ctk.CTkButton(controls, text='Import Excel', **btn_style, command=self.import_excel).grid(row=0, column=6, padx=(10, 0))
        # Table frame
        table_frame = ctk.CTkFrame(self, fg_color="white", border_width=1, border_color="#4B2ED5")
        table_frame.pack(fill=ctk.BOTH, expand=True, padx=8, pady=(10, 8))
        # Canvas for scrolling
        self.canvas = ctk.CTkCanvas(table_frame, bg="white", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll = ctk.CTkScrollbar(table_frame, orientation="vertical", command=self.canvas.yview)
        self.h_scroll = ctk.CTkScrollbar(table_frame, orientation="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.inner = ctk.CTkFrame(self.canvas, fg_color="white")
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self._draw_table()
        self.inner.bind("<Configure>", lambda e: self._update_scrollregion())
        def _on_mousewheel(event):
            if self.canvas.bbox("all") and self.canvas.winfo_height() < self.canvas.bbox("all")[3]:
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        def _on_shift_mousewheel(event):
            if self.canvas.bbox("all") and self.canvas.winfo_width() < self.canvas.bbox("all")[2]:
                self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)
        self._canvas = self.canvas

    def _draw_table(self):
        for widget in self.inner.winfo_children():
            widget.destroy()
        header_font = ("Arial", 10, "bold")
        cell_font = ("Arial", 10)
        for r in range(self.rows):
            for c in range(self.cols):
                if r == 0 and c == 0:
                    entry = ctk.CTkLabel(self.inner, text="", width=30, fg_color="#e6e6e6")
                elif r == 0:
                    col_letter = chr(64 + c) if c <= 26 else chr(64 + (c-1)//26) + chr(65 + (c-1)%26)
                    entry = ctk.CTkLabel(self.inner, text=col_letter, width=90, fg_color="#22C32A", text_color="white", font=header_font)
                elif c == 0:
                    entry = ctk.CTkLabel(self.inner, text=str(r), width=30, fg_color="#e6e6e6", font=header_font)
                else:
                    entry = ctk.CTkEntry(self.inner, width=90, justify="center", border_width=1, corner_radius=0)
                    entry.configure(font=cell_font, fg_color="white", text_color="black")
                    # Always format to two decimals on focus out
                    def format_decimal(event, e=entry):
                        val = e.get()
                        try:
                            if val.strip() != "":
                                num = float(val)
                                e.delete(0, tk.END)
                                e.insert(0, f"{num:.2f}")
                        except:
                            pass
                    entry.bind("<FocusOut>", format_decimal)
                    # Fill with data if present
                    if hasattr(self, 'data') and r-1 < len(self.data) and c-1 < len(self.data[r-1]):
                        val = self.data[r-1][c-1]
                        try:
                            if val != '' and val is not None:
                                val = float(val)
                                entry.insert(0, f"{val:.2f}")
                            else:
                                entry.insert(0, "")
                        except:
                            entry.insert(0, str(val))
                entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
                self.inner.grid_columnconfigure(c, weight=1)
            self.inner.grid_rowconfigure(r, weight=1)

    def enforce_two_decimal_all(self):
        # for decimal formatting
        for r in range(1, self.rows):
            for c in range(1, self.cols):
                widgets = self.inner.grid_slaves(row=r, column=c)
                for widget in widgets:
                    if isinstance(widget, ctk.CTkEntry):
                        value = widget.get()
                        if value.strip() == "":
                            continue
                        try:
                            num = float(value.replace(',', ''))
                            formatted = f"{num:,.2f}"
                            widget.delete(0, 'end')
                            widget.insert(0, formatted)
                        except Exception:
                            continue

    def _update_header(self, col, value):
        self.headers[col] = value

    def _update_data(self, row, col, value):
        self.data[row][col] = value

    def _update_scrollregion(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        if self.canvas.bbox("all") and self.canvas.winfo_height() >= self.canvas.bbox("all")[3]:
            self.canvas.yview_moveto(0)
            self.canvas.configure(yscrollcommand=lambda *args: None)
        else:
            self.canvas.configure(yscrollcommand=self.v_scroll.set)
        # Horizontal scroll logic
        if self.canvas.bbox("all") and self.canvas.winfo_width() >= self.canvas.bbox("all")[2]:
            self.canvas.xview_moveto(0)
            self.canvas.configure(xscrollcommand=lambda *args: None)
        else:
            self.canvas.configure(xscrollcommand=self.h_scroll.set)

    def add_row(self):
        # Add a new row at the end 
        header_font = ("Arial", 10, "bold")
        r = self.rows
        # Add row header
        entry = ctk.CTkLabel(self.inner, text=str(r), width=30, fg_color="#e6e6e6", font=header_font)
        entry.grid(row=r, column=0, sticky="nsew", padx=0, pady=0, ipady=6)
        # Add cells for each column
        for c in range(1, self.cols):
            entry = ctk.CTkEntry(self.inner, width=90, justify="center", border_width=1, corner_radius=0)
            entry.configure(font=("Arial", 10), fg_color="white", text_color="black")
            entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
        self.inner.grid_rowconfigure(r, weight=1)
        self.rows += 1
        self.inner.update_idletasks()
        self.enforce_two_decimal_all()

    def remove_row(self):
        if self.rows <= 2:
            return
        r = self.rows - 1
        for c in range(self.cols):
            widgets = self.inner.grid_slaves(row=r, column=c)
            for widget in widgets:
                widget.destroy()
        self.rows -= 1
        self.inner.update_idletasks()

    def add_col(self):
        c = self.cols
        header_font = ("Arial", 10, "bold")
        # Add new column header
        col_letter = chr(64 + c) if c <= 26 else chr(64 + (c-1)//26) + chr(65 + (c-1)%26)
        entry = ctk.CTkLabel(self.inner, text=col_letter, width=90, fg_color="#22C32A", text_color="white", font=header_font)
        entry.grid(row=0, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
        # Add new cells for each row
        for r in range(1, self.rows):
            entry = ctk.CTkEntry(self.inner, width=90, justify="center", border_width=1, corner_radius=0)
            entry.configure(font=("Arial", 10), fg_color="white", text_color="black")
            entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
        self.inner.grid_columnconfigure(c, weight=1)
        self.cols += 1
        self.inner.update_idletasks()
        self.enforce_two_decimal_all()

    def remove_col(self):
        if self.cols <= 2:
            return
        c = self.cols - 1
        for r in range(self.rows):
            widgets = self.inner.grid_slaves(row=r, column=c)
            for widget in widgets:
                widget.destroy()
        self.cols -= 1
        self.inner.update_idletasks()

    def set_aggregated_data(self, headers, data):
        """
        Set the table headers and data to the provided aggregated results and redraw the table.
        headers: list of str
        data: list of lists (rows)
        """
        # Dynamically set rows and cols to fit the data
        self.cols = max(len(headers), max((len(row) for row in data), default=0))
        self.rows = len(data) + 1  
        # Treat the loaded header row as the first data row
        self.headers = [chr(65 + c) if c < 26 else chr(65 + (c // 26) - 1) + chr(65 + (c % 26)) for c in range(self.cols)]
        # Format numeric cells with thousands comma and two decimals
        formatted_data = []
        for row in [headers] + data:
            formatted_row = []
            for val in row[:self.cols]:
                try:
                    # Only format if it's a number and not empty
                    if val is not None and str(val).strip() != "":
                        num = float(str(val).replace(',', ''))
                        formatted_row.append(f"{num:,.2f}")
                    else:
                        formatted_row.append("")
                except Exception:
                    formatted_row.append(val)
            # Pad if row is short
            formatted_row += ["" for _ in range(self.cols - len(formatted_row))]
            formatted_data.append(formatted_row)
        self.data = formatted_data
        while len(self.data) < self.rows:
            self.data.append(["" for _ in range(self.cols)])
        self._draw_table()
        self.inner.update_idletasks()

    def go_to_reports_page(self):
        if self.controller:
            # Get excel data from the table 
            excel_data = []
            for r in range(1, self.rows):
                row = []
                for c in range(1, self.cols):
                    widgets = self.inner.grid_slaves(row=r, column=c)
                    val = ""
                    for widget in widgets:
                        if isinstance(widget, ctk.CTkEntry):
                            v = widget.get()
                            # Remove commas for numeric values
                            try:
                                v_clean = v.replace(",", "")
                                float(v_clean)
                                val = v_clean
                            except Exception:
                                val = v
                    row.append(val)
                excel_data.append(row)
            # Save to controller for access in ReportsPage
            self.controller.excel_data = excel_data
            if hasattr(self.controller, 'get_page'):
                reports_page = self.controller.get_page('ReportsPage')
                if reports_page:
                    reports_page.set_excel_aggregated(excel_data)
            self.controller.show_frame("ReportsPage")

    def go_to_generate_payslip(self):
        if self.controller:
            self.controller.show_frame("GeneratePayslipPage")

    def import_excel(self):
        import tkinter.filedialog as fd
        import pandas as pd
        import tkinter.messagebox as messagebox
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
                return
            df = pd.read_excel(file_path, sheet_name=selected_sheet)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")
            return
        rows, cols = df.shape
        # Adjust table size
        while self.rows > rows+2:  
            self.remove_row()
        while self.cols > cols+1:
            self.remove_col()
        while self.rows < rows+2:
            self.add_row()
        while self.cols < cols+1:
            self.add_col()
        for c in range(cols):
            entry = self.inner.grid_slaves(row=1, column=c+1)
            if entry:
                entry[0].delete(0, tk.END)
                entry[0].insert(0, str(df.columns[c]))
                entry[0].update()
        # Fill data rows 
        for r in range(rows):
            for c in range(cols):
                value = df.iat[r, c]
                if isinstance(value, float):
                    value = f"{value:.2f}"
                entry = self.inner.grid_slaves(row=r+2, column=c+1)
                if entry:
                    entry[0].delete(0, tk.END)
                    entry[0].insert(0, str(value))
        # After loading table_data:
        table_data = self.get_table_data() if hasattr(self, 'get_table_data') else None
        self.save_table_to_history(table_data)
        self.enforce_two_decimal_all()

    def get_table_data(self):
        """
        Returns the table as a list of lists, where the first row is the Excel header (row=1, col=1:cols),
        and subsequent rows are the data (row=2, col=1:cols). Skips the top/left label row and col.
        Removes commas from numbers for correct downstream processing.
        """
        data = []
        for r in range(1, self.rows):
            row = []
            for c in range(1, self.cols):
                widgets = self.inner.grid_slaves(row=r, column=c)
                val = ""
                for widget in widgets:
                    if isinstance(widget, ctk.CTkEntry):
                        v = widget.get()
                        v_clean = v.replace(",", "")
                        val = v_clean
                row.append(val)
            data.append(row)
        return data

    def save_table_to_history(self, table_data):
        import datetime
        if not table_data or len(table_data) < 2:
            return
        if not os.path.exists(self.history_dir):
            os.makedirs(self.history_dir)
        # Use date for filename and a unique ID for the day
        today = datetime.datetime.now().strftime('%Y%m%d')
        # Find all files for today
        existing = [f for f in os.listdir(self.history_dir) if f.startswith(f"excelHistory_{today}_") and f.endswith('.csv')]
        # Extract IDs and find next available
        ids = []
        for f in existing:
            try:
                id_part = f.split('_')[-1].replace('.csv','')
                ids.append(int(id_part))
            except:
                continue
        next_id = max(ids) + 1 if ids else 1
        fname = f"excelHistory_{today}_{next_id:04d}.csv"
        fpath = os.path.join(self.history_dir, fname)
        with open(fpath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for row in table_data:
                writer.writerow(row)
        # Refresh HistoryPage if it exists 
        if self.controller and hasattr(self.controller, 'get_page'):
            history_page = self.controller.get_page('HistoryPage')
            if history_page and hasattr(history_page, 'refresh'):
                history_page.refresh()

    def save_imported_table_to_history(self, data, headers):
        # Determine history folder path
        history_dir = self.history_dir
        os.makedirs(history_dir, exist_ok=True)
        # List all CSVs in the folder
        csv_files = sorted(
            glob.glob(os.path.join(history_dir, '*.csv')),
            key=os.path.getmtime
        )
        if len(csv_files) >= 100:
            # Get oldest file
            oldest_file = csv_files[0]
            oldest_filename = os.path.basename(oldest_file)
            # Prompt user
            def on_delete():
                try:
                    os.remove(oldest_file)
                except Exception as e:
                    ctk.CTkMessagebox(title="Error", message=f"Failed to delete {oldest_filename}: {e}", icon="cancel")
                    popup.destroy()
                    return
                popup.destroy()
                self._do_save_imported_table_to_history(data, headers, history_dir)
            def on_cancel():
                popup.destroy()
            popup = tk.Toplevel(self)
            popup.title("History Limit Reached")
            popup.geometry("420x180")
            popup.grab_set()
            tk.Label(popup, text=f"History is full (100 files).\nDelete oldest file to continue?", font=("Arial", 12, "bold")).pack(pady=(18, 8))
            tk.Label(popup, text=f"Oldest: {oldest_filename}", font=("Arial", 11)).pack(pady=(0, 10))
            btn_frame = tk.Frame(popup)
            btn_frame.pack(pady=10)
            tk.Button(btn_frame, text="Delete & Continue", font=("Arial", 11, "bold"), width=16, command=on_delete).pack(side="left", padx=8)
            tk.Button(btn_frame, text="Cancel", font=("Arial", 11), width=10, command=on_cancel).pack(side="left", padx=8)
            return  
        # If under limit, save directly
        self._do_save_imported_table_to_history(data, headers, history_dir)

    def _do_save_imported_table_to_history(self, data, headers, history_dir):
        import csv
        from datetime import datetime
        today = datetime.now().strftime("%Y%m%d")
        existing = [f for f in os.listdir(history_dir) if f.startswith(f"excelHistory_{today}_") and f.endswith('.csv')]
        ids = []
        for f in existing:
            try:
                id_part = f.split('_')[-1].replace('.csv','')
                ids.append(int(id_part))
            except:
                continue
        next_id = max(ids) + 1 if ids else 1
        filename = f"excelHistory_{today}_{next_id:04d}.csv"
        filepath = os.path.join(history_dir, filename)
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if headers:
                writer.writerow(headers)
            writer.writerows(data)
        if hasattr(self.controller, 'frames') and 'HistoryPage' in self.controller.frames:
            self.controller.frames['HistoryPage'].refresh()

    def destroy(self):
        if hasattr(self, '_canvas'):
            self._canvas.unbind_all("<MouseWheel>")
        super().destroy()
