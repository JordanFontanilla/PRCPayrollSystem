import customtkinter as ctk
import tkinter as tk
import os
import csv

class ReportsPage(ctk.CTkFrame):
    def __init__(self, parent, controller=None):
        super().__init__(parent, fg_color="white")
        self.controller = controller
        self.rows = 8  # REMEMBERRRR: 7+1 because of header row
        self.cols = 11
        self.header_font = ("Arial", 10, "bold")
        self.cell_font = ("Arial", 10)
        self.default_col_headers = [
            "PAP / UAC CODE", "", "MO.SALARY", "PERA AMOUNT", "GSIS", "PHIC", "HDMF", "OTHER DEDUCTIONS"
        ]
        settings_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "deductionSettings.csv")
        loaded_uac = []
        self._selected_other_deduction_cols = []
        self._deduction_colnames = []
        self._deduction_colnames_saved = []
        if os.path.exists(settings_path):
            import csv
            with open(settings_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    if row and row[0].startswith("UAC:"):
                        loaded_uac.append(row[0][4:])
                    elif row and row[0].startswith("DEDCOLS:"):
                        parts = row[0].split(":", 1)[1].split("|")
                        # parts: [colidx0, colidx1, ..., '', colname0, colname1, ...]
                        col_indices = []
                        colnames = []
                        for x in parts[1:]:
                            if x.isdigit():
                                col_indices.append(int(x))
                            else:
                                break
                        # Find the split between indices and colnames
                        idx_split = 1 + len(col_indices)
                        colnames = parts[idx_split:]
                        self._selected_other_deduction_cols = col_indices
                        self._deduction_colnames_saved = [c for c in colnames if c]
        if loaded_uac:
            self.default_row_headers = loaded_uac
        else:
            self.default_row_headers = [
                "A.I.a.1", "A.II.a.1", "A.III.a.4", "A.III.b.6", "A.III.b.7", "A.III.b.8"
            ]
        self._build_ui()

    def _build_ui(self):
        # Top bar
        topbar = ctk.CTkFrame(self, fg_color="white")
        topbar.pack(fill=ctk.X, pady=(5, 0))
        def go_to_excel_import():
            if self.controller:
                self.controller.show_frame("ExcelImportPage")
        ctk.CTkButton(topbar, text='>', font=('Consolas', 18, 'bold'), fg_color="white", text_color="#0041C2", hover_color="#e6e6e6", width=40, height=40, corner_radius=20, command=go_to_excel_import).pack(side=ctk.LEFT, padx=(10, 20))
        btn_style = {'fg_color': '#0041C2', 'text_color': 'white', 'hover_color': '#003399', 'font': ("Arial", 12, "bold"), 'corner_radius': 8, 'width': 160, 'height': 32}
        ctk.CTkLabel(topbar, text='Reports', font=("Montserrat", 32, "normal"), text_color="#222").pack(side=ctk.LEFT, pady=10)
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
        # Add Adjustments button
        ctk.CTkButton(controls, text='Adjustments', **btn_style, command=self.show_adjustments_popup).grid(row=0, column=7, padx=(10,0))
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
        # Default values for the table (table structure from excelImport, easier)
        for r in range(self.rows):
            for c in range(self.cols):
                if r == 0 and c == 0:
                    entry = ctk.CTkLabel(self.inner, text="", width=30, fg_color="#e6e6e6")
                elif r == 0:
                    col_letter = chr(64 + c) if c <= 26 else chr(64 + (c-1)//26) + chr(65 + (c-1)%26)
                    if c == 2:
                        # MO.SALARY header
                        entry = ctk.CTkLabel(self.inner, text=col_letter, width=90, fg_color="#22C32A", text_color="white", font=self.header_font)
                        if c < len(self.default_col_headers):
                            entry.configure(text=self.default_col_headers[c])
                        entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
                    elif c == 1:
                        fg_color = '#e6e6e6'
                        entry = ctk.CTkLabel(self.inner, text=col_letter, width=90, fg_color=fg_color, text_color="white", font=self.header_font)
                        if c < len(self.default_col_headers):
                            entry.configure(text=self.default_col_headers[c])
                        entry.grid(row=r, column=c, sticky="nsew", padx=(0,0), pady=0, ipady=6)
                    else:
                        fg_color = '#22C32A'
                        entry = ctk.CTkLabel(self.inner, text=col_letter, width=90, fg_color=fg_color, text_color="white", font=self.header_font)
                        if c < len(self.default_col_headers):
                            entry.configure(text=self.default_col_headers[c])
                        entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
                elif c == 0:
                    # Numbers column
                    entry = ctk.CTkLabel(self.inner, text=str(r), width=30, fg_color="#e6e6e6", font=self.header_font)
                elif c == 1 and r > 0:
                    # Default row header values in the second column
                    if r <= len(self.default_row_headers):
                        entry = ctk.CTkLabel(self.inner, text=self.default_row_headers[r-1], width=120, fg_color="#e6e6e6", font=self.header_font)
                    else:
                        entry = ctk.CTkLabel(self.inner, text="", width=120, fg_color="#e6e6e6", font=self.header_font)
                else:
                    entry = ctk.CTkEntry(self.inner, width=90, justify="center", border_width=1, corner_radius=0)
                    entry.configure(font=self.cell_font, fg_color="white", text_color="black")
                    if hasattr(self, 'data') and r-1 < len(self.data) and c-2 < len(self.data[r-1]) and c > 1:
                        entry.insert(0, self.data[r-1][c-2])
                entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
                self.inner.grid_columnconfigure(c, weight=1)
            self.inner.grid_rowconfigure(r, weight=1)

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
        # Add a new row at the end (to follow logic of adding and removing)
        r = self.rows
        # Add row header
        entry = ctk.CTkLabel(self.inner, text=str(r), width=30, fg_color="#e6e6e6", font=self.header_font)
        entry.grid(row=r, column=0, sticky="nsew", padx=0, pady=0, ipady=6)
        # Add cells for each column (except header col)
        for c in range(1, self.cols):
            entry = ctk.CTkEntry(self.inner, width=90, justify="center", border_width=1, corner_radius=0)
            entry.configure(font=("Arial", 10), fg_color="white", text_color="black")
            entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
        self.inner.grid_rowconfigure(r, weight=1)
        self.rows += 1
        self.inner.update_idletasks()

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
        # Add new column header
        col_letter = chr(64 + c) if c <= 26 else chr(64 + (c-1)//26) + chr(65 + (c-1)%26)
        entry = ctk.CTkLabel(self.inner, text=col_letter, width=90, fg_color="#22C32A", text_color="white", font=self.header_font)
        entry.grid(row=0, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
        # Add new cells for each row
        for r in range(1, self.rows):
            entry = ctk.CTkEntry(self.inner, width=90, justify="center", border_width=1, corner_radius=0)
            entry.configure(font=("Arial", 10), fg_color="white", text_color="black")
            entry.grid(row=r, column=c, sticky="nsew", padx=0, pady=0, ipady=6)
        self.inner.grid_columnconfigure(c, weight=1)
        self.cols += 1
        self.inner.update_idletasks()

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

    #Data aggregation
    def set_aggregated_data(self, headers, data):
        """
        Set the table headers and data to the provided aggregated results and redraw the table.
        headers: list of str
        data: list of lists (rows)
        """
        self.headers = headers[:self.cols] + ["" for _ in range(self.cols - len(headers))]
        self.data = [row[:self.cols] + ["" for _ in range(self.cols - len(row))] for row in data[:self.rows-1]]
        while len(self.data) < self.rows-1:
            self.data.append(["" for _ in range(self.cols)])
        self._draw_table()
        self.enforce_number_format_all()
        self.inner.update_idletasks()

    def enforce_number_format_all(self):
        # Format all non-header cells to two decimals with commas for thousands
        for r in range(1, self.rows):
            for c in range(2, self.cols):  # Skipping row headers because duhh
                widgets = self.inner.grid_slaves(row=r, column=c)
                for widget in widgets:
                    if isinstance(widget, ctk.CTkEntry):
                        try:
                            value = widget.get()
                            if value.strip() == "":
                                continue
                            num = float(value.replace(',', ''))
                            formatted = f"{num:,.2f}"
                            widget.delete(0, 'end')
                            widget.insert(0, formatted)
                        except Exception:
                            continue

    def show_adjust_other_deductions_popup(self):
        # For the user to be able to select columns for 'OTHER Deductions' from available columns in excelImport table
        popup = tk.Toplevel(self)
        popup.title("Adjust Other Deductions Columns")
        popup.geometry("650x520")
        popup.grab_set()
        tk.Label(popup, text="Select columns to sum for 'Other Deductions':", font=("Arial", 11, "bold")).pack(pady=(18, 8))
        # Container canvas for scrollable area and buttons
        container = tk.Frame(popup)
        container.pack(fill="both", expand=True, padx=20)
        # Scrollable frame for checkboxes
        canvas = tk.Canvas(container, borderwidth=0, height=180)
        frame = tk.Frame(canvas)
        v_scroll = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
        canvas.create_window((0, 0), window=frame, anchor="nw")
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame.bind("<Configure>", on_frame_configure)
        # Get available columns from excelImportPage IF possible
        available_cols = []
        if hasattr(self.controller, 'frames') and 'ExcelImportPage' in self.controller.frames:
            excel_page = self.controller.frames['ExcelImportPage']
            if hasattr(excel_page, 'inner'):
                for c in range(1, excel_page.cols):
                    widgets = excel_page.inner.grid_slaves(row=1, column=c)
                    for widget in widgets:
                        if isinstance(widget, ctk.CTkEntry):
                            col_name = widget.get().strip()
                            if col_name:
                                available_cols.append((c, col_name))
        selected_vars = {}
        if not self._selected_other_deduction_cols:
            self._selected_other_deduction_cols = list(range(9, 21))  # J (9) to U (20) inclusive
        current = set(self._selected_other_deduction_cols)
        for c, col_name in available_cols:
            var = tk.BooleanVar(value=((c-1) in current))
            cb = tk.Checkbutton(frame, text=col_name, variable=var, font=("Arial", 11))
            cb.pack(anchor="w")
            selected_vars[c-1] = var
        error_label = tk.Label(popup, text="", font=("Arial", 10), fg="#d0021b")
        error_label.pack()
        btns_frame = tk.Frame(popup)
        btns_frame.pack(pady=16)
        def on_apply():
            selected = [c for c, var in selected_vars.items() if var.get()]
            if not selected:
                error_label.config(text="Please select at least one column.")
                return
            self._selected_other_deduction_cols = selected
            settings_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "deductionSettings.csv")
            colnames = []
            if hasattr(self.controller, 'frames') and 'ExcelImportPage' in self.controller.frames:
                excel_page = self.controller.frames['ExcelImportPage']
                if hasattr(excel_page, 'inner'):
                    for c in range(1, excel_page.cols):
                        widgets = excel_page.inner.grid_slaves(row=1, column=c)
                        for widget in widgets:
                            if isinstance(widget, ctk.CTkEntry):
                                col_name = widget.get().strip()
                                if col_name:
                                    colnames.append(col_name)
            # Always overwrite the file with only UAC lines and the new DEDCOLS line
            lines = []
            if os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            uac_lines = [line for line in lines if line.startswith('UAC:')]
            try:
                with open(settings_path, 'w', encoding='utf-8', newline='') as f:
                    for line in uac_lines:
                        f.write(line)
                    f.write("DEDCOLS:" + "|".join(str(x) for x in selected) + "|" + "|".join([""] + colnames) + "\n")
            except Exception as e:
                error_label.config(text=f"Error saving settings: {e}")
                return
            self._deduction_colnames_saved = colnames
            popup.destroy()
            self.refresh_aggregation()
        tk.Button(btns_frame, text="Apply", font=("Arial", 12, "bold"), width=12, command=on_apply).pack(side="left", padx=8)
        tk.Button(btns_frame, text="Cancel", font=("Arial", 11), width=10, command=popup.destroy).pack(side="left", padx=8)

    def show_adjustments_popup(self):
        popup = tk.Toplevel(self)
        popup.title("Adjustments")
        popup.geometry("450x220")
        popup.grab_set()
        tk.Label(popup, text="Adjustments", font=("Arial", 14, "bold")).pack(pady=(18, 8))
        btn_add_remove = tk.Button(popup, text="Add/Remove UAC Codes", font=("Arial", 12), width=22, command=lambda: [popup.destroy(), self.show_add_remove_uac_popup()])
        btn_add_remove.pack(pady=10)
        btn_adjust_deductions = tk.Button(popup, text="Adjust deductions", font=("Arial", 12), width=22, command=lambda: [popup.destroy(), self.show_adjust_other_deductions_popup()])
        btn_adjust_deductions.pack(pady=10)
        tk.Button(popup, text="Close", font=("Arial", 11), width=10, command=popup.destroy).pack(pady=10)

    def show_add_remove_uac_popup(self):
        popup = tk.Toplevel(self)
        popup.title("Add/Remove UAC Codes")
        popup.geometry("500x500")
        popup.grab_set()
        tk.Label(popup, text="Edit UAC Codes:", font=("Arial", 12, "bold")).pack(pady=(18, 8))

        # Frame for scrollable UAC list
        list_frame = tk.Frame(popup)
        list_frame.pack(fill=tk.BOTH, expand=False, padx=10)
        canvas = tk.Canvas(list_frame, height=220)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        uac_inner = tk.Frame(canvas)
        uac_inner.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=uac_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # For showing column info
        col_info_label = tk.Label(popup, text="", font=("Arial", 11), fg="#0041C2")
        col_info_label.pack(pady=(8, 0))

        # Store UAC codes locally for editing
        uac_codes = list(self.default_row_headers)

        def refresh_uac_list():
            for widget in uac_inner.winfo_children():
                widget.destroy()
            for idx, code in enumerate(uac_codes):
                row = tk.Frame(uac_inner)
                row.pack(fill=tk.X, pady=2)
                tk.Label(row, text=code, font=("Arial", 12), width=24, anchor="w").pack(side=tk.LEFT, padx=(2, 8))
                del_btn = tk.Button(row, text="Delete", font=("Arial", 10), fg="#d0021b", width=8,
                                    command=lambda i=idx: on_delete(i))
                del_btn.pack(side=tk.LEFT)
                info_btn = tk.Button(row, text="Show Column", font=("Arial", 10), width=12,
                                     command=lambda i=idx: on_show_col(i))
                info_btn.pack(side=tk.LEFT, padx=(8, 0))

        def on_delete(idx):
            code = uac_codes[idx]
            uac_codes.pop(idx)
            refresh_uac_list()
            col_info_label.config(text=f"Deleted UAC: {code}")

        def on_show_col(idx):
            code = uac_codes[idx]
            col_info_label.config(text=f"UAC '{code}' reflects row {idx+1} in the report table.")

        refresh_uac_list()

        # Add new UAC code
        add_frame = tk.Frame(popup)
        add_frame.pack(pady=(16, 0))
        tk.Label(add_frame, text="Add new UAC code:", font=("Arial", 11)).pack(side=tk.LEFT, padx=(0, 6))
        new_uac_var = tk.StringVar()
        new_uac_entry = tk.Entry(add_frame, textvariable=new_uac_var, font=("Arial", 12), width=18)
        new_uac_entry.pack(side=tk.LEFT)
        def on_add():
            val = new_uac_var.get().strip()
            if not val:
                col_info_label.config(text="UAC code cannot be empty.")
                return
            if val in uac_codes:
                col_info_label.config(text="UAC code already exists.")
                return
            uac_codes.append(val)
            new_uac_var.set("")
            refresh_uac_list()
            col_info_label.config(text=f"Added UAC: {val}")
        tk.Button(add_frame, text="Add", font=("Arial", 11, "bold"), width=8, command=on_add).pack(side=tk.LEFT, padx=(8, 0))

        # Save/Cancel buttons
        btns_frame = tk.Frame(popup)
        btns_frame.pack(pady=18)
        def on_apply():
            if not uac_codes:
                col_info_label.config(text="At least one UAC code is required.")
                return
            prev_count = len(self.default_row_headers)
            new_count = len(uac_codes)
            self.default_row_headers = list(uac_codes)
            # Save to deductionSettings.csv in the correct format
            settings_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "deductionSettings.csv")
            # Read existing lines to preserve DEDCOLS if present
            dedcols_line = None
            if os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith("DEDCOLS:"):
                            dedcols_line = line.rstrip('\n')
            # Write UAC codes and preserve DEDCOLS line if present
            with open(settings_path, 'w', encoding='utf-8', newline='') as f:
                for code in uac_codes:
                    f.write(f"UAC:{code}\n")
                if dedcols_line:
                    f.write(dedcols_line + "\n")
            # Adjust rows to match new UAC code count
            while self.rows - 1 < new_count:
                self.add_row()
            while self.rows - 1 > new_count:
                self.remove_row()
            popup.destroy()
            self.refresh_aggregation()
        tk.Button(btns_frame, text="Apply", font=("Arial", 12, "bold"), width=12, command=on_apply).pack(side="left", padx=8)
        tk.Button(btns_frame, text="Cancel", font=("Arial", 11), width=10, command=popup.destroy).pack(side="left", padx=8)

    def refresh_aggregation(self):
        # If excel data is available in controller, re-aggregate and redraw
        if hasattr(self.controller, 'excel_data'):
            self.set_excel_aggregated(self.controller.excel_data)

    def _aggregate_excel_data(self, excel_data):
        """
        Aggregates excelImport data by PAP/UAC CODE, summing all numeric columns for each code.
        Returns a dict: {pap_code: [mo_salary, pera, gsis, phic, hdmf, other_deductions]}
        """
        from collections import defaultdict
        agg = defaultdict(lambda: [0, 0, 0, 0, 0, 0])
        # unsure, but this is for the default J - U columns (maybe lang)
        if hasattr(self, '_selected_other_deduction_cols') and self._selected_other_deduction_cols:
            other_deduction_cols = self._selected_other_deduction_cols
        else:
            other_deduction_cols = list(range(9, 21))  
        # Try to get header row and map columns
        if excel_data and all(isinstance(x, str) for x in excel_data[0]):
            headers = [h.strip().upper() for h in excel_data[0]]
            def find_col(name):
                try:
                    return headers.index(name)
                except ValueError:
                    return None
            idx_salary = find_col('SALARY')
            idx_pera = find_col('PERA')
            idx_gsis = find_col('GSIS SHARE')
            idx_phic = find_col('PHILHEALTH')
            idx_pagibig = find_col('PAGIBIG')
            idx_pagibig2 = find_col('PAGIBIG II')
            idx_pap = find_col('PAP / UAC CODE')
            data_rows = excel_data[1:]
        else:
            idx_salary = 2
            idx_pera = 3
            idx_gsis = 4
            idx_phic = 5
            idx_pagibig = 6
            idx_pagibig2 = 7
            idx_pap = 1
            data_rows = excel_data
        for row in data_rows:
            if len(row) < 22:
                continue
            pap = str(row[idx_pap]).strip() if idx_pap is not None and len(row) > idx_pap else ''
            # MO.SALARY
            try:
                mo_salary = float(row[idx_salary]) if idx_salary is not None and row[idx_salary] not in (None, "") else 0
            except:
                mo_salary = 0
            # PERA
            try:
                pera = float(row[idx_pera]) if idx_pera is not None and row[idx_pera] not in (None, "") else 0
            except:
                pera = 0
            # GSIS
            try:
                gsis = float(row[idx_gsis]) if idx_gsis is not None and row[idx_gsis] not in (None, "") else 0
            except:
                gsis = 0
            # PHIC
            try:
                phic = float(row[idx_phic]) if idx_phic is not None and row[idx_phic] not in (None, "") else 0
            except:
                phic = 0
            # HDMF (sum PAGIBIG and PAGIBIG II)
            hdmf = 0
            try:
                hdmf1 = float(row[idx_pagibig]) if idx_pagibig is not None and row[idx_pagibig] not in (None, "") else 0
            except:
                hdmf1 = 0
            try:
                hdmf2 = float(row[idx_pagibig2]) if idx_pagibig2 is not None and row[idx_pagibig2] not in (None, "") else 0
            except:
                hdmf2 = 0
            hdmf = hdmf1 + hdmf2
            # OTHER DEDUCTIONS (WTAX to UMD, or user selection)
            other = 0
            for idx in other_deduction_cols:
                try:
                    v = float(row[idx]) if len(row) > idx and row[idx] not in (None, "") else 0
                except:
                    v = 0
                other += v
            # Sum into agg
            agg[pap][0] += mo_salary
            agg[pap][1] += pera
            agg[pap][2] += gsis
            agg[pap][3] += phic
            agg[pap][4] += hdmf
            agg[pap][5] += other
        return agg

    def set_excel_aggregated(self, excel_data):
        """
        Combines all PAP/UAC CODE rows in excelImport and sets the aggregated data in the report table.
        Only rows matching the default_row_headers are shown, in the order of default_row_headers.
        """
        # --- NEW LOGIC: check if excel_data headers/cols match saved settings ---
        headers = []
        if excel_data and all(isinstance(x, str) for x in excel_data[0]):
            headers = [h.strip() for h in excel_data[0]]
        # Compare with saved deduction colnames
        match = False
        if self._deduction_colnames_saved and headers:
            # Only compare if both exist
            if len(self._deduction_colnames_saved) == len(headers):
                match = all(a == b for a, b in zip(self._deduction_colnames_saved, headers))
        if not match:
            # Reset to default deduction columns if structure does not match
            self._selected_other_deduction_cols = list(range(9, 21))  # J (9) to U (20) inclusive
            self._deduction_colnames_saved = headers
            # Optionally, notify user here (e.g., popup or print)
        agg = self._aggregate_excel_data(excel_data)
        headers_out = self.default_col_headers[:1] + ["MO.SALARY", "PERA AMOUNT", "GSIS", "PHIC", "HDMF", "OTHER DEDUCTIONS"]
        data = []
        for pap in self.default_row_headers:
            vals = agg.get(pap, [0,0,0,0,0,0])
            row = [f"{v:.2f}" for v in vals]
            data.append(row)
        # Pad/truncate to fit table size
        while len(data) < self.rows-1:
            data.append([""] * 6)
        data = data[:self.rows-1]
        self.set_aggregated_data(headers_out, data)
