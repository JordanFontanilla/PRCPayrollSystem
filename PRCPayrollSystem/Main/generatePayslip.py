import customtkinter as ctk
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as messagebox
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib import colors
from reportlab.lib.units import mm
import csv
import os
import sys
from PRCPayrollSystem.Main.resource_utils import resource_path
import re

class GeneratePayslipPage(ctk.CTkFrame):
    def __init__(self, parent, controller=None):
        super().__init__(parent, fg_color="white")
        self.controller = controller
        self.load_updated_fields()  # Load field adjustments on init
        self._build_ui()

    def load_updated_fields(self):
        """Load custom and removed fields from updatedFields.csv and apply to the payslip config."""
        updated_fields_path = resource_path("PRCPayrollSystem/settingsAndFields/updatedFields.csv")
        custom_map = {}
        custom_types = {}
        removed = set()
        if os.path.exists(updated_fields_path):
            with open(updated_fields_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                # Make fieldnames lowercase for robustness
                reader.fieldnames = [field.strip().lower() for field in reader.fieldnames] if reader.fieldnames else []
                for row in reader:
                    typ = row.get("type", "").strip().lower()
                    field = row.get("field", "").strip()
                    if typ == "custom" and field:
                        cols = [c.strip() for c in row.get("columns", "").split(',') if c.strip()]
                        field_type = row.get("fieldtype", "earnings").strip().lower()
                        custom_map[field] = cols
                        custom_types[field] = field_type
                    elif typ == "removed" and field:
                        removed.add(field.upper()) # Always store removed as uppercase
        self._custom_field_map = custom_map
        self._custom_field_types = custom_types
        self._removed_default_fields = removed

    def _build_ui(self):
        # Top bar
        topbar = ctk.CTkFrame(self, fg_color="white")
        topbar.pack(fill=ctk.X, pady=(5, 0))
        def go_to_main():
            if self.controller:
                self.controller.show_frame("ExcelImportPage")
        ctk.CTkButton(topbar, text='>', font=('Consolas', 18, 'bold'), fg_color="white", text_color="#0041C2", hover_color="#e6e6e6", width=40, height=40, corner_radius=20, command=go_to_main).pack(side=ctk.LEFT, padx=(10, 20))
        ctk.CTkButton(topbar, text='Load payslip records', fg_color="#0a47b1", text_color="white", hover_color="#003580", font=("Arial", 13, "bold"), corner_radius=16, width=180, height=36, command=self.load_payslip_records).pack(side=ctk.LEFT, padx=(0, 10))
        ctk.CTkButton(topbar, text='Download PDF', fg_color="#1877F2", text_color="white", hover_color="#1456a0", font=("Arial", 13, "bold"), corner_radius=16, width=160, height=36, command=self.download_pdf).pack(side=ctk.RIGHT, padx=(0, 10))
        ctk.CTkButton(topbar, text='Download all PDF', fg_color="#1877F2", text_color="white", hover_color="#1456a0", font=("Arial", 13, "bold"), corner_radius=16, width=160, height=36, command=self.download_all_pdf).pack(side=ctk.RIGHT, padx=(0, 10))
        ctk.CTkButton(topbar, text='Adjust payslip', fg_color="#0a47b1", text_color="white", hover_color="#003580", font=("Arial", 13, "bold"), corner_radius=16, width=160, height=36, command=self.show_adjust_payslip_popup).pack(side=ctk.LEFT, padx=(0, 10))

        # Main content frame
        content = ctk.CTkFrame(self, fg_color="#f5f7fa")
        content.pack(fill=ctk.BOTH, expand=True)

        # Left: Employee list
        left_frame = ctk.CTkFrame(content, fg_color="#f8fbff", border_width=0)
        left_frame.place(relx=0.01, rely=0.05, relwidth=0.18, relheight=0.9)
        ctk.CTkLabel(left_frame, text="Name:", font=("Montserrat", 15, "bold"), text_color="#0a47b1", fg_color="#e6eaff", anchor="w", height=38, corner_radius=8).pack(fill=ctk.X, pady=(8, 0), padx=8)
        self.emp_listbox = tk.Listbox(left_frame, font=("Montserrat", 13), height=12, activestyle='none', bg="#f8fbff", fg="#222", highlightthickness=0, bd=0, relief="flat", selectbackground="#dbeafe", selectforeground="#0a47b1")
        for i in range(7):
            self.emp_listbox.insert(tk.END, f"Name {i+1}")
        self.emp_listbox.pack(fill=ctk.BOTH, expand=True, padx=8, pady=8, ipadx=2, ipady=2)
        self.emp_listbox.configure(borderwidth=0, highlightbackground="#e6eaff", highlightcolor="#e6eaff")

        # Right: Payslip form (scrollable)
        payslip_frame = ctk.CTkFrame(content, fg_color="white", border_width=0)
        payslip_frame.place(relx=0.21, rely=0.05, relwidth=0.78, relheight=0.9)
        # Scrollable canvas (add horizontal scroll)
        self.canvas = ctk.CTkCanvas(payslip_frame, bg="white", highlightthickness=0)
        v_scroll = ctk.CTkScrollbar(payslip_frame, orientation="vertical", command=self.canvas.yview)
        h_scroll = ctk.CTkScrollbar(payslip_frame, orientation="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        payslip_frame.grid_rowconfigure(0, weight=1)
        payslip_frame.grid_columnconfigure(0, weight=1)
        # Frame inside canvas for the payslip
        self.inner_frame = ctk.CTkFrame(self.canvas, fg_color="white")
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", lambda e: self._update_scrollregion())
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
        # Payslip layout
        self._draw_payslip()

    def _normalize_name(self, name):
        # Convert to uppercase, remove punctuation/extra whitespace to create a consistent key.
        return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', '', str(name))).upper().strip()

    def _update_scrollregion(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        if self.canvas.bbox("all") and self.canvas.winfo_height() >= self.canvas.bbox("all")[3]:
            self.canvas.yview_moveto(0)
            self.canvas.configure(yscrollcommand=lambda *args: None)
        else:
            self.canvas.configure(yscrollcommand=self._canvas.master.children['!ctkscrollbar'].set)

    def get_current_earning_and_deduction_fields(self):
        # Returns (earning
        removed = set(getattr(self, '_removed_default_fields', set()))
        default_earning_labels = [
            "Basic Salary", "PERA"
        ]
        default_deduction_labels = [
            "Withholding Tax", "GSIS Employee Share", "PhilHealth Employee Share", "PAGIBIG Employee Share", "Landbank Salary Loan", "GSIS MPL", "GSIS MPL-LITE", "GSIS GFAL", "GSIS Policy Loan", "GSIS Emergency Loan", "GSIS CPL", "PAGIBIG Calamity Loan", "PAGIBIG MPL"
        ]
        custom_fields = list(getattr(self, '_custom_field_map', {}).keys())
        custom_types = getattr(self, '_custom_field_types', {})
        # Remove removed fields and custom fields from defaults
        earning_labels = [f for f in default_earning_labels if f.upper() not in removed and f not in custom_fields]
        deduction_labels = [f for f in default_deduction_labels if f.upper() not in removed and f not in custom_fields]
        # Add custom fields
        for f in custom_fields:
            t = custom_types.get(f, "earnings")
            if t == "earnings":
                earning_labels.append(f)
            else:
                deduction_labels.append(f)
        # Remove duplicates, preserve order
        earning_labels = list(dict.fromkeys(earning_labels))
        deduction_labels = list(dict.fromkeys(deduction_labels))
        return earning_labels, deduction_labels

    def _draw_payslip(self):
        self.load_updated_fields()  # Always reload adjustments before drawing
        # Clear all widgets from inner_frame to avoid artifacts when fields are removed/added
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        # Logo and header
        logo_path = None
        try:
            import os
            logo_path = resource_path("PRCPayrollSystem/Components/PRClogo.png")
            logo_img = Image.open(logo_path).resize((60, 60), Image.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(self.inner_frame, image=self.logo, bg="white", borderwidth=0)
            logo_label.grid(row=0, column=0, columnspan=6, pady=(10, 0), sticky="w")
        except Exception:
            pass
        ctk.CTkLabel(self.inner_frame, text="Professional Regulation Commission\nCordillera Administrative Region", font=("Arial", 15, "bold"), text_color="#222", fg_color="white").grid(row=0, column=1, columnspan=6, pady=(10, 0), sticky="w")
        # Name/Designation/Salary Grade
        ctk.CTkLabel(self.inner_frame, text="Name", font=("Arial", 12, "bold"), fg_color="white", anchor="w").grid(row=1, column=0, sticky="nsew", padx=2, pady=2)
        self.name_entry = ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120)
        self.name_entry.grid(row=1, column=1, sticky="nsew", padx=2, pady=2)
        ctk.CTkLabel(self.inner_frame, text="Designation", font=("Arial", 12, "bold"), fg_color="white", anchor="w").grid(row=1, column=2, sticky="nsew", padx=2, pady=2)
        self.designation_entry = ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120)
        self.designation_entry.grid(row=1, column=3, sticky="nsew", padx=2, pady=2)
        ctk.CTkLabel(self.inner_frame, text="Salary Grade", font=("Arial", 12, "bold"), fg_color="white", anchor="w").grid(row=1, column=4, sticky="nsew", padx=2, pady=2)
        self.salary_grade_entry = ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=80)
        self.salary_grade_entry.grid(row=1, column=5, sticky="nsew", padx=2, pady=2)

        # Table header
        ctk.CTkLabel(self.inner_frame, text="Earnings", font=("Arial", 12, "bold"), fg_color="#e6e6e6", anchor="center").grid(row=2, column=0, columnspan=2, sticky="nsew", padx=1, pady=1)
        ctk.CTkLabel(self.inner_frame, text="Deductions", font=("Arial", 12, "bold"), fg_color="#e6e6e6", anchor="center").grid(row=2, column=2, columnspan=2, sticky="nsew", padx=1, pady=1)
        # Remove extra empty header columns for better alignment
        # Table columns
        ctk.CTkLabel(self.inner_frame, text="", fg_color="#e6e6e6").grid(row=2, column=3, columnspan=2, sticky="nsew", padx=1, pady=1)
        ctk.CTkLabel(self.inner_frame, text="", fg_color="#e6e6e6").grid(row=2, column=4, columnspan=2, sticky="nsew", padx=1, pady=1)



        # Get earning and deduction fields using the new helper
        earning_labels, deduction_labels = self.get_current_earning_and_deduction_fields()
        # Padding to keep table shape and alignment
        max_rows = max(len(earning_labels), len(deduction_labels))
        earning_labels += [""] * (max_rows - len(earning_labels))
        deduction_labels += [""] * (max_rows - len(deduction_labels))
        for i in range(max_rows):
            e_text = earning_labels[i]
            d_text = deduction_labels[i]
            # Only create widgets for non-blank fields
            if e_text.strip():
                ctk.CTkLabel(self.inner_frame, text=e_text, font=("Arial", 12), fg_color="white", anchor="w", width=160).grid(row=3+i, column=0, sticky="nsew", padx=1, pady=1)
                ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120).grid(row=3+i, column=1, sticky="nsew", padx=1, pady=1)
            # Deductions remain unchanged
            if d_text.strip():
                ctk.CTkLabel(self.inner_frame, text=d_text, font=("Arial", 12), fg_color="white", anchor="w", width=220).grid(row=3+i, column=2, sticky="nsew", padx=1, pady=1)
                ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120).grid(row=3+i, column=3, sticky="nsew", padx=1, pady=1)
        # Totals and netpay (as table rows)
        row_offset = 3 + max_rows
        ctk.CTkLabel(self.inner_frame, text="Total Earnings", font=("Arial", 12, "bold"), fg_color="white", anchor="w").grid(row=row_offset, column=0, sticky="w", padx=(10, 0), pady=(10, 0))
        ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120).grid(row=row_offset, column=1, sticky="w", padx=(0, 0), pady=(10, 0))
        ctk.CTkLabel(self.inner_frame, text="Total Deductions", font=("Arial", 12, "bold"), fg_color="white", anchor="w").grid(row=row_offset, column=2, sticky="w", padx=(10, 0), pady=(10, 0))
        ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120).grid(row=row_offset, column=3, sticky="w", padx=(0, 0), pady=(10, 0))
        # Netpay rows
        ctk.CTkLabel(self.inner_frame, text="Netpay (1st half)", font=("Arial", 12, "bold"), fg_color="white", anchor="w").grid(row=row_offset+1, column=0, sticky="w", padx=(10, 0), pady=(10, 0))
        ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120).grid(row=row_offset+1, column=1, sticky="w", padx=(0, 0), pady=(10, 0))
        ctk.CTkLabel(self.inner_frame, text="Netpay (2nd half)", font=("Arial", 12, "bold"), fg_color="white", anchor="w").grid(row=row_offset+1, column=2, sticky="w", padx=(10, 0), pady=(10, 0))
        ctk.CTkEntry(self.inner_frame, font=("Arial", 12), fg_color="white", border_width=0, corner_radius=0, width=120).grid(row=row_offset+1, column=3, sticky="w", padx=(0, 0), pady=(10, 0))
        # ...existing code for adjust fields: apply the same border_width=0, corner_radius=0, fg_color="white" to all CTkEntry widgets...

    def set_employee_info(self, name, designation, salary_grade):
        # Set the values in the payslip table entries for Name, Designation, Salary Grade
        if hasattr(self, 'name_entry'):
            self.name_entry.delete(0, 'end')
            self.name_entry.insert(0, name)
        if hasattr(self, 'designation_entry'):
            self.designation_entry.delete(0, 'end')
            self.designation_entry.insert(0, designation)
        if hasattr(self, 'salary_grade_entry'):
            self.salary_grade_entry.delete(0, 'end')
            self.salary_grade_entry.insert(0, salary_grade)

    def set_employee_names(self, names):
        # Clear the listbox
        self.emp_listbox.delete(0, 'end')
        # Insert new names with enumeration
        for idx, name in enumerate(names, 1):
            self.emp_listbox.insert('end', f"{idx}. {name}")
        # Store the names for lookup
        self._employee_names = names
        # Bind selection event
        self.emp_listbox.bind('<<ListboxSelect>>', self._on_employee_select)

    def _on_employee_select(self, event):
        # Get selected index
        selection = event.widget.curselection()
        if not selection:
            return
        idx = selection[0]
        # Lookup details from ImportEmployeePage table if available
        if hasattr(self.controller, 'frames') and 'ImportEmployeePage' in self.controller.frames:
            import_page = self.controller.frames['ImportEmployeePage']
            # Find the row in the table corresponding to the selected index (skip header)
            row_idx = idx + 1
            name = import_page.table.cells.get((row_idx, 0)).get() if import_page.table.cells.get((row_idx, 0)) else ''
            designation = import_page.table.cells.get((row_idx, 1)).get() if import_page.table.cells.get((row_idx, 1)) else ''
            salary_grade = import_page.table.cells.get((row_idx, 2)).get() if import_page.table.cells.get((row_idx, 2)) else ''
            self.set_employee_info(name, designation, salary_grade)

    def load_payslip_records(self):
        # 1. Load master employee list from ImportEmployeePage (the single source of truth for employees)
        master_names = []
        employee_details = {}
        if hasattr(self.controller, 'frames') and 'ImportEmployeePage' in self.controller.frames:
            import_page = self.controller.frames['ImportEmployeePage']
            table = getattr(import_page, 'table', None)
            if table:
                # Assuming Name, Designation, Salary Grade are the first three columns
                for r in range(1, table.rows):  # Skip header
                    name_cell = table.cells.get((r, 0))
                    name = name_cell.get().strip() if name_cell and name_cell.get() else ''
                    if name:
                        master_names.append(name)
                        designation_cell = table.cells.get((r, 1))
                        salary_grade_cell = table.cells.get((r, 2))
                        designation = designation_cell.get().strip() if designation_cell and designation_cell.get() else ''
                        salary_grade = salary_grade_cell.get().strip() if salary_grade_cell and salary_grade_cell.get() else ''
                        employee_details[name] = {'designation': designation, 'salary_grade': salary_grade}
        
        self._employee_details = employee_details # Store for reuse

        # 2. Load payslip financial data from the imported Excel file into a temporary dictionary
        raw_payslip_data = {}
        excel_table_loaded = False
        if hasattr(self.controller, 'frames') and 'ExcelImportPage' in self.controller.frames:
            excel_page = self.controller.frames['ExcelImportPage']
            if hasattr(excel_page, 'get_table_data'):
                table_data = excel_page.get_table_data()
                if table_data and len(table_data) > 1:
                    excel_table_loaded = True
                    header_row = [str(h).strip() for h in table_data[0]]
                    name_idx = -1
                    # Find name column index case-insensitively
                    for i, h in enumerate(header_row):
                        if h.strip().upper() in ("NAME", "EMPLOYEE NAME"):
                            name_idx = i
                            break
                    if name_idx == -1:
                         messagebox.showerror("Error", "Could not find 'Name' or 'Employee Name' column in the imported Excel file.")
                         return

                    # Map payslip fields to their indices in Excel
                    field_map = {
                        "BASIC SALARY": ["SALARY", "MO.SALARY", "MONTHLY SALARY"],
                        "PERA": ["PERA"],
                        "WITHHOLDING TAX": ["WTAX", "WITHHOLDING TAX"],
                        "GSIS EMPLOYEE SHARE": ["GSIS SHARE", "GSIS EMPLOYEE SHARE"],
                        "PHILHEALTH EMPLOYEE SHARE": ["PHILHEALTH", "PHIC", "PHILHEALTH EMPLOYEE SHARE"],
                        "PAGIBIG EMPLOYEE SHARE": ["PAGIBIG 1", "PAGIBIG 2", "PAG-IBIG 1", "PAG-IBIG 2", "HDMF", "PAGIBIG", "PAGIBIG EMPLOYEE SHARE"],
                        "LANDBANK SALARY LOAN": ["LANDBANK SALARY LOAN", "LB SALARY LOAN"],
                        "GSIS MPL": ["GSIS MPL", "GSIS_MPL"],
                        "GSIS MPL-LITE": ["GSIS MPL-LITE", "MP-LITE"],
                        "GSIS GFAL": ["GSIS GFAL", "GFAL"],
                        "GSIS POLICY LOAN": ["GSIS POLICY LOAN", "PLREG"],
                        "GSIS EMERGENCY LOAN": ["GSIS EMERGENCY LOAN", "EMRGYLN"],
                        "GSIS CPL": ["GSIS CPL", "CPL"],
                        "PAGIBIG CALAMITY LOAN": ["PAGIBIG CALAMITY LOAN", "CL-MPL"],
                        "PAGIBIG MPL": ["PAGIBIG MPL", "MPL"],
                        "NETPAY (1ST HALF)": ["SAL_WK1"],
                        "NETPAY (2ND HALF)": ["SAL_WK2"]
                    }
                    # Find indices for each field
                    field_indices = {}
                    pagibig_indices = []
                    for i, h in enumerate(header_row):
                        h_upper = h.strip().upper()
                        if "PAGIBIG" in h_upper:
                            pagibig_indices.append(i)
                        for key, aliases in field_map.items():
                            if key == "PAGIBIG EMPLOYEE SHARE": continue
                            for alias in aliases:
                                if h_upper == alias.upper():
                                    field_indices[key] = i
                    # --- Custom field index mapping ---
                    custom_map = getattr(self, '_custom_field_map', {})
                    custom_indices = {}
                    for field, excel_names in custom_map.items():
                        indices = []
                        for col in excel_names:
                            col = col.strip()
                            if '-' in col:
                                # Range, e.g. C-E or 2-5
                                try:
                                    if col[0].isalpha():
                                        # Letter range (Excel style)
                                        start, end = col.split('-')
                                        start_idx = ord(start.upper()) - ord('A')
                                        end_idx = ord(end.upper()) - ord('A')
                                        indices.extend(list(range(start_idx, end_idx+1)))
                                    else:
                                        # Numeric range
                                        start, end = map(int, col.split('-'))
                                        indices.extend(list(range(start-1, end)))
                                except Exception:
                                    continue
                            else:
                                # Try to match by header name first
                                try:
                                    idx = header_row.index(col)
                                    indices.append(idx)
                                except ValueError:
                                    # Try as column letter
                                    if col.isalpha() and len(col) == 1:
                                        indices.append(ord(col.upper()) - ord('A'))
                                    else:
                                        try:
                                            indices.append(int(col)-1)
                                        except Exception:
                                            pass
                        custom_indices[field] = indices
                    # ---
                    for row in table_data[1:]:
                        excel_name = row[name_idx].strip() if name_idx is not None and len(row) > name_idx else ''
                        if excel_name:
                            row_data = {}
                            for key in field_map.keys():
                                if key in ("NAME", "DESIGNATION", "SALARY GRADE"):
                                    continue
                                if key == "PAGIBIG EMPLOYEE SHARE":
                                    pagibig_sum = 0.0
                                    for idx in pagibig_indices:
                                        try:
                                            val = row[idx]
                                            pagibig_sum += float(val) if val not in (None, '', 'nan') else 0.0
                                        except (ValueError, TypeError, IndexError): pass
                                    row_data[key] = f"{pagibig_sum:.2f}" if pagibig_sum != 0 else ''
                                else:
                                    idx = field_indices.get(key)
                                    row_data[key] = row[idx] if idx is not None and len(row) > idx else ''
                            # --- Custom fields value extraction ---
                            for field, indices in custom_indices.items():
                                val_sum = 0.0
                                has_value = False
                                for idx in indices:
                                    try:
                                        v = row[idx]
                                        if v not in (None, '', 'nan'):
                                            val_sum += float(v)
                                            has_value = True
                                    except (ValueError, TypeError, IndexError): continue
                                if has_value:
                                    row_data[field] = f"{val_sum:.2f}" if len(indices) > 1 else str(val_sum) if val_sum != 0 else ''
                                else:
                                    row_data[field] = ''
                            raw_payslip_data[excel_name] = row_data
        
        # 3. Reconcile financial data with master employee list.
        final_payslip_data = {}
        # Create a map from a normalized excel name to its original data.
        raw_payslip_data_map = { self._normalize_name(name): data for name, data in raw_payslip_data.items() }

        for name in master_names:
            normalized_master_name = self._normalize_name(name)
            if normalized_master_name in raw_payslip_data_map:
                # If a direct normalized match is found, use it.
                final_payslip_data[name] = raw_payslip_data_map[normalized_master_name]
            else:
                # If no financial data is found, the employee will have an empty payslip.
                final_payslip_data[name] = {}
        
        # 4. Update UI with reconciled data
        self.set_employee_names(master_names)
        self._payslip_data = final_payslip_data
        self.emp_listbox.bind('<<ListboxSelect>>', self._on_payslip_record_select)

        # Always load the first employee's payslip data if available
        if master_names:
            self.emp_listbox.selection_clear(0, 'end')
            self.emp_listbox.selection_set(0)
            self.emp_listbox.activate(0)
            self._draw_payslip()  # Redraw the payslip layout to reflect new/removed fields
            # Get the first employee's info from our loaded details
            name = master_names[0]
            details = self._employee_details.get(name, {})
            designation = details.get('designation', '')
            salary_grade = details.get('salary_grade', '')
            self.set_employee_info(name, designation, salary_grade)
            self._on_payslip_record_select_from_name(name)
            import tkinter.messagebox as messagebox
            if excel_table_loaded:
                pass
            else:
                messagebox.showwarning("ExcelImport Missing", "No Excel file loaded. Please go to 'Generate Payslip/Reports', import an Excel file, then try again.")

    def _on_payslip_record_select(self, event):
        selection = event.widget.curselection()
        if not selection:
            return
        idx = selection[0]
        name = self.emp_listbox.get(idx)
        if '. ' in name:
            name = name.split('. ', 1)[1]
        self._on_payslip_record_select_from_name(name)

    def _on_payslip_record_select_from_name(self, name):
        # Lookup payslip data and fill payslip fields
        if hasattr(self, '_payslip_data') and name in self._payslip_data:
            self.fill_payslip_fields(self._payslip_data[name])
        # Always update name, designation, and salary grade fields from stored details
        designation = ""
        salary_grade = ""
        if hasattr(self, '_employee_details') and name in self._employee_details:
            details = self._employee_details[name]
            designation = details.get('designation', '')
            salary_grade = details.get('salary_grade', '')
        self.set_employee_info(name, designation, salary_grade)

    def fill_payslip_fields(self, row_data):
        # Only fill the payslip fields (salary to MPL) in the payslip document, do not touch designation or salary grade
        # Map payslip fields to their row/column in the payslip layout
        # Build dynamic field map for custom fields and removed fields
        removed = set(getattr(self, '_removed_default_fields', set()))
        custom_fields = list(getattr(self, '_custom_field_map', {}).keys())
        custom_types = getattr(self, '_custom_field_types', {})
        # Default field order
        earning_fields = ["BASIC SALARY", "PERA"]
        deduction_fields = [
            "WITHHOLDING TAX", "GSIS EMPLOYEE SHARE", "PHILHEALTH EMPLOYEE SHARE", "PAGIBIG EMPLOYEE SHARE", "LANDBANK SALARY LOAN", "GSIS MPL", "GSIS MPL-LITE", "GSIS GFAL", "GSIS POLICY LOAN", "GSIS EMERGENCY LOAN", "GSIS CPL", "PAGIBIG CALAMITY LOAN", "PAGIBIG MPL"
        ]
        earning_fields = [f for f in earning_fields if f not in removed]
        deduction_fields = [f for f in deduction_fields if f not in removed]
        for f in custom_fields:
            t = custom_types.get(f, "earnings")
            if t == "earnings":
                earning_fields.append(f)
            else:
                deduction_fields.append(f)
        # Map fields to rows/columns
        field_map = {}
        max_rows = max(len(earning_fields), len(deduction_fields))
        for i in range(max_rows):
            if i < len(earning_fields) and earning_fields[i]:
                field_map[earning_fields[i]] = (3+i, 1)
            if i < len(deduction_fields) and deduction_fields[i]:
                field_map[deduction_fields[i]] = (3+i, 3)
        # Clear all payslip fields first
        for (row, col) in field_map.values():
            for widget in self.inner_frame.winfo_children():
                info = widget.grid_info()
                if info.get('row') == row and info.get('column') == col and isinstance(widget, ctk.CTkEntry):
                    widget.delete(0, 'end')
        # Now fill with new data and sum earnings/deductions
        total_earnings = 0.0
        total_deductions = 0.0
        for field, (row, col) in field_map.items():
            value = row_data.get(field, "")
            for widget in self.inner_frame.winfo_children():
                info = widget.grid_info()
                if info.get('row') == row and info.get('column') == col and isinstance(widget, ctk.CTkEntry):
                    # Format with commas for thousands if numeric
                    try:
                        v = float(value)
                        widget.insert(0, f"{v:,.2f}")
                    except (ValueError, TypeError):
                        widget.insert(0, value)
            # Sum for totals
            try:
                v = float(value) if value not in (None, '', 'nan') else 0.0
            except (ValueError, TypeError):
                v = 0.0
            if field in earning_fields:
                total_earnings += v
            elif field in deduction_fields:
                total_deductions += v
        # Set totals in the payslip
        row_offset = 3 + max_rows
        for widget in self.inner_frame.winfo_children():
            info = widget.grid_info()
            if info.get('row') == row_offset and info.get('column') == 1 and isinstance(widget, ctk.CTkEntry):
                widget.delete(0, 'end')
                widget.insert(0, f"{total_earnings:,.2f}" if total_earnings != 0 else '')
            if info.get('row') == row_offset and info.get('column') == 3 and isinstance(widget, ctk.CTkEntry):
                widget.delete(0, 'end')
                widget.insert(0, f"{total_deductions:,.2f}" if total_deductions != 0 else '')
        # Netpay fields (case-insensitive lookup)
        row_data_ci = {str(k).strip().upper(): v for k, v in row_data.items()}
        netpay1 = row_data_ci.get("NETPAY (1ST HALF)", "")
        netpay2 = row_data_ci.get("NETPAY (2ND HALF)", "")
        for widget in self.inner_frame.winfo_children():
            info = widget.grid_info()
            if info.get('row') == row_offset+1 and info.get('column') == 1 and isinstance(widget, ctk.CTkEntry):
                widget.delete(0, 'end')
                try:
                    v = float(netpay1)
                    widget.insert(0, f"{v:,.2f}")
                except (ValueError, TypeError):
                    widget.insert(0, netpay1)
            if info.get('row') == row_offset+1 and info.get('column') == 3 and isinstance(widget, ctk.CTkEntry):
                widget.delete(0, 'end')
                try:
                    v = float(netpay2)
                    widget.insert(0, f"{v:,.2f}")
                except (ValueError, TypeError):
                    widget.insert(0, netpay2)

    def download_pdf(self):
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.lib.utils import ImageReader
        import os
        import tkinter as tk
        from tkinter import messagebox, filedialog as fd
        self.load_updated_fields()  # Always reload adjustments before generating
        file_path = fd.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Payslip as PDF")
        if not file_path:
            return
        # --- Get selected employee robustly ---
        name = getattr(self, '_selected_employee', None)
        if not name and hasattr(self, 'emp_listbox'):
            selection = self.emp_listbox.curselection()
            if selection:
                idx = selection[0]
                if hasattr(self, '_employee_names') and idx < len(self._employee_names):
                    name = self._employee_names[idx]
        # --- Ensure payslip data is loaded ---
        payslip_data = getattr(self, '_payslip_data', None)
        if payslip_data is None or not payslip_data:
            # Try to load payslip records if not loaded
            if hasattr(self, 'load_payslip_records'):
                self.load_payslip_records()
                payslip_data = getattr(self, '_payslip_data', None)
        if not name or not payslip_data or name not in payslip_data:
            messagebox.showerror("Error", "No payslip data loaded.")
            return
        c = pdf_canvas.Canvas(file_path, pagesize=A4)
        width, height = A4
        margin = 20 * mm
        logo_height = 18*mm
        logo_width = 18*mm
        logo_x = margin
        logo_y = height - margin - logo_height + 10
        header_y = logo_y - 2
        # --- Use the same field logic as the UI ---
        earning_labels, deduction_labels = self.get_current_earning_and_deduction_fields()
        # --- Case-insensitive mapping for field lookups ---
        row_data = payslip_data.get(name, {})
        row_data_ci = {str(k).strip().upper(): v for k, v in row_data.items()}
        earning_vals = [row_data_ci.get(label.strip().upper(), "0.00") if label.strip() else "" for label in earning_labels]
        deduction_vals = [row_data_ci.get(label.strip().upper(), "0.00") if label.strip() else "" for label in deduction_labels]
        # Padding for uneven columns
        max_rows = max(len(earning_labels), len(deduction_labels))
        earning_labels_pad = earning_labels + [""] * (max_rows - len(earning_labels))
        earning_vals_pad = earning_vals + [""] * (max_rows - len(earning_vals))
        deduction_labels_pad = deduction_labels + [""] * (max_rows - len(deduction_labels))
        deduction_vals_pad = deduction_vals + [""] * (max_rows - len(deduction_vals))
        def safe_float(val):
            try:
                return float(str(val).replace(",", ""))
            except Exception:
                return 0.0
        total_earnings = sum(safe_float(v) for v in earning_vals if v not in (None, ""))
        total_deductions = sum(safe_float(v) for v in deduction_vals if v not in (None, ""))
        netpay1 = row_data_ci.get("NETPAY (1ST HALF)", "")
        netpay2 = row_data_ci.get("NETPAY (2ND HALF)", "")
        y = height - margin
        try:
            logo_path = resource_path("PRCPayrollSystem/Components/PRClogo.png")
            c.drawImage(ImageReader(logo_path), logo_x, logo_y, logo_width, logo_height, mask='auto')
        except Exception:
            pass
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(width/2, header_y, "Professional Regulation Commission")
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString(width/2, header_y - 18, "Cordillera Administrative Region")
        y = header_y - 40
        c.setFont("Helvetica", 12)
        c.drawString(margin, y, f"Name: {name}")
        y -= 16
        # Always get designation and salary grade from stored details
        designation = ""
        salary_grade = ""
        if hasattr(self, '_employee_details') and name in self._employee_details:
            details = self._employee_details[name]
            designation = details.get('designation', '')
            salary_grade = details.get('salary_grade', '')
        c.drawString(margin, y, f"Designation: {designation}")
        c.drawRightString(width - margin, y, f"Salary Grade: {salary_grade}")
        y -= 28
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margin, y, "Earnings")
        # Adjusted x-coordinates for better alignment
        deductions_header_x = margin + 250  # Move header further right
        deductions_value_x = margin + 470  # Move value further right
        c.drawString(deductions_header_x, y, "Deductions")
        y -= 8
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        c.line(margin, y, width-margin, y)
        y -= 16
        c.setFont("Helvetica", 10)
        for i in range(max_rows):
            # Format numbers with comma for thousands
            e_label = earning_labels_pad[i]
            e_val = earning_vals_pad[i]
            d_label = deduction_labels_pad[i]
            d_val = deduction_vals_pad[i]
            def fmt_num(val, label=None):
                # Only show 0.00 if the label is not blank
                if label is not None and not label.strip():
                    return ""
                try:
                    v = float(str(val).replace(',', ''))
                    return f"{v:,.2f}"
                except Exception:
                    return str(val) if val else "0.00"
            c.drawString(margin, y, e_label)
            c.drawRightString(margin+140, y, fmt_num(e_val, e_label))
            c.drawString(deductions_header_x, y, d_label)
            c.drawRightString(deductions_value_x, y, fmt_num(d_val, d_label))
            y -= 14
        # Update total/netpay lines to use comma formatting as well
        c.setFont("Helvetica-Bold", 11)
        c.drawString(margin, y, "Total Earnings:")
        c.drawRightString(margin+140, y, f"{total_earnings:,.2f}")
        c.drawString(deductions_header_x, y, "Total Deductions:")
        c.drawRightString(deductions_value_x, y, f"{total_deductions:,.2f}")
        y -= 22
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margin, y, f"Netpay (1st half): {fmt_num(netpay1)}")
        y -= 16
        c.drawString(margin, y, f"Netpay (2nd half): {fmt_num(netpay2)}")
        c.setFont("Helvetica-Oblique", 9)
        c.setFillColor(colors.grey)
        c.drawString(margin, 30, f"Generated by PRC Payroll System - v1.0")
        c.save()
        messagebox.showinfo("PDF Saved", f"Payslip PDF saved to:\n{file_path}")
        # Save a copy in pastPayslips folder
        try:
            import shutil
            past_payslips_dir = resource_path("PRCPayrollSystem/pastPayslips")
            if not os.path.exists(past_payslips_dir):
                os.makedirs(past_payslips_dir)
            # Check if there are already 100 or more payslips
            payslip_files = [f for f in os.listdir(past_payslips_dir) if f.lower().endswith('.pdf')]
            if len(payslip_files) >= 100:
                from tkinter import messagebox
                messagebox.showwarning(
                    "Payslip Limit Reached",
                    "The pastPayslips folder already contains 100 payslips. Please delete a payslip first using the delete function in the History page before saving new ones."
                )
                return  # Do not save if limit reached
            # Use employee name and date for filename
            from datetime import datetime
            safe_name = name.replace(' ', '_').replace('/', '_')
            dt_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            dest_path = os.path.join(past_payslips_dir, f"{safe_name}_payslip_{dt_str}.pdf")
            shutil.copy(file_path, dest_path)
        except Exception as e:
            print(f"Failed to save payslip copy: {e}")
        # After saving the PDF, refresh the HistoryPage PDF list
        if self.controller:
            history_page = self.controller.get_page('HistoryPage')
            if history_page:
                history_page.refresh()

    def download_all_pdf(self):
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.lib.utils import ImageReader
        import os
        import tkinter as tk
        from tkinter import messagebox, filedialog as fd
        self.load_updated_fields()  # Always reload adjustments before generating
        file_path = fd.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save All Payslips as PDF")
        if not file_path:
            return
        names = getattr(self, '_employee_names', [])
        payslip_data = getattr(self, '_payslip_data', {})
        if not names or not payslip_data:
            messagebox.showerror("Error", "No payslip data loaded.")
            return
        c = pdf_canvas.Canvas(file_path, pagesize=A4)
        width, height = A4
        margin = 20 * mm
        logo_height = 18*mm
        logo_width = 18*mm
        logo_x = margin
        logo_y = height - margin - logo_height + 10
        header_y = logo_y - 2
        earning_labels, deduction_labels = self.get_current_earning_and_deduction_fields()
        max_rows = max(len(earning_labels), len(deduction_labels))
        earning_labels_pad = earning_labels + [""] * (max_rows - len(earning_labels))
        deduction_labels_pad = deduction_labels + [""] * (max_rows - len(deduction_labels))
        for idx in range(0, len(names), 2):
            # Draw up to two payslips per page
            for payslip_num in range(2):
                if idx + payslip_num >= len(names):
                    break
                name = names[idx + payslip_num]
                row_data = payslip_data.get(name, {})
                row_data_ci = {str(k).strip().upper(): v for k, v in row_data.items()}
                earning_vals = [row_data_ci.get(label.strip().upper(), "0.00") if label.strip() else "" for label in earning_labels]
                deduction_vals = [row_data_ci.get(label.strip().upper(), "0.00") if label.strip() else "" for label in deduction_labels]
                earning_vals_pad = earning_vals + [""] * (max_rows - len(earning_vals))
                deduction_vals_pad = deduction_vals + [""] * (max_rows - len(deduction_vals))
                def safe_float(val):
                    try:
                        return float(str(val).replace(",", ""))
                    except Exception:
                        return 0.0
                total_earnings = sum(safe_float(v) for v in earning_vals if v not in (None, ""))
                total_deductions = sum(safe_float(v) for v in deduction_vals if v not in (None, ""))
                netpay1 = row_data_ci.get("NETPAY (1ST HALF)", "")
                netpay2 = row_data_ci.get("NETPAY (2ND HALF)", "")
                # Calculate payslip vertical space
                payslip_height = (height - 2*margin) / 2
                if payslip_num == 0:
                    y = height - margin
                    header_y = y - 10
                else:
                    y = height - margin - payslip_height - 15*mm  # Lower the second payslip further
                    header_y = y - 10
                # Draw logo for each payslip
                try:
                    logo_path = resource_path("PRCPayrollSystem/Components/PRClogo.png")
                    c.drawImage(ImageReader(logo_path), logo_x, header_y + 10, logo_width, logo_height, mask='auto')
                except Exception:
                    pass
                # Draw payslip content
                c.setFont("Helvetica-Bold", 20)
                c.setFillColor(colors.black)
                c.drawCentredString(width/2, header_y, "Professional Regulation Commission")
                c.setFont("Helvetica-Bold", 13)
                c.drawCentredString(width/2, header_y - 18, "Cordillera Administrative Region")
                y = header_y - 40
                c.setFont("Helvetica", 12)
                c.drawString(margin, y, f"Name: {name}")
                y -= 16
                designation = ""
                salary_grade = ""
                # Always get designation and salary grade from stored details
                if hasattr(self, '_employee_details') and name in self._employee_details:
                    details = self._employee_details[name]
                    designation = details.get('designation', '')
                    salary_grade = details.get('salary_grade', '')
                c.drawString(margin, y, f"Designation: {designation}")
                c.drawRightString(width - margin, y, f"Salary Grade: {salary_grade}")
                y -= 28
                c.setFont("Helvetica-Bold", 12)
                c.drawString(margin, y, "Earnings")
                # Adjusted x-coordinates for better alignment
                deductions_header_x = margin + 250  # Move header further right
                deductions_value_x = margin + 470   # Move value further right
                c.drawString(deductions_header_x, y, "Deductions")
                y -= 8
                c.setStrokeColor(colors.black)
                c.setLineWidth(1)
                c.line(margin, y, width-margin, y)
                y -= 16
                c.setFont("Helvetica", 10)
                for i in range(max_rows):
                    # Format numbers with comma for thousands
                    e_label = earning_labels_pad[i]
                    e_val = earning_vals_pad[i]
                    d_label = deduction_labels_pad[i]
                    d_val = deduction_vals_pad[i]
                    def fmt_num(val, label=None):
                        # Only show 0.00 if the label is not blank
                        if label is not None and not label.strip():
                            return ""
                        try:
                            v = float(str(val).replace(',', ''))
                            return f"{v:,.2f}"
                        except Exception:
                            return str(val) if val else "0.00"
                    c.drawString(margin, y, e_label)
                    c.drawRightString(margin+140, y, fmt_num(e_val, e_label))
                    c.drawString(deductions_header_x, y, d_label)
                    c.drawRightString(deductions_value_x, y, fmt_num(d_val, d_label))
                    y -= 14
                # Update total/netpay lines to use comma formatting as well
                c.setFont("Helvetica-Bold", 11)
                c.drawString(margin, y, "Total Earnings:")
                c.drawRightString(margin+140, y, f"{total_earnings:,.2f}")
                c.drawString(deductions_header_x, y, "Total Deductions:")
                c.drawRightString(deductions_value_x, y, f"{total_deductions:,.2f}")
                y -= 22
                c.setFont("Helvetica-Bold", 12)
                c.drawString(margin, y, f"Netpay (1st half): {fmt_num(netpay1)}")
                y -= 16
                c.drawString(margin, y, f"Netpay (2nd half): {fmt_num(netpay2)}")
                # Draw footer only for the last payslip on the page or last payslip overall
                if payslip_num == 1 or (idx + payslip_num == len(names) - 1):
                    c.setFont("Helvetica-Oblique", 9)
                    c.setFillColor(colors.grey)
                    c.drawString(margin, 30, f"Generated by PRC Payroll System - v1.0")
                c.setFillColor(colors.black)  # Reset color for next payslip
            if idx + 2 < len(names):
                c.showPage()  # New page for next two payslips
        c.save()
        messagebox.showinfo("PDF Saved", f"All payslips PDF saved to:\n{file_path}")
        # Save a copy in pastPayslips folder
        try:
            import shutil
            past_payslips_dir = resource_path("PRCPayrollSystem/pastPayslips")
            if not os.path.exists(past_payslips_dir):
                os.makedirs(past_payslips_dir)
            from datetime import datetime
            dt_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            dest_path = os.path.join(past_payslips_dir, f"all_payslips_{dt_str}.pdf")
            shutil.copy(file_path, dest_path)
        except Exception as e:
            print(f"Failed to save all payslips copy: {e}")
        # After saving the PDFs, refresh the HistoryPage PDF list
        if self.controller:
            history_page = self.controller.get_page('HistoryPage')
            if history_page:
                history_page.refresh()

    def show_adjust_payslip_popup(self):
        popup = tk.Toplevel(self)
        popup.title("Adjust Payslip Fields")
        popup.geometry("480x350")
        popup.grab_set()
        tk.Label(popup, text="Adjust Payslip Fields", font=("Arial", 14, "bold")).pack(pady=(18, 8))
        btn_add = tk.Button(popup, text="Add Field", font=("Arial", 12), width=18, command=lambda: [popup.destroy(), self.show_add_payslip_field_popup()])
        btn_add.pack(pady=10)
        btn_remove = tk.Button(popup, text="Remove Field", font=("Arial", 12), width=18, command=lambda: [popup.destroy(), self.show_remove_payslip_field_popup()])
        btn_remove.pack(pady=10)
        tk.Button(popup, text="Close", font=("Arial", 11), width=10, command=popup.destroy).pack(pady=10)

    def save_updated_fields(self):
        """Save custom and removed fields to updatedFields.csv and update payslipSettings.csv."""
        updated_fields_path = resource_path("PRCPayrollSystem/settingsAndFields/updatedFields.csv")
        payslip_settings_path = resource_path("PRCPayrollSystem/settingsAndFields/payslipSettings.csv")
        custom_map = getattr(self, '_custom_field_map', {})
        custom_types = getattr(self, '_custom_field_types', {})
        removed = getattr(self, '_removed_default_fields', set())
        # Save to updatedFields.csv (for UI logic)
        with open(updated_fields_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Type", "Field", "Columns", "FieldType"])
            for field, cols in custom_map.items():
                field_type = custom_types.get(field, "earnings")
                writer.writerow(["custom", field, ','.join(cols), field_type])
            for field in removed:
                writer.writerow(["removed", field, "", ""])
        # Save to payslipSettings.csv (for persistent config)
        # Compose the current field list (order, type, excel columns)
        earning_labels, deduction_labels = self.get_current_earning_and_deduction_fields()
        # Build a list of tuples: (Field, Type, Columns)
        field_rows = []
        for field in earning_labels:
            if field in custom_map:
                field_type = custom_types.get(field, "earnings")
                cols = ','.join(custom_map[field])
                field_rows.append([field, field_type, cols])
            else:
                field_rows.append([field, "earnings", ""])
        for field in deduction_labels:
            if field in custom_map:
                field_type = custom_types.get(field, "deductions")
                cols = ','.join(custom_map[field])
                field_rows.append([field, field_type, cols])
            else:
                field_rows.append([field, "deductions", ""])
        with open(payslip_settings_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Field", "Type", "Columns"])
            for row in field_rows:
                writer.writerow(row)

    def show_add_payslip_field_popup(self):
        popup = ctk.CTkToplevel(self) if hasattr(ctk, 'CTkToplevel') else tk.Toplevel(self)
        popup.title("Add Payslip Field")
        popup.geometry("420x320")
        popup.grab_set()
        ctk.CTkLabel(popup, text="Field Name (Display)", font=("Arial", 12)).pack(pady=(18, 4))
        field_name_var = tk.StringVar()
        ctk.CTkEntry(popup, textvariable=field_name_var, font=("Arial", 12), width=220).pack(pady=2)
        ctk.CTkLabel(popup, text="Excel Column Names (comma separated)", font=("Arial", 12)).pack(pady=(12, 4))
        excel_names_var = tk.StringVar()
        ctk.CTkEntry(popup, textvariable=excel_names_var, font=("Arial", 12), width=220).pack(pady=2)
        # Earnings/Deductions checkbox
        type_var = tk.StringVar(value="earnings")
        frame_type = ctk.CTkFrame(popup)
        frame_type.pack(pady=(12, 2))
        ctk.CTkLabel(frame_type, text="Display and calculate as:", font=("Arial", 12)).pack(side=tk.LEFT, padx=(0, 8))
        earnings_cb = ctk.CTkRadioButton(frame_type, text="Earnings", variable=type_var, value="earnings", font=("Arial", 11))
        deductions_cb = ctk.CTkRadioButton(frame_type, text="Deductions", variable=type_var, value="deductions", font=("Arial", 11))
        earnings_cb.pack(side=tk.LEFT)
        deductions_cb.pack(side=tk.LEFT, padx=(10, 0))
        error_label = ctk.CTkLabel(popup, text="", font=("Arial", 10), text_color="#d0021b")
        error_label.pack()
        def on_apply():
            field = field_name_var.get().strip()
            excel_names = [x.strip() for x in excel_names_var.get().split(',') if x.strip()]
            field_type = type_var.get()
            if not field or not excel_names:
                error_label.configure(text="Both fields are required.")
                return
            if not hasattr(self, '_custom_field_map'):
                self._custom_field_map = {}
            if not hasattr(self, '_custom_field_types'):
                self._custom_field_types = {}
            self._custom_field_map[field] = excel_names
            self._custom_field_types[field] = field_type
            self.save_updated_fields()  # Save first
            if hasattr(self, 'load_payslip_records'):
                self.load_payslip_records()  # Then reload records/UI
            popup.destroy()
        ctk.CTkButton(popup, text="Add", font=("Arial", 12, "bold"), width=120, command=on_apply).pack(pady=12)
        ctk.CTkButton(popup, text="Cancel", font=("Arial", 11), width=100, command=popup.destroy).pack()

    def show_remove_payslip_field_popup(self):
        popup = ctk.CTkToplevel(self) if hasattr(ctk, 'CTkToplevel') else tk.Toplevel(self)
        popup.title("Remove Payslip Field")
        popup.geometry("520x420")
        popup.grab_set()
        ctk.CTkLabel(popup, text="Select Field to Remove", font=("Arial", 12, "bold")).pack(pady=(18, 8))
        default_fields = [
            "BASIC SALARY", "PERA", "WITHHOLDING TAX", "GSIS EMPLOYEE SHARE", "PHILHEALTH EMPLOYEE SHARE", "PAGIBIG EMPLOYEE SHARE",
            "LANDBANK SALARY LOAN", "GSIS MPL", "GSIS MPL-LITE", "GSIS GFAL", "GSIS POLICY LOAN", "GSIS EMERGENCY LOAN", "GSIS CPL",
            "PAGIBIG CALAMITY LOAN", "PAGIBIG MPL", "NETPAY (1ST HALF)", "NETPAY (2ND HALF)"
        ]
        custom_fields = list(getattr(self, '_custom_field_map', {}).keys())
        all_fields = default_fields + custom_fields
        listbox = tk.Listbox(popup, font=("Arial", 12), height=10, selectmode=tk.SINGLE)
        for f in all_fields:
            listbox.insert(tk.END, f)
        listbox.pack(pady=8, padx=10, fill=tk.X)
        error_label = ctk.CTkLabel(popup, text="", font=("Arial", 10), text_color="#d0021b")
        error_label.pack()
        def on_remove():
            sel = listbox.curselection()
            if not sel:
                error_label.configure(text="Select a field to remove.")
                return
            idx = sel[0]
            field = all_fields[idx]
            # Remove from custom or default fields
            if field in custom_fields:
                # Remove from custom field map and types
                if field in self._custom_field_map:
                    del self._custom_field_map[field]
                if hasattr(self, '_custom_field_types') and field in self._custom_field_types:
                    del self._custom_field_types[field]
            elif field in default_fields:
                # Remove from default fields by tracking removed fields
                if not hasattr(self, '_removed_default_fields'):
                    self._removed_default_fields = set()
                self._removed_default_fields.add(field.upper()) # Store as upper to match check
            # After removal, immediately update CSV to reflect only current fields
            self.save_updated_fields()
            if hasattr(self, 'load_payslip_records'):
                self.load_payslip_records()
            popup.destroy()
        ctk.CTkButton(popup, text="Remove", font=("Arial", 12, "bold"), width=120, command=on_remove).pack(pady=12)
        ctk.CTkButton(popup, text="Cancel", font=("Arial", 11), width=100, command=popup.destroy).pack()