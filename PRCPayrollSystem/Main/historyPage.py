import customtkinter as ctk
import tkinter as tk
import os
import csv
import sys
from PRCPayrollSystem.Main.resource_utils import resource_path

class HistoryPage(ctk.CTkFrame):
    def __init__(self, parent, controller=None):
        super().__init__(parent, fg_color="white")
        self.controller = controller
        self._loading_summary = False  # Flag to prevent re-entrant selection
        self.history_dir = resource_path("PRCPayrollSystem/pastLoadedHistory")
        # Directory for payslip PDFs
        self.payslip_dir = resource_path("PRCPayrollSystem/pastPayslips")
        self._build_ui()
        self.load_history_files()

    def _build_ui(self):
        topbar = ctk.CTkFrame(self, fg_color="white")
        topbar.pack(fill=ctk.X, pady=(5, 0))
        def go_to_main():
            if self.controller:
                self.controller.show_frame("MainMenu")
        ctk.CTkButton(topbar, text='>', font=('Consolas', 18, 'bold'), fg_color="white", text_color="#0041C2", hover_color="#e6e6e6", width=40, height=40, corner_radius=20, command=go_to_main).pack(side=ctk.LEFT, padx=(10, 20))
        ctk.CTkLabel(topbar, text='History', font=("Montserrat", 32, "normal"), text_color="#222").pack(side=ctk.LEFT, pady=10)
        # Open Full Table button (white text)
        self._open_full_table_btn = ctk.CTkButton(topbar, text="Open Full Table", text_color="white", fg_color="#0a47b1", hover_color="#003580", state="disabled", command=lambda: None)
        self._open_full_table_btn.pack(side=ctk.RIGHT, padx=(0, 8), pady=8)
        # Delete button (remodeled for checkbox multi-delete)
        self._delete_mode = False
        self._delete_btn = ctk.CTkButton(topbar, text="Delete", text_color="white", fg_color="#d0021b", hover_color="#a30010", state="normal", command=self.toggle_delete_mode)
        self._delete_btn.pack(side=ctk.RIGHT, padx=(0, 8), pady=8)

        # --- Modern, organized file list UI with its own scrollbar ---
        listbox_frame = ctk.CTkFrame(self, fg_color="#f3f6fa", border_width=1, border_color="#dbeafe", corner_radius=12)
        listbox_frame.pack(fill=ctk.BOTH, expand=True, padx=16, pady=(10, 0))
        ctk.CTkLabel(listbox_frame, text="History Files", font=("Montserrat", 16, "bold"), text_color="#0a47b1").grid(row=0, column=0, sticky="w", padx=12, pady=(8, 0))
        ctk.CTkLabel(listbox_frame, text="Payslip PDFs", font=("Montserrat", 16, "bold"), text_color="#0a47b1").grid(row=0, column=1, sticky="w", padx=12, pady=(8, 0))
        # CSV Listbox
        listbox_container = ctk.CTkFrame(listbox_frame, fg_color="#f3f6fa")
        listbox_container.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        listbox_scrollbar = ctk.CTkScrollbar(listbox_container, orientation="vertical")
        listbox_scrollbar.pack(side=ctk.RIGHT, fill=ctk.Y)
        # Set explicit width for the listbox container to match the delete mode canvas width
        listbox_container.update_idletasks()
        default_width = 600  # Adjust this value to match the delete mode canvas visually
        listbox_container.configure(width=default_width)
        self.listbox = tk.Listbox(listbox_container, font=("Consolas", 13), height=12, activestyle='none', bg="#f8fbff", fg="#0a47b1", highlightthickness=0, bd=0, relief="flat", selectbackground="#dbeafe", selectforeground="#0a47b1", borderwidth=0, yscrollcommand=listbox_scrollbar.set)
        self.listbox.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True)
        listbox_scrollbar.configure(command=self.listbox.yview)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        # Payslip PDF Listbox
        payslip_container = ctk.CTkFrame(listbox_frame, fg_color="#f3f6fa")
        payslip_container.grid(row=1, column=1, sticky="nsew", padx=8, pady=(0, 8))
        payslip_scrollbar = ctk.CTkScrollbar(payslip_container, orientation="vertical")
        payslip_scrollbar.pack(side=ctk.RIGHT, fill=ctk.Y)
        payslip_container.update_idletasks()
        payslip_default_width = 600  # Adjust as needed for symmetry
        payslip_container.configure(width=payslip_default_width)
        self.payslip_listbox = tk.Listbox(payslip_container, font=("Consolas", 13), height=12, activestyle='none', bg="#f8fbff", fg="#0a47b1", highlightthickness=0, bd=0, relief="flat", selectbackground="#dbeafe", selectforeground="#0a47b1", borderwidth=0, yscrollcommand=payslip_scrollbar.set)
        self.payslip_listbox.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True)
        payslip_scrollbar.configure(command=self.payslip_listbox.yview)
        self.payslip_listbox.bind('<<ListboxSelect>>', self.on_payslip_select)
        self.payslip_listbox.bind('<Double-Button-1>', self.on_payslip_double_click)
        # Add a subtle border and rounded corners to the listboxes using the frame
        listbox_frame.grid_columnconfigure(0, weight=1)
        listbox_frame.grid_columnconfigure(1, weight=1)
        self.summary_frame = ctk.CTkFrame(self, fg_color="white")
        self.summary_frame.pack(fill=ctk.BOTH, expand=True, padx=16, pady=8)

    def load_history_files(self):
        self.listbox.delete(0, tk.END)
        self.history_files = []
        if not os.path.exists(self.history_dir):
            os.makedirs(self.history_dir, exist_ok=True)
        files = [fname for fname in sorted(os.listdir(self.history_dir)) if fname.endswith('.csv')]
        # Show most recent first
        files = sorted(files, reverse=True)
        for i, fname in enumerate(files):
            self.history_files.append(fname)
            # Add index and date for clarity
            display_name = f"{i+1:02d}.  {fname}"
            self.listbox.insert(tk.END, display_name)
        # If no files, show a placeholder
        if not files:
            self.listbox.insert(tk.END, "No history files found.")
            self.listbox.configure(state="disabled")
        else:
            self.listbox.configure(state="normal")

        # Load payslip PDFs
        self.payslip_listbox.delete(0, tk.END)
        self.payslip_files = []
        if not os.path.exists(self.payslip_dir):
            os.makedirs(self.payslip_dir, exist_ok=True)
        pdfs = [fname for fname in sorted(os.listdir(self.payslip_dir)) if fname.lower().endswith('.pdf')]
        pdfs = sorted(pdfs, reverse=True)
        for i, fname in enumerate(pdfs):
            self.payslip_files.append(fname)
            display_name = f"{i+1:02d}.  {fname}"
            self.payslip_listbox.insert(tk.END, display_name)
        if not pdfs:
            self.payslip_listbox.insert(tk.END, "No payslip PDFs found.")
            self.payslip_listbox.configure(state="disabled")
        else:
            self.payslip_listbox.configure(state="normal")
        self._history_vars = []
        self._payslip_vars = []
        self._history_checkboxes = []
        self._payslip_checkboxes = []
        for i, fname in enumerate(self.history_files):
            self.listbox.delete(i)
            display_name = f"{i+1:02d}.  {fname}"
            self.listbox.insert(tk.END, display_name)
        for i, fname in enumerate(self.payslip_files):
            self.payslip_listbox.delete(i)
            display_name = f"{i+1:02d}.  {fname}"
            self.payslip_listbox.insert(tk.END, display_name)

    def on_select(self, event):
        # Disable old single-file delete logic; only enable Open Full Table
        selection = event.widget.curselection()
        if not selection or not self.history_files:
            self._open_full_table_btn.configure(state="disabled", command=lambda: None)
            return
        idx = selection[0]
        if idx >= len(self.history_files):
            self._open_full_table_btn.configure(state="disabled", command=lambda: None)
            return
        fname = self.history_files[idx]
        self.show_history_summary(fname)
        fpath = os.path.join(self.history_dir, fname)
        self._open_full_table_btn.configure(state="normal", command=lambda: self.open_full_table(fpath, fname))
        # Do not enable delete button for single selection

    def on_payslip_select(self, event):
        # Disable old single-file delete logic for payslips
        return

    def confirm_delete_file(self, fpath, fname, is_pdf=False):
        import tkinter.messagebox as messagebox
        answer = messagebox.askquestion("Delete File", f"Are you sure you want to delete '{fname}'?", icon='warning')
        if answer == 'yes':
            try:
                os.remove(fpath)
                self.refresh()
                messagebox.showinfo("Deleted", f"'{fname}' has been deleted.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete file: {e}")
        # else: do nothing (cancel)

    def show_history_summary(self, fname):
        fpath = os.path.join(self.history_dir, fname)
        self._open_full_table_btn.configure(state="normal", command=lambda: self.open_full_table(fpath, fname))
        self.after(150, lambda: self._do_show_history_summary(fname))

    def _do_show_history_summary(self, fname):
        for widget in self.summary_frame.winfo_children():
            widget.destroy()
        fpath = os.path.join(self.history_dir, fname)
        try:
            import csv
            with open(fpath, newline='', encoding='utf-8') as f:
                reader = list(csv.reader(f))
                if not reader:
                    ctk.CTkLabel(self.summary_frame, text="No data in file").pack()
                    return
                headers = reader[0]
                data_rows = reader[1:]

                #scrollable canvas
                container = tk.Frame(self.summary_frame, bg="white")
                container.pack(fill=ctk.BOTH, expand=True, pady=(0, 0))
                container.grid_rowconfigure(0, weight=1)
                container.grid_columnconfigure(0, weight=1)
                canvas = tk.Canvas(container, bg="white", highlightthickness=0, height=600)
                canvas.grid(row=0, column=0, sticky="nsew")
                v_scroll = ctk.CTkScrollbar(container, orientation="vertical", command=canvas.yview, width=18)
                v_scroll.grid(row=0, column=1, sticky="ns")
                h_scroll = ctk.CTkScrollbar(container, orientation="horizontal", command=canvas.xview, height=18)
                h_scroll.grid(row=1, column=0, sticky="ew")
                canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

                table_frame = tk.Frame(canvas, bg="white")
                table_window = canvas.create_window((0, 0), window=table_frame, anchor="nw")

                # --- Table content ---
                header_font = ("Arial", 10, "bold")
                cell_font = ("Arial", 10)
                # Title
                title = tk.Label(table_frame, text=f"Summary of {fname}", font=("Arial", 14, "bold"), bg="white", anchor="w")
                title.grid(row=0, column=0, columnspan=len(headers), sticky="w", pady=(0, 4))
                # Headers
                for c, h in enumerate(headers):
                    lbl = tk.Label(table_frame, text=h, font=header_font, bg="#39d353", fg="white", padx=8, pady=4, borderwidth=1, relief="solid")
                    lbl.grid(row=1, column=c, sticky="nsew", padx=1, pady=1)
                    table_frame.grid_columnconfigure(c, weight=1, minsize=100)
                # Data rows
                for r, row in enumerate(data_rows):
                    for c, val in enumerate(row):
                        lbl = tk.Label(table_frame, text=val, font=cell_font, bg="white", fg="#222", padx=8, pady=4, borderwidth=1, relief="solid")
                        lbl.grid(row=2 + r, column=c, sticky="nsew", padx=1, pady=1)
                # Fill empty cells if row is short
                for r, row in enumerate(data_rows):
                    for c in range(len(row), len(headers)):
                        lbl = tk.Label(table_frame, text="", font=cell_font, bg="white", padx=8, pady=4, borderwidth=1, relief="solid")
                        lbl.grid(row=2 + r, column=c, sticky="nsew", padx=1, pady=1)

                # --- Update scrollregion on table_frame resize ---
                def _update_scrollregion(event=None):
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    # Only set the window width if the table is smaller than the canvas
                    bbox = canvas.bbox(table_window)
                    if bbox:
                        table_width = bbox[2] - bbox[0]
                        canvas_width = canvas.winfo_width()
                        if table_width < canvas_width:
                            canvas.itemconfig(table_window, width=canvas_width)
                        else:
                            canvas.itemconfig(table_window, width=table_width)
                table_frame.bind("<Configure>", _update_scrollregion)
                canvas.bind("<Configure>", _update_scrollregion)

                # Mousewheel scrolling (vertical and horizontal)
                def _on_mousewheel(event):
                    if event.state & 0x1:  # Shift is held for horizontal scroll
                        canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
                    else:
                        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                canvas.bind("<MouseWheel>", _on_mousewheel)

        except Exception as e:
            ctk.CTkLabel(self.summary_frame, text=f"Error loading file: {e}").pack()

    def on_payslip_double_click(self, event):
        # Still allow opening the PDF externally
        selection = event.widget.curselection()
        if not selection or not self.payslip_files:
            return
        idx = selection[0]
        if idx >= len(self.payslip_files):
            return
        fname = self.payslip_files[idx]
        pdf_path = os.path.join(self.payslip_dir, fname)
        import subprocess
        import sys
        if sys.platform.startswith('win'):
            os.startfile(pdf_path)
        elif sys.platform.startswith('darwin'):
            subprocess.call(['open', pdf_path])
        else:
            subprocess.call(['xdg-open', pdf_path])

    def open_full_table(self, fpath, fname):
        # Read the CSV file
        with open(fpath, newline='', encoding='utf-8') as f:
            reader = list(csv.reader(f))
            if not reader:
                return
            headers = reader[0]
            data = reader[1:]
        # Pass data to ExcelImportPage and navigate
        if self.controller:
            excel_import_page = self.controller.get_page('ExcelImportPage')
            if excel_import_page:
                excel_import_page.set_aggregated_data(headers, data)
            self.controller.show_frame('ExcelImportPage')

    def refresh(self):
        # Reload the file list and also clear and reload the payslip list
        self.load_history_files()
        # Optionally, clear the summary frame to avoid showing stale previews or summaries
        for widget in self.summary_frame.winfo_children():
            widget.destroy()

    def toggle_delete_mode(self):
        if not self._delete_mode:
            self._delete_mode = True
            self._delete_btn.configure(text="Confirm Delete", fg_color="#a30010")
            self.show_delete_checkboxes()
        else:
            self.perform_delete_checked()
            self._delete_mode = False
            self._delete_btn.configure(text="Delete", fg_color="#d0021b")
            self.hide_delete_checkboxes()

    def show_delete_checkboxes(self):
        # Hide original listboxes
        self.listbox.pack_forget()
        self.payslip_listbox.pack_forget()
        # Get the current size of the listbox widgets themselves (not the container)
        self.listbox.update_idletasks()
        self.payslip_listbox.update_idletasks()
        listbox_width = self.listbox.winfo_width()
        listbox_height = self.listbox.winfo_height()
        payslip_width = self.payslip_listbox.winfo_width()
        payslip_height = self.payslip_listbox.winfo_height()
        # --- History Files with checkboxes ---
        self._history_canvas = tk.Canvas(self.listbox.master, bg="#f8fbff", highlightthickness=0,
                                         width=listbox_width, height=listbox_height)
        self._history_canvas.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True)
        self.listbox.master.children['!ctkscrollbar'].configure(command=self._history_canvas.yview)
        self._history_canvas.configure(yscrollcommand=self.listbox.master.children['!ctkscrollbar'].set)
        self._history_frame = tk.Frame(self._history_canvas, bg="#f8fbff")
        self._history_canvas.create_window((0, 0), window=self._history_frame, anchor="nw")
        self._history_vars = []
        for i, fname in enumerate(self.history_files):
            var = tk.BooleanVar()
            row = tk.Frame(self._history_frame, bg="#f8fbff")
            # Set a fixed width for the label so the row width matches the listbox, accounting for the checkbox
            label = tk.Label(row, text=f"{i+1:02d}.  {fname}", font=("Consolas", 13), bg="#f8fbff", fg="#0a47b1", anchor="w", width=max(1, int(listbox_width/9)-4))
            label.pack(side=tk.LEFT, fill=tk.X, expand=True)
            cb = tk.Checkbutton(row, variable=var, bg="#f8fbff")
            cb.pack(side=tk.RIGHT, padx=2)
            row.pack(fill=tk.X, padx=2, pady=1)
            self._history_vars.append((var, os.path.join(self.history_dir, fname)))
        self._history_frame.update_idletasks()
        self._history_canvas.config(scrollregion=self._history_canvas.bbox("all"))
        def _on_history_frame_configure(event):
            self._history_canvas.config(scrollregion=self._history_canvas.bbox("all"))
        self._history_frame.bind("<Configure>", _on_history_frame_configure)
        # --- Payslip PDFs with checkboxes ---
        self._payslip_canvas = tk.Canvas(self.payslip_listbox.master, bg="#f8fbff", highlightthickness=0,
                                         width=payslip_width, height=payslip_height)
        self._payslip_canvas.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True)
        self.payslip_listbox.master.children['!ctkscrollbar'].configure(command=self._payslip_canvas.yview)
        self._payslip_canvas.configure(yscrollcommand=self.payslip_listbox.master.children['!ctkscrollbar'].set)
        self._payslip_frame = tk.Frame(self._payslip_canvas, bg="#f8fbff")
        self._payslip_canvas.create_window((0, 0), window=self._payslip_frame, anchor="nw")
        self._payslip_vars = []
        for i, fname in enumerate(self.payslip_files):
            var = tk.BooleanVar()
            row = tk.Frame(self._payslip_frame, bg="#f8fbff")
            label = tk.Label(row, text=f"{i+1:02d}.  {fname}", font=("Consolas", 13), bg="#f8fbff", fg="#0a47b1", anchor="w", width=max(1, int(payslip_width/9)-4))
            label.pack(side=tk.LEFT, fill=tk.X, expand=True)
            cb = tk.Checkbutton(row, variable=var, bg="#f8fbff")
            cb.pack(side=tk.RIGHT, padx=2)
            row.pack(fill=tk.X, padx=2, pady=1)
            self._payslip_vars.append((var, os.path.join(self.payslip_dir, fname)))
        self._payslip_frame.update_idletasks()
        self._payslip_canvas.config(scrollregion=self._payslip_canvas.bbox("all"))
        def _on_payslip_frame_configure(event):
            self._payslip_canvas.config(scrollregion=self._payslip_canvas.bbox("all"))
        self._payslip_frame.bind("<Configure>", _on_payslip_frame_configure)
        # Prevent the canvas from resizing to its contents
        self._history_canvas.pack_propagate(False)
        self._payslip_canvas.pack_propagate(False)

    def hide_delete_checkboxes(self):
        # Remove the custom frames and restore the original listboxes
        if hasattr(self, '_history_canvas'):
            self._history_canvas.pack_forget()
            self._history_canvas.destroy()
        if hasattr(self, '_payslip_canvas'):
            self._payslip_canvas.pack_forget()
            self._payslip_canvas.destroy()
        # Restore the original scrollbar commands to the listboxes
        self.listbox.master.children['!ctkscrollbar'].configure(command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=self.listbox.master.children['!ctkscrollbar'].set)
        self.payslip_listbox.master.children['!ctkscrollbar'].configure(command=self.payslip_listbox.yview)
        self.payslip_listbox.configure(yscrollcommand=self.payslip_listbox.master.children['!ctkscrollbar'].set)
        self.listbox.pack(side=ctk.LEFT, fill=ctk.X, expand=True)
        self.payslip_listbox.pack(side=ctk.LEFT, fill=ctk.X, expand=True)
        self._history_vars = []
        self._payslip_vars = []
        self.load_history_files()

    def perform_delete_checked(self):
        import tkinter.messagebox as messagebox
        deleted = False
        # Delete checked history files
        for var, fpath in getattr(self, '_history_vars', []):
            if var.get() and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                    deleted = True
                except Exception:
                    pass
        # Delete checked payslip files
        for var, fpath in getattr(self, '_payslip_vars', []):
            if var.get() and os.path.exists(fpath):
                try:
                    os.remove(fpath)
                    deleted = True
                except Exception:
                    pass
        if deleted:
            self.refresh()
            messagebox.showinfo("Deleted", "Selected files have been deleted.")
        else:
            messagebox.showinfo("No Selection", "No files selected for deletion.")
