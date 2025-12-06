import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, InvalidOperation
import locale
import os

from logic import calculate_amortization_schedule

# --- Constants ---
DATE_FORMAT = "%Y-%m-%d"
SUPPORTED_DATE_FORMATS = [DATE_FORMAT, "%Y%m%d"]

METHOD_STRAIGHT_LINE = 'Straight-Line'
METHOD_DECLINING_BALANCE = 'Double Declining Balance'
METHOD_SOYD = "Sum-of-the-Years Digits"
AMORTIZATION_METHODS = [METHOD_STRAIGHT_LINE, METHOD_DECLINING_BALANCE, METHOD_SOYD]

COLUMNS = ('Period', 'Date', 'Description', 'Expense', 'Accumulated', 'Book Value')
NUMERIC_COLUMNS = ['Amortization Expense', 'Accumulated Amortization', 'Book Value']

# --- Utility Functions ---

def parse_flexible_date(date_string: str) -> datetime:
    """
    Parses a date string from one of the supported formats ('%Y-%m-%d' or '%Y%m%d').
    Raises ValueError if the string cannot be parsed with any of the formats.
    """
    # Clean whitespace from the input string
    date_string = date_string.strip()    
    formats_to_try = SUPPORTED_DATE_FORMATS
    for fmt in formats_to_try:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    raise ValueError("Invalid date format. Use 'YYYY-MM-DD' or 'YYYYMMDD'.")


class AmortizationApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("AmorPy - Amortization & Depreciation Calculator")
        
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "icon.ico")
            self.iconbitmap(icon_path)
        except tk.TclError:
            print("Warning: 'icon.ico' not found. The application will use a default icon.")
            
        self.geometry("700x550")
        
        self._setup_styles()
        
        self._create_menu()
        
        self.setup_locale()

        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        input_frame = ttk.Labelframe(main_frame, text="Input Data", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 5))

        self.asset_name_var = tk.StringVar(value="Your Asset Name")
        self.total_cost_var = tk.StringVar()
        self.method_var = tk.StringVar(value=METHOD_STRAIGHT_LINE)
        self.start_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-01'))
        self.end_date_var = tk.StringVar(value=(datetime.now() + relativedelta(years=1, days=-1)).strftime(DATE_FORMAT))
        self.schedule_df: pd.DataFrame | None = None

        self._create_input_widgets(input_frame)
        self._create_action_buttons(main_frame)

        output_frame = ttk.Labelframe(main_frame, text="Schedule", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self._create_output_widgets(output_frame)

        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5, style="Status.TLabel")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        style.configure("Status.TLabel", background="#f0f0f0", foreground="#333333")
        style.configure("Accent.TButton", font=('Helvetica', 9, 'bold'))

    def _create_menu(self) -> None:
        """Creates the main menu bar for the application."""
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Reset Form", command=self._clear_form)
        file_menu.add_command(label="Save Schedule", command=self._save_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about_dialog)

    def _show_about_dialog(self) -> None:
        """Displays the 'About' dialog box."""
        messagebox.showinfo(
            "About AmorPy",
            "AmorPy - Asset Amortization & Depreciation Calculator\nA simple tool to generate amortization schedules.\n\nVersion: 1.5\nCalculation methods:\n1. Double Declining Balance (DDB)\n2. Sum-of-the-Years Digits (SOYD)\n3. Straight-Line\n\nAuthor: Bintang\nContact: gaben@5222.de (XMPP/Jabber)\nGithub: https://github.com/bintanngg"
        )

    def _clear_form(self) -> None:
        """Resets all input fields and clears the output tree."""
        self.asset_name_var.set("Your Asset Name")
        self.total_cost_var.set("")
        self.method_var.set(METHOD_STRAIGHT_LINE)
        self.start_date_var.set(datetime.now().strftime('%Y-%m-01'))
        self.end_date_var.set((datetime.now() + relativedelta(years=1, days=-1)).strftime(DATE_FORMAT))
        
        self.schedule_df = None
        self.tree.delete(*self.tree.get_children())
        self.status_var.set("Form cleared. Ready for new calculation.")
        self.save_button.config(state='disabled')
        
    def _format_date_entry(self, date_var: tk.StringVar) -> None:
        try:
            date_str = date_var.get()
            if date_str:
                parsed_date = parse_flexible_date(date_str)
                date_var.set(parsed_date.strftime(DATE_FORMAT))
        except ValueError as e:
            messagebox.showerror("Incorrect Date Format", str(e))
            date_var.set("")

    def _format_total_cost(self, *args) -> None:
        self.total_cost_var.trace_remove('write', self.total_cost_trace_id)

        try:
            current_val = self.total_cost_var.get()
            cursor_pos = self.total_cost_entry.index(tk.INSERT)
            len_before = len(current_val)

            cleaned_val = ''.join(filter(str.isdigit, current_val))

            if cleaned_val:
                num_val = int(cleaned_val)
                formatted_val = locale.format_string("%d", num_val, grouping=True)
            else:
                formatted_val = ""

            self.total_cost_var.set(formatted_val)
            len_after = len(formatted_val)
            cursor_pos += (len_after - len_before)
            self.total_cost_entry.icursor(max(0, min(cursor_pos, len_after)))

        finally:
            self.total_cost_trace_id = self.total_cost_var.trace_add('write', self._format_total_cost)

    def setup_locale(self) -> None:
        try:
            locale.setlocale(locale.LC_ALL, 'id_ID')
            self.format_currency = lambda val: locale.format_string('%.2f', val, grouping=True)
        except locale.Error:
            self.format_currency = lambda val: f"{val:,.2f}".replace(',', '#').replace('.', ',').replace('#', '.')

    def _create_input_widgets(self, parent_frame: ttk.Frame) -> None:
        parent_frame.columnconfigure(1, weight=1)

        # Nama Aset
        ttk.Label(parent_frame, text="Asset Name").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(parent_frame, textvariable=self.asset_name_var, width=40).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

        # Harga Perolehan
        ttk.Label(parent_frame, text="Initial Value").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        total_cost_entry = ttk.Entry(parent_frame, textvariable=self.total_cost_var, width=40)
        total_cost_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        self.total_cost_entry = total_cost_entry
        self.total_cost_trace_id = self.total_cost_var.trace_add('write', self._format_total_cost)

        # Metode Amortisasi
        ttk.Label(parent_frame, text="Amortization Method").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        method_menu = ttk.Combobox(parent_frame, textvariable=self.method_var, values=AMORTIZATION_METHODS, state='readonly', width=38)
        method_menu.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)

        # Tanggal Mulai
        ttk.Label(parent_frame, text="Start Date (YYYY-MM-DD)").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        start_date_entry = ttk.Entry(parent_frame, textvariable=self.start_date_var, width=40)
        start_date_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)
        start_date_entry.bind('<FocusOut>', lambda event: self._format_date_entry(self.start_date_var))

        # Tanggal Selesai
        ttk.Label(parent_frame, text="End Date (YYYY-MM-DD)").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        end_date_entry = ttk.Entry(parent_frame, textvariable=self.end_date_var, width=40)
        end_date_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=5)
        end_date_entry.bind('<FocusOut>', lambda event: self._format_date_entry(self.end_date_var))

    def _create_action_buttons(self, parent_frame: ttk.Frame) -> None:
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill=tk.X, pady=(5, 10))

        ttk.Button(button_frame, text="Calculate", command=self._perform_calculation, style="Accent.TButton").pack(side=tk.LEFT)
        
        self.save_button = ttk.Button(button_frame, text="Save to xlsx", command=self._save_to_excel, state='disabled')
        self.save_button.pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame, text="Exit", command=self.destroy).pack(side=tk.RIGHT)

    def _create_output_widgets(self, parent_frame: ttk.Frame) -> None:
        self.tree = ttk.Treeview(parent_frame, show='headings')
        self.tree["columns"] = COLUMNS
        
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col, anchor=tk.W)
            anchor = tk.E if col not in ['Description', 'Date'] else tk.W
            self.tree.column(col, anchor=anchor, width=80)

        v_scroll = ttk.Scrollbar(parent_frame, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(parent_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        parent_frame.grid_rowconfigure(0, weight=1)
        parent_frame.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')

    def _validate_and_get_inputs(self) -> tuple | None:
        try:
            asset_name = self.asset_name_var.get()
            if not asset_name:
                raise ValueError("Asset Name cannot be empty.")
                
            total_cost_str = ''.join(filter(str.isdigit, self.total_cost_var.get()))
            if not total_cost_str:
                raise ValueError("Initial value cannot be empty.")

            total_cost = Decimal(total_cost_str)

            if total_cost <= 0:
                raise ValueError("Initial value must be greater than zero.")

            start_date = parse_flexible_date(self.start_date_var.get())
            end_date = parse_flexible_date(self.end_date_var.get())

            method = self.method_var.get()
            
            if end_date <= start_date:
                raise ValueError("End date must be after the start date.")
            
            return total_cost, start_date, end_date, method
            
        except InvalidOperation:
            messagebox.showerror("Input Error", "Initial value is invalid. Please enter a correct number.")
            return None
        except ValueError as e:
            messagebox.showerror("Error Input", str(e))
            return None

    def _populate_treeview(self) -> None:
        """Clears and populates the treeview with data from self.schedule_df."""
        self.tree.delete(*self.tree.get_children())
        if self.schedule_df is None:
            return

        for _, row in self.schedule_df.iterrows():
            formatted_values = [
                row['Period'],
                row['Date'].strftime(DATE_FORMAT),
                row['Description'],
                self.format_currency(row['Amortization Expense']),
                self.format_currency(row['Accumulated Amortization']),
                self.format_currency(row['Book Value'])
            ]
            self.tree.insert("", "end", values=formatted_values)

    def _perform_calculation(self) -> None:
        self.status_var.set("")
        self.save_button.config(state='disabled')
        self.schedule_df = None
        self.tree.delete(*self.tree.get_children())

        inputs = self._validate_and_get_inputs()
        if not inputs:
            return
        
        total_cost, start_date, end_date, method = inputs
        
        df, error_msg = calculate_amortization_schedule(total_cost, start_date, end_date, method)

        if error_msg:
            messagebox.showerror("Calculation Error", error_msg)
            return

        self.schedule_df = df
        self._populate_treeview()
        self.status_var.set("Calculation successful! You can now save to Excel.")
        self.save_button.config(state='normal')

    def _save_to_excel(self) -> None:
        if self.schedule_df is None:
            messagebox.showwarning("No Data", "There is no data to save. Please perform a calculation first.")
            return

        asset_name = self.asset_name_var.get() or "Unnamed Asset"
        method = self.method_var.get()

        output_filename = f'Result {method} Method - {asset_name}.xlsx'

        try:
            df_excel = self.schedule_df.copy()
            for col in NUMERIC_COLUMNS:
                df_excel[col] = df_excel[col].astype(float)

            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                df_excel.to_excel(writer, sheet_name='Schedule', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Schedule']

                number_format = '#,##0.00'

                for col_letter in ['D', 'E', 'F']:
                    for cell in worksheet[col_letter][1:]:
                        cell.number_format = number_format
            
            self.status_var.set(f"Success! Schedule saved to '{output_filename}'")
            messagebox.showinfo("Success", f"The amortization schedule has been saved to the file:\n{output_filename}")
        except Exception as e:
            self.status_var.set(f"Error: Failed to save Excel file.")
            messagebox.showerror("Save Error", f"Failed to save Excel file: {e}")
