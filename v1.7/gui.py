import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, InvalidOperation
import locale
import os

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from logic import calculate_amortization_schedule
from tkcalendar import DateEntry

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
    """Parses a date string from one of the supported formats."""
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
        self.title("AmorPy v1.7 - Amortization & Depreciation Calculator")
        
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "icon.ico")
            self.iconbitmap(icon_path)
        except tk.TclError:
            pass # Default icon
            
        self.geometry("800x650")
        
        self._setup_styles()
        self._create_menu()
        self.setup_locale()

        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Input Section ---
        input_frame = ttk.Labelframe(main_frame, text="Input Data", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        self.asset_name_var = tk.StringVar(value="New Asset")
        self.total_cost_var = tk.StringVar()
        self.salvage_value_var = tk.StringVar(value="0")
        self.method_var = tk.StringVar(value=METHOD_STRAIGHT_LINE)
        self.start_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-01'))
        self.end_date_var = tk.StringVar(value=(datetime.now() + relativedelta(years=1, days=-1)).strftime(DATE_FORMAT))
        self.schedule_df: pd.DataFrame | None = None

        self._create_input_widgets(input_frame)
        self._create_action_buttons(main_frame)

        # --- Output Section (Tabs) ---
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Tab 1: Schedule (Table)
        self.tab_schedule = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(self.tab_schedule, text="Schedule")
        self._create_schedule_widgets(self.tab_schedule)
        
        # Tab 2: Chart (Graph)
        self.tab_chart = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(self.tab_chart, text="Chart")
        self.chart_canvas = None

        # --- Status Bar ---
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5, style="Status.TLabel")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        try:
            self.tk.call("source", "azure.tcl") # Optional: Try to load a theme if available, otherwise default
            self.tk.call("set_theme", "light") 
        except:
            pass
            
        style.configure("Status.TLabel", background="#f0f0f0", foreground="#333333")
        style.configure("Accent.TButton", font=('Helvetica', 9, 'bold'))

    def _create_menu(self) -> None:
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
        messagebox.showinfo(
            "About AmorPy",
            "AmorPy v1.7\n\nAdd Calendar Picker.\n\nAuthor: Bintang\nGithub: https://github.com/bintanngg"
        )

    def _clear_form(self) -> None:
        self.asset_name_var.set("New Asset")
        self.total_cost_var.set("")
        self.salvage_value_var.set("0")
        # self.start_date_var and self.end_date_var are now handled by DateEntry widgets directly
        # but we can still reset them if we keep references to widgets or just recreate them.
        # Simpler approach: if we keep the vars, DateEntry updates them? 
        # DateEntry has textvariable support.
        self.start_date_var.set(datetime.now().strftime('%Y-%m-01'))
        self.end_date_var.set((datetime.now() + relativedelta(years=1, days=-1)).strftime(DATE_FORMAT))
        
        # We need to explicitly set the date on the widgets if they exist
        if hasattr(self, 'start_date_entry'):
            self.start_date_entry.set_date(datetime.now().replace(day=1))
        if hasattr(self, 'end_date_entry'):
            self.end_date_entry.set_date(datetime.now() + relativedelta(years=1, days=-1))
        
        self.schedule_df = None
        self.tree.delete(*self.tree.get_children())
        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()
            self.chart_canvas = None
            
        self.status_var.set("Form cleared.")
        self.save_button.config(state='disabled')

    def setup_locale(self) -> None:
        try:
            locale.setlocale(locale.LC_ALL, 'id_ID')
            self.format_currency = lambda val: locale.format_string('%.2f', val, grouping=True)
        except locale.Error:
            self.format_currency = lambda val: f"{val:,.2f}"

    def _format_currency_input(self, var: tk.StringVar):
        """Standardizes input to allow for calculations while keeping user friendly formatting."""
        # This is a simplified version. For a robust app, use a proper diverse validation.
        # Here we just want to NOT strip decimals.
        pass 

    def _create_input_widgets(self, parent_frame: ttk.Frame) -> None:
        parent_frame.columnconfigure(1, weight=1)
        parent_frame.columnconfigure(3, weight=1)

        # Row 0: Asset Name & Method
        ttk.Label(parent_frame, text="Asset Name").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(parent_frame, textvariable=self.asset_name_var).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

        ttk.Label(parent_frame, text="Method").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(parent_frame, textvariable=self.method_var, values=AMORTIZATION_METHODS, state='readonly').grid(row=0, column=3, sticky=tk.EW, padx=5, pady=5)

        # Row 1: Cost & Salvage
        ttk.Label(parent_frame, text="Total Cost").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(parent_frame, textvariable=self.total_cost_var).grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)

        ttk.Label(parent_frame, text="Salvage Value").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(parent_frame, textvariable=self.salvage_value_var).grid(row=1, column=3, sticky=tk.EW, padx=5, pady=5)

        # Row 2: Dates
        ttk.Label(parent_frame, text="Start Date").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.start_date_entry = DateEntry(parent_frame, textvariable=self.start_date_var, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='y-mm-dd', state='readonly')
        self.start_date_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        self.start_date_entry.bind('<Button-1>', lambda e: self.start_date_entry.drop_down())

        ttk.Label(parent_frame, text="End Date").grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        self.end_date_entry = DateEntry(parent_frame, textvariable=self.end_date_var, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='y-mm-dd', state='readonly')
        self.end_date_entry.grid(row=2, column=3, sticky=tk.EW, padx=5, pady=5)
        self.end_date_entry.bind('<Button-1>', lambda e: self.end_date_entry.drop_down())

    def _create_action_buttons(self, parent_frame: ttk.Frame) -> None:
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(button_frame, text="Calculate & Plot", command=self._perform_calculation, style="Accent.TButton").pack(side=tk.LEFT)
        self.save_button = ttk.Button(button_frame, text="Save to Excel", command=self._save_to_excel, state='disabled')
        self.save_button.pack(side=tk.LEFT, padx=10)

    def _create_schedule_widgets(self, parent_frame: ttk.Frame) -> None:
        self.tree = ttk.Treeview(parent_frame, show='headings')
        self.tree["columns"] = COLUMNS
        
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col, anchor=tk.W)
            anchor = tk.E if col not in ['Description', 'Date'] else tk.W
            self.tree.column(col, anchor=anchor, width=90)

        v_scroll = ttk.Scrollbar(parent_frame, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(parent_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        parent_frame.grid_rowconfigure(0, weight=1)
        parent_frame.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')

    def _parse_currency(self, val: str) -> Decimal:
        # Remove currency symbols and grouping separators, keep decimal point
        clean = ''.join(c for c in val if c.isdigit() or c == '.' or c == ',')
        # If both . and , are present, assume last one is decimal
        if '.' in clean and ',' in clean:
             if clean.rfind('.') > clean.rfind(','):
                 clean = clean.replace(',', '') # remove thousands sep
             else:
                 clean = clean.replace('.', '').replace(',', '.') # swap
        elif ',' in clean and '.' not in clean: 
            # Check if likely a decimal or thousands (simple heuristic: if 2 decimals, likely cents)
             parts = clean.split(',')
             if len(parts) == 2 and len(parts[1]) == 2:
                 clean = clean.replace(',', '.')
        
        if not clean:
             return Decimal("0")
        return Decimal(clean)

    def _validate_and_get_inputs(self) -> tuple | None:
        try:
            asset_name = self.asset_name_var.get()
            
            # IMPROVED PARSING
            total_cost = self._parse_currency(self.total_cost_var.get())
            salvage_value = self._parse_currency(self.salvage_value_var.get())

            if total_cost <= 0:
                raise ValueError("Total Cost must be > 0.")
            if salvage_value < 0:
                raise ValueError("Salvage Value cannot be negative.")

            start_date = parse_flexible_date(self.start_date_var.get())
            end_date = parse_flexible_date(self.end_date_var.get())
            method = self.method_var.get()
            
            if end_date <= start_date:
                raise ValueError("End Date must be after Start Date.")
            
            return total_cost, salvage_value, start_date, end_date, method
            
        except InvalidOperation:
            messagebox.showerror("Input Error", "Invalid number format.")
            return None
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
            return None

    def _perform_calculation(self) -> None:
        self.schedule_df = None
        self.tree.delete(*self.tree.get_children())
        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()
            self.chart_canvas = None

        inputs = self._validate_and_get_inputs()
        if not inputs:
            return
        
        total_cost, salvage_value, start_date, end_date, method = inputs
        
        df, error_msg = calculate_amortization_schedule(total_cost, salvage_value, start_date, end_date, method)

        if error_msg:
            messagebox.showerror("Calculation Error", error_msg)
            return

        self.schedule_df = df
        
        # Populate Table
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

        # Plot Chart
        self._plot_chart(df)

        self.status_var.set("Calculation successful.")
        self.save_button.config(state='normal')
        
    def _plot_chart(self, df: pd.DataFrame) -> None:
        fig = Figure(figsize=(5, 4), dpi=100)
        ax = fig.add_subplot(111)
        
        dates = df['Date']
        book_value = df['Book Value']
        accumulated = df['Accumulated Amortization']
        
        ax.plot(dates, book_value, label='Book Value', color='#2196F3', linewidth=2)
        ax.plot(dates, accumulated, label='Accumulated Amortization', color='#F44336', linewidth=2)
        
        ax.set_title('Amortization Schedule')
        ax.set_xlabel('Date')
        ax.set_ylabel('Currency')
        ax.legend()
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # Embed in Tkinter
        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.tab_chart)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def _save_to_excel(self) -> None:
        if self.schedule_df is None:
            return
            
        try:
            default_filename = f"Schedule_{self.asset_name_var.get()}.xlsx".replace(" ", "_")
            self.schedule_df.to_excel(default_filename, index=False)
            messagebox.showinfo("Success", f"Saved to {default_filename}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

