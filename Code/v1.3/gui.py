import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttk
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, InvalidOperation
import locale
import os

# Import the core calculation function from the logic file
from logic import calculate_amortization_schedule

def parse_flexible_date(date_string: str) -> datetime:
    """
    Parses a date string from one of the supported formats ('%Y-%m-%d' or '%Y%m%d').
    Raises ValueError if the string cannot be parsed with any of the formats.
    """
    # Clean whitespace from the input string
    date_string = date_string.strip()
    
    formats_to_try = ["%Y-%m-%d", "%Y%m%d"]
    for fmt in formats_to_try:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    # If all formats failed, raise an error with clear instructions
    raise ValueError("Format tanggal tidak valid. Gunakan 'YYYY-MM-DD' atau 'YYYYMMDD'.")


class AmortizationApp(ttk.Window):
    def __init__(self, themename="litera"):
        super().__init__(themename=themename)
        self.title("AmorPy v1.3 | Bintang")
        
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "icon.ico")
            self.iconbitmap(icon_path)
        except tk.TclError:
            print("Warning: 'icon.ico' not found. The application will use a default icon.")
            
        self.geometry("750x650")
        
        self.setup_locale()

        # Main frame setup
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input Frame
        input_frame = ttk.Labelframe(main_frame, text="Input Data", padding="10")
        input_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

        # Initialize variables before creating widgets
        self.asset_name_var = tk.StringVar(value="Your Asset Name")
        self.total_cost_var = tk.StringVar()
        self.method_var = tk.StringVar(value='Straight-Line')
        self.start_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-01'))
        self.end_date_var = tk.StringVar(value=(datetime.now() + relativedelta(years=1, days=-1)).strftime('%Y-%m-%d'))
        self.schedule_df = None # To store the dataframe

        self._create_input_widgets(input_frame)
        self._create_action_buttons(main_frame)

        # Output Frame
        output_frame = ttk.Labelframe(main_frame, text="Jadwal Amortisasi", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._create_output_widgets(output_frame)

        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5, bootstyle="inverse-light")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _format_date_entry(self, date_var: tk.StringVar):
        """Formats the date in a StringVar to YYYY-MM-DD on focus out."""
        try:
            date_str = date_var.get()
            if date_str:  # only format if not empty
                parsed_date = parse_flexible_date(date_str)
                date_var.set(parsed_date.strftime('%Y-%m-%d'))
        except ValueError as e:
            # If parsing fails, show an error immediately for better UX.
            # Clear the invalid entry to force user correction.
            messagebox.showerror("Format Tanggal Salah", str(e))
            date_var.set("")

    def _format_total_cost(self, *args):
        """Formats the total cost entry with thousand separators as the user types."""
        # Temporarily remove the trace to avoid recursion
        self.total_cost_var.trace_remove('write', self.total_cost_trace_id)

        try:
            current_val = self.total_cost_var.get()
            cursor_pos = self.total_cost_entry.index(tk.INSERT)
            len_before = len(current_val)

            # Remove non-digit chars
            cleaned_val = ''.join(filter(str.isdigit, current_val))

            if cleaned_val: # If there are any digits
                num_val = int(cleaned_val)
                # Format with locale-specific grouping (e.g., '1.000.000')
                formatted_val = locale.format_string("%d", num_val, grouping=True)
            else: # If no digits are left, the field should be empty
                formatted_val = ""
            
            self.total_cost_var.set(formatted_val)
            
            # Adjust cursor position based on the change in length
            len_after = len(formatted_val)
            cursor_pos += (len_after - len_before)
            self.total_cost_entry.icursor(max(0, min(cursor_pos, len_after)))

        finally:
            # Always re-add the trace
            self.total_cost_trace_id = self.total_cost_var.trace_add('write', self._format_total_cost)



    def setup_locale(self):
        """Sets up locale for currency formatting with a fallback."""
        try:
            locale.setlocale(locale.LC_ALL, 'id_ID')
            self.format_currency = lambda val: locale.format_string('%.2f', val, grouping=True)
        except locale.Error:
            # Fallback if 'id_ID' locale is not installed on the system
            self.format_currency = lambda val: f"{val:,.2f}".replace(',', '#').replace('.', ',').replace('#', '.')

    def _create_input_widgets(self, parent_frame):
        """Creates and places all the input widgets in the parent frame."""
        # Grid configuration
        parent_frame.columnconfigure(1, weight=1)

        # Nama Aset
        ttk.Label(parent_frame, text="Asset Name").grid(row=0, column=0, sticky=tk.W, padx=5, pady=8)
        ttk.Entry(parent_frame, textvariable=self.asset_name_var, width=40).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=8)

        # Harga Perolehan
        ttk.Label(parent_frame, text="Initial Value").grid(row=1, column=0, sticky=tk.W, padx=5, pady=8)
        total_cost_entry = ttk.Entry(parent_frame, textvariable=self.total_cost_var, width=40)
        total_cost_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=8)
        self.total_cost_entry = total_cost_entry
        self.total_cost_trace_id = self.total_cost_var.trace_add('write', self._format_total_cost)

        # Metode Amortisasi
        ttk.Label(parent_frame, text="Amortization Method").grid(row=2, column=0, sticky=tk.W, padx=5, pady=8)
        method_options = ['Straight-Line', 'Double Declining Balance']
        method_menu = ttk.Combobox(parent_frame, textvariable=self.method_var, values=method_options, state='readonly', width=38)
        method_menu.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=8)

        # Tanggal Mulai
        ttk.Label(parent_frame, text="Start Date (YYYY-MM-DD)").grid(row=3, column=0, sticky=tk.W, padx=5, pady=8)
        start_date_entry = ttk.Entry(parent_frame, textvariable=self.start_date_var, width=40)
        start_date_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=8)
        start_date_entry.bind('<FocusOut>', lambda event: self._format_date_entry(self.start_date_var))

        # Tanggal Selesai
        ttk.Label(parent_frame, text="End Date (YYYY-MM-DD)").grid(row=4, column=0, sticky=tk.W, padx=5, pady=8)
        end_date_entry = ttk.Entry(parent_frame, textvariable=self.end_date_var, width=40)
        end_date_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=8)
        end_date_entry.bind('<FocusOut>', lambda event: self._format_date_entry(self.end_date_var))

    def _create_action_buttons(self, parent_frame):
        """Creates and places the main action buttons."""
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill=tk.X, padx=15, pady=(5, 15))

        ttk.Button(button_frame, text="Calculate", command=self._perform_calculation, bootstyle="primary").pack(side=tk.LEFT)
        
        self.save_button = ttk.Button(button_frame, text="Save to xlsx", command=self._save_to_excel, state='disabled', bootstyle="success")
        self.save_button.pack(side=tk.LEFT, padx=10)

        ttk.Button(button_frame, text="Exit", command=self.quit, bootstyle="danger-outline").pack(side=tk.RIGHT)

    def _create_output_widgets(self, parent_frame):
        """Creates the treeview for displaying the schedule."""
        self.tree = ttk.Treeview(parent_frame, show='headings')
        self.tree["columns"] = ('Periode', 'Tanggal', 'Deskripsi', 'Beban Amortisasi', 'Akumulasi Amortisasi', 'Nilai Buku')
        
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            anchor = tk.E if col not in ['Deskripsi', 'Tanggal'] else tk.W
            self.tree.column(col, anchor=anchor, width=80)

        v_scroll = ttk.Scrollbar(parent_frame, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(parent_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        parent_frame.grid_rowconfigure(0, weight=1)
        parent_frame.grid_columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')

    def _perform_calculation(self):
        self.tree.delete(*self.tree.get_children())
        self.status_var.set("")
        self.save_button.config(state='disabled')
        self.schedule_df = None

        try:
            asset_name = self.asset_name_var.get()
            if not asset_name:
                raise ValueError("Nama Aset tidak boleh kosong.")
                
            total_cost_str = self.total_cost_var.get().replace('.', '').replace(',', '.')
            if not total_cost_str:
                raise ValueError("Harga perolehan tidak boleh kosong.")

            total_cost = Decimal(total_cost_str)

            if total_cost <= 0:
                raise ValueError("Harga perolehan harus lebih besar dari nol.")

            # Use the flexible date parsing function
            start_date = parse_flexible_date(self.start_date_var.get())
            end_date = parse_flexible_date(self.end_date_var.get())

            method = self.method_var.get()
            
            if end_date <= start_date:
                raise ValueError("Tanggal selesai harus setelah tanggal mulai.")

        except InvalidOperation:
            messagebox.showerror("Error Input", "Harga perolehan tidak valid. Harap masukkan angka yang benar (misal: 10.000,50).")
            return
        except ValueError as e:
            # Display the specific error message from parsing or validation
            messagebox.showerror("Error Input", str(e))
            return
        
        self.schedule_df, error_msg = calculate_amortization_schedule(total_cost, start_date, end_date, method)

        if error_msg:
            messagebox.showerror("Error Perhitungan", error_msg)
            return

        # Populate the treeview with formatted data
        for index, row in self.schedule_df.iterrows():
            formatted_values = [
                row['Periode'],
                row['Tanggal'].strftime('%Y-%m-%d'),
                row['Deskripsi'],
                self.format_currency(row['Beban Amortisasi']),
                self.format_currency(row['Akumulasi Amortisasi']),
                self.format_currency(row['Nilai Buku'])
            ]
            self.tree.insert("", "end", values=formatted_values)
        
        self.status_var.set("Perhitungan berhasil! Anda sekarang dapat menyimpan ke Excel.")
        self.save_button.config(state='normal')

    def _save_to_excel(self):
        if self.schedule_df is None:
            messagebox.showwarning("Tidak Ada Data", "Tidak ada data untuk disimpan. Harap lakukan perhitungan terlebih dahulu.")
            return

        asset_name = self.asset_name_var.get() or "Tanpa Nama"
        method = self.method_var.get()

        output_filename = f'Hasil Amortisasi {method} - {asset_name}.xlsx'

        try:
            # Create a new DataFrame for excel export
            df_excel = self.schedule_df.copy()
            
            # Convert Decimal to float for Excel compatibility
            numeric_cols = ['Beban Amortisasi', 'Akumulasi Amortisasi', 'Nilai Buku']
            for col in numeric_cols:
                df_excel[col] = df_excel[col].astype(float)

            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                df_excel.to_excel(writer, sheet_name='Jadwal Amortisasi', index=False)
                
                # Get the workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Jadwal Amortisasi']

                # Define the number format
                # Format ini akan menampilkan pemisah ribuan (titik) dan dua desimal (koma)
                # seperti yang umum di Indonesia.
                number_format = '#,##0.00'

                # Apply the format to the numeric columns (e.g., D, E, F)
                for col_letter in ['D', 'E', 'F']:
                    for cell in worksheet[col_letter][1:]: # Skip header
                        cell.number_format = number_format
            
            self.status_var.set(f"Sukses! Jadwal disimpan ke '{output_filename}'")
            messagebox.showinfo("Sukses", f"Jadwal amortisasi telah disimpan ke file:\n{output_filename}")
        except Exception as e:
            self.status_var.set(f"Error: Gagal menyimpan file Excel.")
            messagebox.showerror("Error Penyimpanan", f"Gagal menyimpan file Excel: {e}")
