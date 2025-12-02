import tkinter as tk
from tkinter import ttk, messagebox
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


class AmortizationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AmorPy v1.2 | Bintang")
        
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "icon.ico")
            self.iconbitmap(icon_path)
        except tk.TclError:
            print("Warning: 'icon.ico' not found. The application will use a default icon.")
            
        self.geometry("950x650")
        
        self.setup_locale()

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure("TLabel", padding=6, font=('Helvetica', 10))
        self.style.configure("TEntry", padding=6, font=('Helvetica', 10))
        self.style.configure("TButton", padding=6, font=('Helvetica', 10, 'bold'))
        self.style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold'))

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        input_frame = ttk.LabelFrame(main_frame, text="Input Data", padding="10")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.asset_name_var = tk.StringVar(value="Sewa Gedung Kantor")
        ttk.Label(input_frame, text="Nama Aset:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.asset_name_var, width=40).grid(row=0, column=1, padx=5, pady=5)

        self.total_cost_var = tk.StringVar()
        ttk.Label(input_frame, text="Harga Perolehan (Rp):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        total_cost_entry = ttk.Entry(input_frame, textvariable=self.total_cost_var, width=40)
        total_cost_entry.grid(row=1, column=1, padx=5, pady=5)
        self.total_cost_entry = total_cost_entry  # Save reference for cursor management
        self.total_cost_trace_id = self.total_cost_var.trace_add('write', self._format_total_cost)

        # Start Date Entry
        self.start_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-01'))
        ttk.Label(input_frame, text="Tgl Mulai (YYYY-MM-DD):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        start_date_entry = ttk.Entry(input_frame, textvariable=self.start_date_var, width=40)
        start_date_entry.grid(row=2, column=1, padx=5, pady=5)
        start_date_entry.bind('<FocusOut>', lambda event: self._format_date_entry(self.start_date_var))

        # End Date Entry
        self.end_date_var = tk.StringVar(value=(datetime.now() + relativedelta(years=1, days=-1)).strftime('%Y-%m-%d'))
        ttk.Label(input_frame, text="Tgl Selesai (YYYY-MM-DD):").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        end_date_entry = ttk.Entry(input_frame, textvariable=self.end_date_var, width=40)
        end_date_entry.grid(row=3, column=1, padx=5, pady=5)
        end_date_entry.bind('<FocusOut>', lambda event: self._format_date_entry(self.end_date_var))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=15, pady=10)

        ttk.Button(button_frame, text="Hitung & Simpan ke Excel", command=self.run_calculation).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Keluar", command=self.quit).pack(side=tk.RIGHT)

        output_frame = ttk.LabelFrame(main_frame, text="Jadwal Amortisasi", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(output_frame, show='headings')
        self.tree["columns"] = ('Periode', 'Tanggal', 'Deskripsi', 'Beban Amortisasi', 'Akumulasi Amortisasi', 'Nilai Buku')
        
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            anchor = tk.E if col not in ['Deskripsi', 'Tanggal'] else tk.W
            self.tree.column(col, anchor=anchor, width=140)

        v_scroll = ttk.Scrollbar(output_frame, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(output_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        output_frame.grid_rowconfigure(0, weight=1)
        output_frame.grid_columnconfigure(0, weight=1)

        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _format_date_entry(self, date_var: tk.StringVar):
        """Formats the date in a StringVar to YYYY-MM-DD on focus out."""
        try:
            date_str = date_var.get()
            if date_str:  # only format if not empty
                parsed_date = parse_flexible_date(date_str)
                date_var.set(parsed_date.strftime('%Y-%m-%d'))
        except ValueError:
            # If parsing fails, do nothing. Let the main validation handle it
            # upon calculation to avoid being intrusive during typing.
            pass

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

            if cleaned_val:
                num_val = int(cleaned_val)
                # Format with locale-specific grouping (e.g., '1.000.000')
                formatted_val = locale.format_string("%d", num_val, grouping=True)
                
                self.total_cost_var.set(formatted_val)
                
                # Adjust cursor position based on the change in length
                len_after = len(formatted_val)
                cursor_pos += (len_after - len_before)
                self.total_cost_entry.icursor(max(0, min(cursor_pos, len_after)))

            # If cleaned_val is empty but current_val is not, it means the user deleted everything.
            elif current_val:
                 self.total_cost_var.set("")

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

    def run_calculation(self):
        self.tree.delete(*self.tree.get_children())
        self.status_var.set("")

        try:
            asset_name = self.asset_name_var.get()
            if not asset_name:
                raise ValueError("Nama Aset tidak boleh kosong.")
                
            total_cost_str = self.total_cost_var.get().replace('.', '').replace(',', '.')
            total_cost = Decimal(total_cost_str)

            if total_cost <= 0:
                raise ValueError("Harga perolehan harus lebih besar dari nol.")

            # Use the flexible date parsing function
            start_date = parse_flexible_date(self.start_date_var.get())
            end_date = parse_flexible_date(self.end_date_var.get())
            
            if end_date <= start_date:
                raise ValueError("Tanggal selesai harus setelah tanggal mulai.")

        except InvalidOperation:
            messagebox.showerror("Error Input", "Harga perolehan tidak valid. Harap masukkan angka yang benar (misal: 10.000,50).")
            return
        except ValueError as e:
            # Display the specific error message from parsing or validation
            messagebox.showerror("Error Input", str(e))
            return
        
        schedule_df, error_msg = calculate_amortization_schedule(total_cost, start_date, end_date)

        if error_msg:
            messagebox.showerror("Error Perhitungan", error_msg)
            return

        # Populate the treeview with formatted data
        for index, row in schedule_df.iterrows():
            formatted_values = [
                row['Periode'],
                row['Tanggal'].strftime('%Y-%m-%d'),
                row['Deskripsi'],
                self.format_currency(row['Beban Amortisasi']),
                self.format_currency(row['Akumulasi Amortisasi']),
                self.format_currency(row['Nilai Buku'])
            ]
            self.tree.insert("", "end", values=formatted_values)
        
        self.save_to_excel(schedule_df, asset_name)

    def save_to_excel(self, df: pd.DataFrame, asset_name: str):
        output_filename = f'Hasil Amortisasi {asset_name}.xlsx'
        try:
            # Create a new DataFrame for excel export
            df_for_export = df.copy()

            # Convert numbers to localized strings for Excel output.
            # This stores them as TEXT, making them unusable for formulas in Excel.
            for col in ['Beban Amortisasi', 'Akumulasi Amortisasi', 'Nilai Buku']:
                df_for_export[col] = df_for_export[col].apply(self.format_currency)
            
            df_for_export['Tanggal'] = df_for_export['Tanggal'].dt.strftime('%Y-%m-%d')

            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                df_for_export.to_excel(writer, sheet_name='Jadwal Amortisasi', index=False)
            
            self.status_var.set(f"Sukses! Jadwal disimpan ke '{output_filename}'")
            messagebox.showinfo("Sukses", f"Jadwal amortisasi telah disimpan ke file:\n{output_filename}")
        except Exception as e:
            self.status_var.set(f"Error: Gagal menyimpan file Excel.")
            messagebox.showerror("Error Penyimpanan", f"Gagal menyimpan file Excel: {e}")

