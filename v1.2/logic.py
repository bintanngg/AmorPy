import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, getcontext

# Set precision for Decimal calculations
getcontext().prec = 28

def calculate_amortization_schedule(total_cost: Decimal, start_date: datetime, end_date: datetime):
    """
    Menghitung jadwal amortisasi menggunakan tipe data Decimal untuk akurasi finansial.
    Mengembalikan DataFrame dengan tipe data numerik mentah.
    """
    num_months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) + 1
    
    if num_months <= 0:
        return None, "Periode sewa tidak valid. Tanggal selesai harus setelah tanggal mulai."
        
    monthly_amortization = (total_cost / Decimal(num_months)).quantize(Decimal("0.01"))
    
    schedule_data = []
    accumulated_amortization = Decimal("0.00")
    book_value = total_cost
    
    # Baris awal (Periode 0)
    schedule_data.append([
        0,
        start_date,
        'Initial Value',
        Decimal("0.00"),
        Decimal("0.00"),
        book_value
    ])
    
    # Membuat jadwal untuk setiap bulan
    for i in range(1, num_months + 1):
        current_date = start_date + relativedelta(months=i-1)
        
        # Memastikan nilai buku adalah nol pada akhir periode
        if i == num_months:
            # Sisa amortisasi untuk bulan terakhir untuk memastikan nilai buku menjadi 0
            amortization_this_month = total_cost - accumulated_amortization
            accumulated_amortization = total_cost
            book_value = Decimal("0.00")
        else:
            amortization_this_month = monthly_amortization
            accumulated_amortization += monthly_amortization
            book_value -= monthly_amortization
            
        schedule_data.append([
            i,
            current_date,
            'Amortization',
            amortization_this_month,
            accumulated_amortization,
            book_value
        ])
        
    columns = [
        'Periode', 'Tanggal', 'Deskripsi',
        'Beban Amortisasi', 'Akumulasi Amortisasi', 'Nilai Buku'
    ]
    
    schedule_df = pd.DataFrame(schedule_data, columns=columns)
    
    return schedule_df, None
