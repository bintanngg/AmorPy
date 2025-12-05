import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, getcontext

from calendar import monthrange
# Set precision for Decimal calculations
getcontext().prec = 28

def _calculate_straight_line_schedule(total_cost: Decimal, start_date: datetime, end_date: datetime, num_months: int):
    """
    Menghitung jadwal amortisasi menggunakan metode garis lurus.
    """
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
        # Hitung tanggal untuk periode saat ini
        target_date = start_date + relativedelta(months=i-1)
        # Menangani hari di akhir bulan (misal: 31 Jan -> 28/29 Feb)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)

        
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
        
    return schedule_data

def _calculate_declining_balance_schedule(total_cost: Decimal, start_date: datetime, end_date: datetime, num_months: int):
    """
    Menghitung jadwal amortisasi menggunakan metode saldo menurun ganda.
    """
    if num_months == 0:
        return []

    # Tingkat penyusutan bulanan untuk saldo menurun ganda
    # (2 / masa manfaat dalam bulan)
    depreciation_rate = Decimal(2) / Decimal(num_months)

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
        # Hitung tanggal untuk periode saat ini
        target_date = start_date + relativedelta(months=i-1)
        # Menangani hari di akhir bulan (misal: 31 Jan -> 28/29 Feb)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)
        
        current_book_value = total_cost - accumulated_amortization

        # Memastikan nilai buku adalah nol pada akhir periode
        if i == num_months:
            amortization_this_month = current_book_value
        else:
            # Amortisasi untuk bulan ini
            amortization_this_month = (current_book_value * depreciation_rate).quantize(Decimal("0.01"))

            # Jangan biarkan beban amortisasi melebihi sisa nilai buku.
            # Ini penting jika tingkat penyusutan sangat tinggi pada aset dengan masa manfaat pendek.
            if amortization_this_month > current_book_value:
                amortization_this_month = current_book_value

        accumulated_amortization += amortization_this_month
        book_value -= amortization_this_month

        # Ensure book value doesn't become a tiny negative number on the last run
        if i == num_months:
            book_value = Decimal("0.00")
        
        schedule_data.append([
            i,
            current_date,
            'Amortization',
            amortization_this_month,
            accumulated_amortization,
            book_value if book_value > Decimal("0") else Decimal("0.00")
        ])

    return schedule_data

def calculate_amortization_schedule(total_cost: Decimal, start_date: datetime, end_date: datetime, method: str):
    """
    Menghitung jadwal amortisasi berdasarkan metode yang dipilih.
    Mengembalikan DataFrame dengan tipe data numerik mentah.
    """
    num_months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) + 1
    
    if num_months <= 0:
        return None, "Periode sewa tidak valid. Tanggal selesai harus setelah tanggal mulai."

    if method == 'Straight-Line':
        schedule_data = _calculate_straight_line_schedule(total_cost, start_date, end_date, num_months)
    elif method == 'Double Declining Balance':
        schedule_data = _calculate_declining_balance_schedule(total_cost, start_date, end_date, num_months)
    else:
        return None, f"Metode amortisasi tidak diketahui: {method}"

    columns = [
        'Periode', 'Tanggal', 'Deskripsi',
        'Beban Amortisasi', 'Akumulasi Amortisasi', 'Nilai Buku'
    ]
    
    schedule_df = pd.DataFrame(schedule_data, columns=columns)
    
    return schedule_df, None
