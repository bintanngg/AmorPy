import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, getcontext

from calendar import monthrange
# Set precision for Decimal calculations
getcontext().prec = 28

def _calculate_straight_line_schedule(total_cost: Decimal, start_date: datetime, end_date: datetime, num_months: int):
    """Calculates the amortization schedule using the straight-line method."""
    monthly_amortization = (total_cost / Decimal(num_months)).quantize(Decimal("0.01")) if num_months > 0 else Decimal("0")
    
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
    
    # Create schedule for each month
    for i in range(1, num_months + 1):
        # Calculate date for the current period
        target_date = start_date + relativedelta(months=i-1)
        # Handle end-of-month days (e.g., Jan 31 -> Feb 28/29)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)

        amortization_this_month = monthly_amortization
        # On the last month, adjust the expense to exactly zero out the book value, handling any rounding differences.
        if i == num_months:
            amortization_this_month = book_value

        accumulated_amortization += amortization_this_month
        book_value -= amortization_this_month
            
        schedule_data.append([
            i,
            current_date,
            'Amortization',
            amortization_this_month,
            accumulated_amortization,
            book_value if book_value > Decimal("0") else Decimal("0.00")
        ])
        
    return schedule_data

def _calculate_declining_balance_schedule(total_cost: Decimal, start_date: datetime, end_date: datetime, num_months: int):
    """Calculates the amortization schedule using the double-declining balance method."""
    if num_months == 0:
        return []

    # Monthly depreciation rate for double-declining balance (2 / useful life in months)
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

    # Create schedule for each month
    for i in range(1, num_months + 1):
        # Calculate date for the current period
        target_date = start_date + relativedelta(months=i-1)
        # Handle end-of-month days (e.g., Jan 31 -> Feb 28/29)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)
        
        remaining_life = num_months - i + 1
        current_book_value = book_value # Use book_value from previous period

        # Ensure book value is zero at the end of the period
        if i == num_months:
            amortization_this_month = current_book_value
        else:
            # Calculate depreciation using both DDB and Straight-Line on remaining value
            ddb_amortization = (current_book_value * depreciation_rate).quantize(Decimal("0.01"))
            
            # Straight-line depreciation on the remaining book value over the remaining life
            sl_amortization = (current_book_value / Decimal(remaining_life)).quantize(Decimal("0.01"))

            # Switch to straight-line if it provides a larger depreciation expense
            amortization_this_month = max(ddb_amortization, sl_amortization)

            # Do not let the amortization expense exceed the remaining book value.
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

def _calculate_soyd_schedule(total_cost: Decimal, start_date: datetime, end_date: datetime, num_months: int):
    """Calculates the amortization schedule using the Sum-of-the-Years' Digits (SOYD) method."""
    if num_months == 0:
        return []

    # Calculate the Sum-of-the-Years' Digits (for months)
    soyd = Decimal(num_months * (num_months + 1)) / Decimal(2)

    schedule_data = []
    accumulated_amortization = Decimal("0.00")
    book_value = total_cost

    # Initial row (Period 0)
    schedule_data.append([
        0,
        start_date,
        'Initial Value',
        Decimal("0.00"),
        Decimal("0.00"),
        book_value
    ])

    # Create schedule for each month
    for i in range(1, num_months + 1):
        # Calculate date for the current period
        target_date = start_date + relativedelta(months=i-1)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)

        # On the last month, the amortization is simply the remaining book value.
        if i == num_months:
            amortization_this_month = book_value
        else:
            # Remaining life in months
            remaining_life = Decimal(num_months - i + 1)
            # Amortization for this month
            amortization_this_month = (total_cost * (remaining_life / soyd)).quantize(Decimal("0.01"))

        # Ensure expense does not exceed the book value (important for rounding)
        if amortization_this_month > book_value and i != num_months:
            amortization_this_month = book_value

        accumulated_amortization += amortization_this_month
        book_value -= amortization_this_month

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
    """Calculates the amortization schedule based on the selected method."""
    num_months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) + 1
    
    if num_months <= 0:
        return None, "Invalid lease period. The end date must be after the start date."

    if method == 'Straight-Line':
        schedule_data = _calculate_straight_line_schedule(total_cost, start_date, end_date, num_months)
    elif method == 'Double Declining Balance':
        schedule_data = _calculate_declining_balance_schedule(total_cost, start_date, end_date, num_months)
    elif method == "Sum-of-the-Years Digits":
        schedule_data = _calculate_soyd_schedule(total_cost, start_date, end_date, num_months)
    else:
        return None, f"Unknown amortization method: {method}"

    columns = [
        'Period', 'Date', 'Description', 'Amortization Expense',
        'Accumulated Amortization', 'Book Value'
    ]
    
    schedule_df = pd.DataFrame(schedule_data, columns=columns)
    
    return schedule_df, None