import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, getcontext
from calendar import monthrange

# Set precision for Decimal calculations
getcontext().prec = 28

def _create_schedule_row(period, date, description, expense, accumulated, book_value):
    return [period, date, description, expense, accumulated, book_value]

def _calculate_straight_line_schedule(total_cost: Decimal, salvage_value: Decimal, start_date: datetime, num_months: int):
    """Calculates the amortization schedule using the straight-line method."""
    depreciable_base = total_cost - salvage_value
    
    if num_months <= 0:
        return []
        
    monthly_amortization = (depreciable_base / Decimal(num_months)).quantize(Decimal("0.01"))
    
    schedule_data = []
    accumulated_amortization = Decimal("0.00")
    book_value = total_cost
    
    # Initial row
    schedule_data.append(_create_schedule_row(
        0, start_date, 'Initial Value', Decimal("0.00"), Decimal("0.00"), book_value
    ))
    
    for i in range(1, num_months + 1):
        target_date = start_date + relativedelta(months=i-1)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)

        amortization_this_month = monthly_amortization
        
        # Adjust last month to match salvage value exactly
        if i == num_months:
            amortization_this_month = book_value - salvage_value
            
        # Safety check: don't depreciate below salvage
        if book_value - amortization_this_month < salvage_value:
             amortization_this_month = book_value - salvage_value

        accumulated_amortization += amortization_this_month
        book_value -= amortization_this_month
            
        schedule_data.append(_create_schedule_row(
            i, current_date, 'Amortization', amortization_this_month, accumulated_amortization, book_value
        ))
        
    return schedule_data

def _calculate_declining_balance_schedule(total_cost: Decimal, salvage_value: Decimal, start_date: datetime, num_months: int):
    """Calculates the amortization schedule using the double-declining balance method."""
    if num_months == 0:
        return []

    # Monthly depreciation rate for DDB (2 / useful life in months)
    depreciation_rate = Decimal(2) / Decimal(num_months)

    schedule_data = []
    accumulated_amortization = Decimal("0.00")
    book_value = total_cost

    # Initial row
    schedule_data.append(_create_schedule_row(
        0, start_date, 'Initial Value', Decimal("0.00"), Decimal("0.00"), book_value
    ))

    for i in range(1, num_months + 1):
        target_date = start_date + relativedelta(months=i-1)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)
        
        current_book_value = book_value

        # Stop if already at or below salvage value
        if current_book_value <= salvage_value:
             amortization_this_month = Decimal("0.00")
        else:
            # DDB Calculation
            amortization_this_month = (current_book_value * depreciation_rate).quantize(Decimal("0.01"))
            
            # Switch to Straight Line check? 
            # Standard DDB often switches to SL when SL > DDB to ensure full depreciation.
            # Using remaining life and remaining depreciable base.
            remaining_life = num_months - i + 1
            if remaining_life > 0:
                remaining_depreciable = current_book_value - salvage_value
                sl_amortization = (remaining_depreciable / Decimal(remaining_life)).quantize(Decimal("0.01"))
                amortization_this_month = max(amortization_this_month, sl_amortization)

            # Cap so we don't go below salvage
            if current_book_value - amortization_this_month < salvage_value:
                amortization_this_month = current_book_value - salvage_value

        accumulated_amortization += amortization_this_month
        book_value -= amortization_this_month

        schedule_data.append(_create_schedule_row(
            i, current_date, 'Amortization', amortization_this_month, accumulated_amortization, book_value
        ))

    return schedule_data

def _calculate_soyd_schedule(total_cost: Decimal, salvage_value: Decimal, start_date: datetime, num_months: int):
    """Calculates the amortization schedule using the Sum-of-the-Years' Digits (SOYD) method."""
    if num_months == 0:
        return []

    soyd = Decimal(num_months * (num_months + 1)) / Decimal(2)
    depreciable_base = total_cost - salvage_value

    schedule_data = []
    accumulated_amortization = Decimal("0.00")
    book_value = total_cost

    # Initial row
    schedule_data.append(_create_schedule_row(
        0, start_date, 'Initial Value', Decimal("0.00"), Decimal("0.00"), book_value
    ))

    for i in range(1, num_months + 1):
        target_date = start_date + relativedelta(months=i-1)
        _, days_in_month = monthrange(target_date.year, target_date.month)
        day = min(start_date.day, days_in_month)
        current_date = target_date.replace(day=day)

        if book_value <= salvage_value:
             amortization_this_month = Decimal("0.00")
        else:
            remaining_life = Decimal(num_months - i + 1)
            amortization_this_month = (depreciable_base * (remaining_life / soyd)).quantize(Decimal("0.01"))

            # Handing rounding on the last month or if exceeding salvage
            if book_value - amortization_this_month < salvage_value:
                amortization_this_month = book_value - salvage_value

        accumulated_amortization += amortization_this_month
        book_value -= amortization_this_month

        schedule_data.append(_create_schedule_row(
            i, current_date, 'Amortization', amortization_this_month, accumulated_amortization, book_value
        ))

    return schedule_data

def calculate_amortization_schedule(total_cost: Decimal, salvage_value: Decimal, start_date: datetime, end_date: datetime, method: str):
    """Calculates the amortization schedule based on the selected method."""
    num_months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) + 1
    
    if num_months <= 0:
        return None, "Invalid lease period. The end date must be after the start date."
        
    if salvage_value >= total_cost:
         return None, "Salvage value cannot be greater than or equal to the total cost."

    if method == 'Straight-Line':
        schedule_data = _calculate_straight_line_schedule(total_cost, salvage_value, start_date, num_months)
    elif method == 'Double Declining Balance':
        schedule_data = _calculate_declining_balance_schedule(total_cost, salvage_value, start_date, num_months)
    elif method == "Sum-of-the-Years Digits":
        schedule_data = _calculate_soyd_schedule(total_cost, salvage_value, start_date, num_months)
    else:
        return None, f"Unknown amortization method: {method}"

    columns = [
        'Period', 'Date', 'Description', 'Amortization Expense',
        'Accumulated Amortization', 'Book Value'
    ]
    
    schedule_df = pd.DataFrame(schedule_data, columns=columns)
    
    return schedule_df, None
