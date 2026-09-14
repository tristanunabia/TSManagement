import datetime
import calendar

WEEKDAY_MAP = {
    0: "M",   # Monday
    1: "T",   # Tuesday
    2: "W",   # Wednesday
    3: "TH",  # Thursday
    4: "F",   # Friday
    5: "S",   # Saturday
    6: "SU"   # Sunday
}

def get_previous_month(year: int, month: int) -> tuple[int, int]:
    """
    Given a year and month, return the (year, month) of the previous month.
    """
    if month == 1:
        return year - 1, 12
    else:
        return year, month - 1

def generate_cutoff_dates(year: int, month: int, cutoff_type: str) -> list[dict]:
    """
    Generates a list of dictionaries for each date in the cutoff period.
    Each dictionary has keys:
      - 'date': datetime.date object
      - 'day_num': day of the month string or int (e.g. 26, 27, ..., 1, 2, ..., 10)
      - 'day_abbr': custom weekday abbreviation
    
    Cutoff types:
      - '11-25': Covers the 11th through the 25th of the selected target month.
      - '26-10': Covers the 26th of the previous month through the 10th of the target month.
    """
    if cutoff_type == "11-25":
        start_date = datetime.date(year, month, 11)
        end_date = datetime.date(year, month, 25)
    elif cutoff_type == "26-10":
        prev_year, prev_month = get_previous_month(year, month)
        # Determine the number of days in the previous month to ensure 26 is valid
        # (e.g. February has 28 or 29 days, etc., 26 is always valid for all months)
        start_date = datetime.date(prev_year, prev_month, 26)
        end_date = datetime.date(year, month, 10)
    else:
        raise ValueError(f"Invalid cutoff type: {cutoff_type}. Must be '11-25' or '26-10'.")

    dates = []
    current_date = start_date
    while current_date <= end_date:
        dates.append({
            'date': current_date,
            'day_num': current_date.day,
            'day_abbr': WEEKDAY_MAP[current_date.weekday()]
        })
        current_date += datetime.timedelta(days=1)
        
    return dates
