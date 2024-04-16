from datetime import datetime


# Formats an integer amount into currency format with the Korean Won symbol.
def format_cost(amount):
    return f"{amount:,.0f}"


# Convert Date Value into format on SQL
def format_date(date_str: str):
    # 1. '20230513'
    # 2. '230513'
    # 3. '2023/05/13'
    # 4. '2023-05-13'
    # 5. '23/05/13'
    # 6. '23-05-13'
    # 7. '23/5/13'
    # 8. '23-5-13'

    # All input cases(#1~#8) should be converted like '1996-05-13'
    # But if the date value s not valid, return None

    # Remove any non-digit characters
    digits = ''.join(filter(str.isdigit, date_str))

    # Determine the year, month, and day based on the length of the digits
    if len(digits) == 8:  # Formats like '20230513'
        year, month, day = int(digits[:4]), int(digits[4:6]), int(digits[6:])
    elif len(digits) == 6:  # Formats like '230513'
        year, month, day = int(digits[:2]), int(digits[2:4]), int(digits[4:])
        # Convert two-digit year to four-digit year
        year += 2000 if year < 50 else 1900
    else:
        return None  # Invalid format

    # Validate and format the date
    try:
        formatted_date = datetime(year, month, day).strftime('%Y-%m-%d')
        return formatted_date
    except ValueError:
        return None  # Invalid date


# Convert the string to a datetime object. The string should be in the format 'YYYY-MM-DD'.
def day_of_week(date_str):
    date = datetime.strptime(date_str, '%Y-%m-%d')
    days = ["Mon", "Tue", "Wed", "Thurs", "Fri", "Sat", "Sun"]
    return days[date.weekday()]
