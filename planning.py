import datetime

import datetime

SHIFTS = [
    ("07:00", "15:15"),
    ("15:15", "23:00"),
    ("23:00", "07:00")
]

SHIFT_DURATION_MIN = {
    0: (15 * 60 + 15) - (7 * 60),       # 495 min
    1: (23 * 60) - (15 * 60 + 15),      # 465 min
    2: (24 * 60 + 7 * 60) - (23 * 60),  # 480 min (shift nuit)
}

PAUSE_MIN = 45

PERSONS = ["Alice", "Bob", "Charlie", "David"]

PERSON_COLORS = {
    "Alice": "#9c99ff",
    "Bob": "#99FF99",
    "Charlie": "#9999FF",
    "David": "#FFCC99"
}


def daterange(start_date, end_date):
    for n in range(int((end_date - start_date).days) + 1):
        yield start_date + datetime.timedelta(n)

def get_weekdays_of_month(year, month):
    start = datetime.date(year, month, 1)
    if month == 12:
        end = datetime.date(year + 1, 1, 1) - datetime.timedelta(1)
    else:
        end = datetime.date(year, month + 1, 1) - datetime.timedelta(1)
    return [d for d in daterange(start, end)]
