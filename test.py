import pandas as pd
from datetime import datetime

def timestamp_to_date_list(timestamp):
    timestamp = timestamp[:10]
    date = datetime.strptime(timestamp.strip(), '%d.%m.%Y')
    return [date.day, date.month, date.year]

print(timestamp_to_date_list('15.09.2026'))