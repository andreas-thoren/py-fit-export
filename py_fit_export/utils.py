from datetime import datetime, time, timedelta, timezone

CET = timezone(timedelta(hours=1))

def excel_safe_datetime(v):
    if isinstance(v, datetime) and v.tzinfo is not None:
        return v.astimezone(CET).replace(tzinfo=None)
    if isinstance(v, time) and v.tzinfo is not None:
        return v.replace(tzinfo=None)
    return v
