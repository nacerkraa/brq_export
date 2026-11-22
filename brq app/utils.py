"""Shared date utilities."""
import datetime as dt

def parse_date(value):
    try:
        return dt.datetime.strptime(value.strip(), "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError("Date invalide. Utilisez le format AAAA-MM-JJ.") from exc
