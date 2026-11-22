"""BRQ report queries and cumulative calculations."""
import datetime as dt
import pandas as pd
from config import ANNUAL_OBJECTIVE, RECAP_QUERY

def fetch_period_totals(connection, start_date, end_date):
    """Return validated category quantities for one inclusive date period."""
    frame = pd.read_sql(
        RECAP_QUERY,
        connection,
        params={"date1": start_date, "date2": end_date}
    )
    categories = ["CHAIR", "PONTE", "RUMINANTS", "DIVERS"]
    if frame.empty:
        return {category: 0.0 for category in categories}
    row = frame.iloc[0]
    return {
        category: (0.0 if pd.isna(row.get(category)) else float(row[category]))
        for category in categories
    }


def build_cumulative_summary(connection, reference_date):
    """
    Calculate daily, month-to-date and year-to-date totals.

    All three periods are anchored to the selected report end date.
    """
    month_start = reference_date.replace(day=1)
    year_start = dt.date(reference_date.year, 1, 1)

    daily = fetch_period_totals(connection, reference_date, reference_date)
    monthly = fetch_period_totals(connection, month_start, reference_date)
    yearly = fetch_period_totals(connection, year_start, reference_date)

    categories = ["CHAIR", "PONTE", "RUMINANTS", "DIVERS"]
    rows = [
        {
            "Catégorie": category,
            "Jour": daily[category],
            "Cumul mois": monthly[category],
            "Cumul année": yearly[category],
            "Objectif annuel (%)": None,
        }
        for category in categories
    ]
    annual_total = sum(yearly.values())
    rows.append({
        "Catégorie": "TOTAL",
        "Jour": sum(daily.values()),
        "Cumul mois": sum(monthly.values()),
        "Cumul année": annual_total,
        "Objectif annuel (%)": annual_total / ANNUAL_OBJECTIVE * 100,
    })
    return pd.DataFrame(rows)
