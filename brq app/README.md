# BRQ Export Manager

## Run

From the repository root:

```powershell
py brq_app.py
```

Or double-click `brq_app.pyw` to launch without a console window.

Install dependencies with:

```powershell
py -m pip install -r "brq app/requirements.txt"
```

## Source layout

- `config.py`: paths, database settings, SQL queries and annual objective.
- `database.py`: MySQL connection and SQL backup import.
- `report_data.py`: cumulative calculations and report data access.
- `exporters.py`: Excel, PDF and PNG generation.
- `gui.py`: Tkinter interface and interactions.
- `utils.py`: shared date utilities.
- `legacy/`: previous standalone scripts retained for reference.
