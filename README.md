# Excel Dashboard Auto-Refresher

A Python automation script that opens Excel workbooks, refreshes all data connections, and saves them — hands-free.

Useful for keeping Power Query dashboards or OLEDB-connected reports up to date on a schedule without manually opening each file.

---

## What It Does

For every `.xlsx` file in a target folder, the script will:

1. Open the workbook using a background Excel instance
2. Trigger `RefreshAll()` to update all data connections
3. Wait for all queries to finish (Power Query, QueryTables, OLEDB)
4. Save the file and close it
5. Move on to the next file

---

## Requirements

- Windows (uses `win32com` to drive Excel)
- Microsoft Excel installed
- Python 3.x
- `pywin32` library

Install the dependency with:

```bash
pip install pywin32
```

---

## Setup

1. Clone the repo:
   ```bash
   git clone https://github.com/your-username/excel-dashboard-refresher.git
   cd excel-dashboard-refresher
   ```

2. Set your target folder path in `refresh_dashboards.py`:
   ```python
   folder_path = r"C:\Your\Dashboard\Folder"
   ```

---

## Usage

```bash
python refresh_dashboards.py
```

Excel will open visibly, process each file one by one, and close when done. Progress is printed to the console.

---

## Automating on a Schedule

To run this every morning automatically, use **Windows Task Scheduler**:

1. Open Task Scheduler → Create Basic Task
2. Set your trigger (e.g., daily at 7:00 AM)
3. Action: `Start a Program`
   - Program: `python`
   - Arguments: `C:\path\to\refresh_dashboards.py`

---

## Notes

- Files starting with `~$` are skipped (Excel temp/lock files)
- `DisplayAlerts` is disabled to prevent Excel pop-ups from blocking the script
- The script polls every second to confirm refreshes are truly complete before saving — `RefreshAll()` is asynchronous and returns immediately without this check
