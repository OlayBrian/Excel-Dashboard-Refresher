import win32com.client
import time
import os

folder_path = r"Q:\Your_Folder_Path"

file_filter = lambda name: name.lower().endswith(".xlsx") and not name.startswith("~$")

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

def is_still_refreshing(wb):
    # Check if any QueryTables are still refreshing
    for sheet in wb.Sheets:
        for lo in sheet.ListObjects:
            try:
                if lo.QueryTable.Refreshing:
                    return True
            except Exception:
                continue
    # Check if any Workbook Connections (Power Query) are still refreshing
    for conn in wb.Connections:
        try:
            if conn.OLEDBConnection.Refreshing:
                return True
        except Exception:
            continue
    return False

for filename in os.listdir(folder_path):
    if not file_filter(filename):
        continue

    filepath = os.path.join(folder_path, filename)
    print(f"\nOpening: {filename}")
    
    try:
        wb = excel.Workbooks.Open(filepath)
        wb.RefreshAll()
        print("Triggered RefreshAll...")

        # Wait until everything finishes refreshing
        while is_still_refreshing(wb):
            print("Refreshing...")
            time.sleep(1)

        wb.Save()
        print(f"Saved: {filename}")
        wb.Close(SaveChanges=False)

    except Exception as e:
        print(f"Error with file {filename}: {e}")

excel.Quit()
print("\nAll files processed. Excel closed.")
