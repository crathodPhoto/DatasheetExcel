import os
import pandas as pd
import win32com.client as win32
import time

# -------------------------
# CONFIGURATION - Paths
# -------------------------
destination_folder = r"C:\Users\crathod\Documents\Datasheet Automation\Script Output"
excel_template_path = r"C:\Users\crathod\Documents\Datasheet Automation\Datasheet Graph Template 1.xlsm"

os.makedirs(destination_folder, exist_ok=True)

# -------------------------
# Load Devices.xlsx
# -------------------------
devices_file = os.path.join(destination_folder, "Devices.xlsx")
print(f"Looking for devices file at: {devices_file}")
print(f"File exists: {os.path.exists(devices_file)}")

if not os.path.exists(devices_file):
    print("ERROR: Devices.xlsx not found! Make sure to run DatasheetAutomationPart1FINAL.py first.")
    exit(1)

devices_df = pd.read_excel(devices_file, sheet_name="Devices")
print(f"Original devices_df shape: {devices_df.shape}")
print(f"Original devices_df columns: {devices_df.columns.tolist()}")
print(f"First few rows:\n{devices_df.head()}")

# Only drop rows where Lot_ID or Dev# are missing (SN and SKU can be empty initially)
devices_df = devices_df.dropna(subset=["Lot_ID", "Dev#"])
print(f"After dropna, devices_df shape: {devices_df.shape}")

if devices_df.empty:
    print("ERROR: No valid devices found in Devices.xlsx after filtering!")
    print("Check that the Excel file has data in the required columns: Lot_ID, Dev#")
    exit(1)

# Fill NaN values in SN and SKU with empty strings for processing
devices_df["SN"] = devices_df["SN"].fillna("")
devices_df["SKU"] = devices_df["SKU"].fillna("")

# -------------------------
# Helpers
# -------------------------
def find_device_file(lot_id, dev_num, phrase):
    for filename in os.listdir(os.path.join(destination_folder, "Other")):
        if lot_id in filename and dev_num in filename and phrase in filename:
            return os.path.join(destination_folder, "Other", filename)
    return None

def paste_text_file_fast(sheet, start_row, start_column, lot_id, dev_num, phrase):
    file_path = find_device_file(lot_id, dev_num, phrase)
    if not file_path:
        print(f"File not found for {phrase} - skipping.")
        return None

    with open(file_path, 'r') as f:
        lines = [line.strip() for line in f.readlines()]

    data = [line.split() for line in lines]
    print(f"Sample data for {phrase}: {data[:5]}")  # Show first 5 rows for inspection

    num_rows = len(data)
    num_cols = max(len(row) for row in data)

    for row in data:
        while len(row) < num_cols:
            row.append("")

    start_cell = sheet.Cells(start_row, start_column)
    end_cell = sheet.Cells(start_row + num_rows - 1, start_column + num_cols - 1)
    sheet.Range(start_cell, end_cell).Value = data

    print(f"Fast-pasted {len(data)} rows for {phrase} starting at row {start_row}, column {start_column}")
    return file_path

def clear_old_data(sheet):
    """Clear specific rows before pasting new data."""
    print("Clearing old data in rows 69-72 and 40-43...")

    # Clear rows 69 to 72 (all columns)
    sheet.Range(f"A69:CB72").ClearContents()

    # Clear rows 40 to 43 (all columns)
    sheet.Range(f"A40:CB43").ClearContents()

    print("Old data cleared.\n")

def log_axis_values(sheet):
    """Log all axis control values for inspection."""
    values = {
        "Primary X Min (D7)": sheet.Cells(7, 4).Value,
        "Primary X Max (E7)": sheet.Cells(7, 5).Value,
        "Primary Y Min (D13)": sheet.Cells(13, 4).Value,
        "Primary Y Max (E13)": sheet.Cells(13, 5).Value,
        "Secondary Y Min (D10)": sheet.Cells(10, 4).Value,
        "Secondary Y Max (E10)": sheet.Cells(10, 5).Value,
    }
    for name, value in values.items():
        print(f"{name}: {value}")
    return values

def update_chart_axes(sheet, chart, chart_number):
    """Updates and logs axis settings for the specified chart."""
    log_axis_values(sheet)

    if chart_number == 1:
        axes_config = {
            'primary_y': (sheet.Cells(13, 4).Value, sheet.Cells(13, 5).Value),
            'primary_x': (sheet.Cells(7, 4).Value, sheet.Cells(7, 5).Value),
            'secondary_y': (sheet.Cells(10, 4).Value, sheet.Cells(10, 5).Value)
        }
    elif chart_number == 2:
        axes_config = {
            'primary_y': (sheet.Cells(8, 4).Value, sheet.Cells(8, 5).Value),
            'primary_x': (sheet.Cells(12, 4).Value, sheet.Cells(12, 5).Value),
            'secondary_y': (sheet.Cells(9, 4).Value, sheet.Cells(9, 5).Value)
        }
    else:
        raise ValueError(f"Invalid chart number: {chart_number}")

    print(f"\n--- Setting axes for Chart{chart_number} ---")

    try:
        chart.Parent.Activate()

        y_min, y_max = axes_config['primary_y']
        x_min, x_max = axes_config['primary_x']
        sy_min, sy_max = axes_config['secondary_y']

        print(f"  Primary Y Axis: Min={y_min}, Max={y_max}")
        chart.Axes(2).MinimumScale = y_min
        chart.Axes(2).MaximumScale = y_max

        print(f"  Primary X Axis: Min={x_min}, Max={x_max}")
        chart.Axes(1).MinimumScale = x_min
        chart.Axes(1).MaximumScale = x_max

        print(f"  Secondary Y Axis: Min={sy_min}, Max={sy_max}")
        try:
            chart.Axes(2, 2).MinimumScale = sy_min
            chart.Axes(2, 2).MaximumScale = sy_max
        except Exception:
            print("  (No secondary Y axis present, skipping)")

        print(f"Updated axes for Chart{chart_number}\n")

    except Exception as e:
        print(f"Failed to set axes for Chart{chart_number}: {e}")

# -------------------------
# Process First Device in Diagnosis Mode
# -------------------------
excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # Show Excel for direct inspection

first_device = devices_df.iloc[0]
lot_id = str(first_device["Lot_ID"]).strip()
dev_num = str(first_device["Dev#"]).strip()
# Handle SN: if it's empty or NaN, keep it as empty string, otherwise convert to int then string
sn_raw = first_device["SN"]
if pd.isna(sn_raw) or str(sn_raw).strip() == "":
    sn = ""
else:
    sn = str(int(sn_raw)).strip()
sku = str(first_device["SKU"]).strip()

try:
    wb = excel.Workbooks.Open(excel_template_path)
    sheet = wb.Sheets("snl")
except Exception as e:
    print(f"Error opening Excel template: {e}")
    print("This might be because the file is already open in Excel.")
    print("Please close all Excel files and try again, or check if the file is in Protected View.")
    excel.Quit()
    exit(1)

print(f"Processing Device: Lot={lot_id}, Dev={dev_num}, SN={sn}, SKU={sku}")

# Clear old data
clear_old_data(sheet)

# Paste Data
paste_text_file_fast(sheet, 17, 1, lot_id, dev_num, "WLT_Wave")
paste_text_file_fast(sheet, 46, 1, lot_id, dev_num, "WLT_SMSR")
paste_text_file_fast(sheet, 79, 1, lot_id, dev_num, "LIV_vs_Temp")

# Set SKU and recalculate
sheet.Cells(1, 2).Value = sku
sheet.Calculate()
excel.CalculateFull()

# Work with charts
charts_sheet = wb.Sheets("Charts")
chart1 = charts_sheet.ChartObjects("Chart1").Chart
chart2 = charts_sheet.ChartObjects("Chart2").Chart

update_chart_axes(sheet, chart1, 1)
update_chart_axes(sheet, chart2, 2)

# Pause so you can inspect everything
print("\nPausing for manual inspection in Excel. Charts should now reflect pasted data.")
print("Review all axis settings, pasted data, and charts.")
print("Press Ctrl+C to end script after review.")

while True:
    time.sleep(10)
