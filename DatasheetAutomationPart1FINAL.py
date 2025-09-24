import os
import shutil
import pandas as pd
import subprocess
import re

# -------------------------
# CONFIGURATION - Set your paths here
# -------------------------
source_folder = r"C:\Users\crathod\Documents\Datasheet Automation\Paste Raw Data HERE"  # Folder A - Source folder
destination_folder = r"C:\Users\crathod\Documents\Datasheet Automation\Script Output"  # Folder B - Destination folder
fix_script_path = r"C:\Users\crathod\Documents\Datasheet Automation\CountandFixtxtfiles.py"  # Full path to the FIX script

# Create subfolders in destination folder
liv_folder = os.path.join(destination_folder, "LIV")
smsr_folder = os.path.join(destination_folder, "SMSR")
other_folder = os.path.join(destination_folder, "Other")

os.makedirs(liv_folder, exist_ok=True)
os.makedirs(smsr_folder, exist_ok=True)
os.makedirs(other_folder, exist_ok=True)

# -------------------------
# SKU LOOKUP FUNCTIONALITY
# -------------------------
def load_sku_lookup_table():
    """Load the SKU lookup table from the Excel template"""
    try:
        template_path = r"C:\Users\crathod\Documents\Datasheet Automation\Datasheet Graph Template 1.xlsm"
        key_df = pd.read_excel(template_path, sheet_name='Key')
        print(f"Loaded SKU lookup table with {len(key_df)} entries")
        return key_df
    except Exception as e:
        print(f"Warning: Could not load SKU lookup table: {e}")
        return None

def find_sku_for_device(lot_id, dev_num, key_df=None):
    """
    Find the appropriate SKU for a device based on its Lot_ID and Dev#
    """
    if key_df is None:
        return None
        
    # Extract wavelength and device type from Lot_ID
    # Pattern: "795-DBRL051525B-G11X" -> wavelength=795, type=DBRL
    match = re.match(r'(\d+(?:\.\d+)?)-([A-Z]+)', lot_id)
    if not match:
        print(f"Could not parse Lot_ID for SKU lookup: {lot_id}")
        return None
    
    wavelength = float(match.group(1))
    device_type = match.group(2)
    
    # Find matching SKUs in the key sheet
    matching_skus = []
    
    for sku in key_df['SKU'].dropna():
        # Check if SKU starts with the wavelength (allowing for close matches)
        sku_match = re.match(r'(\d+(?:\.\d+)?)', str(sku))
        if sku_match:
            sku_wavelength = float(sku_match.group(1))
            # Allow for close wavelength matches (within 5nm)
            if abs(sku_wavelength - wavelength) <= 5:
                # Check if device type is in the SKU
                if device_type in str(sku):
                    # Calculate priority: prefer more specific matches (longer device type names)
                    # Also prefer closer wavelength matches
                    wavelength_diff = abs(sku_wavelength - wavelength)
                    
                    # Extract the device type part from SKU for specificity scoring
                    sku_device_part = re.search(r'[A-Z]+(?:LITE)?', str(sku))
                    specificity_score = len(sku_device_part.group()) if sku_device_part else 0
                    
                    # Higher specificity score is better (DBRLITE > DBRL)
                    matching_skus.append((sku, wavelength_diff, -specificity_score))
    
    if matching_skus:
        # Sort by wavelength closeness, then by specificity (more specific first), then by name length
        matching_skus.sort(key=lambda x: (x[1], x[2], len(x[0])))
        best_match = matching_skus[0][0]
        print(f"Found SKU for {wavelength}nm {device_type}: {best_match} (from {len(matching_skus)} matches)")
        return best_match
    else:
        print(f"No matching SKU found for wavelength={wavelength}, type={device_type}")
        return None

# Load the SKU lookup table once
sku_lookup_table = load_sku_lookup_table()

# -------------------------
# SECTION 1 - Copy Files & Categorize
# -------------------------
for filename in os.listdir(source_folder):
    source_file = os.path.join(source_folder, filename)
    destination_file = os.path.join(destination_folder, filename)

    # Copy all files to B
    shutil.copy2(source_file, destination_file)

    # Now sort into folders based on file type and name
    if filename.endswith(".jpg"):
        if "LIV_vs_Temp" in filename:
            shutil.move(destination_file, os.path.join(liv_folder, filename))
        elif "SMSR_vs_Temp" in filename:
            shutil.move(destination_file, os.path.join(smsr_folder, filename))
    elif filename.endswith(".txt"):
        shutil.move(destination_file, os.path.join(other_folder, filename))

print("SECTION 1 complete: Files copied and organized.")

# -------------------------
# SECTION 2 - Generate Excel file listing devices in LIV
# -------------------------

# Extract Lot_ID and Dev# from filenames and write to Excel
devices = []
device_set = set()  # To avoid duplicates

def parse_filename(filename):
    """
    Parse filename to extract Lot_ID and Dev#
    Example formats:
    - 795-DBRL051525B-G11X_DryEtch-37-131_0.0900A_LIV_vs_Temp.jpg
    - 852-DBRL051723C-G2X-25-79_0.1500A_LIV_vs_Temp.jpg
    """
    try:
        # Remove the file extension and measurement type suffix
        base_name = filename.replace(".jpg", "")
        base_name = base_name.replace("_LIV_vs_Temp", "").replace("_SpecWidth", "").replace("_Wave-SMSR_vs_Temp", "").replace("_WLT_SMSR", "").replace("_WLT_Wave", "")
        
        # Remove the current measurement part (e.g., _0.0900A, _0.1500A)
        if "_0." in base_name:
            base_name = base_name.split("_0.")[0]
        
        # Handle different patterns
        if "G11X_DryEtch-" in base_name:
            # Pattern: 795-DBRL051525B-G11X_DryEtch-37-131
            parts = base_name.split("_DryEtch-")
            lot_id = parts[0]  # 795-DBRL051525B-G11X
            dev_num = parts[1]  # 37-131
        elif "G2X-" in base_name:
            # Pattern: 852-DBRL051723C-G2X-25-79
            parts = base_name.split("-")
            lot_id = "-".join(parts[:3])  # 852-DBRL051723C-G2X
            dev_num = "-".join(parts[3:5])  # 25-79
        else:
            # Generic fallback - look for pattern ending with X followed by dash and device numbers
            import re
            # Look for pattern: (anything)X-(digits)-(digits)
            match = re.match(r'(.+X)(?:_DryEtch)?-(\d+-\d+)', base_name)
            if match:
                lot_id = match.group(1)
                dev_num = match.group(2)
            else:
                # Try to split by dashes and find reasonable components
                parts = base_name.split("-")
                if len(parts) >= 5:
                    # Assume first 3 parts are lot_id, last 2 are device
                    lot_id = "-".join(parts[:3])
                    dev_num = "-".join(parts[-2:])
                else:
                    raise ValueError("Cannot determine lot_id and dev_num from filename structure")
        
        return lot_id, dev_num
    except Exception as e:
        raise ValueError(f"Parsing failed: {e}")

for filename in os.listdir(liv_folder):
    if filename.endswith(".jpg"):
        try:
            lot_id, dev_num = parse_filename(filename)
            
            # Create a unique identifier to avoid duplicates
            device_key = (lot_id, dev_num)
            
            if device_key not in device_set:
                device_set.add(device_key)
                
                # Attempt to find SKU for this device
                sku = find_sku_for_device(lot_id, dev_num, sku_lookup_table)
                
                devices.append({
                    "Lot_ID": lot_id,
                    "Dev#": dev_num,
                    "SN": "",  # Blank column for Serial Number
                    "SKU": sku if sku else ""  # Use found SKU or blank
                })
                
                if sku:
                    print(f"Successfully parsed: {filename} -> Lot_ID: {lot_id}, Dev#: {dev_num}, SKU: {sku}")
                else:
                    print(f"Successfully parsed: {filename} -> Lot_ID: {lot_id}, Dev#: {dev_num} (no SKU found)")
        except ValueError as e:
            print(f"Could not parse {filename}: {e}")

# Create DataFrame
df = pd.DataFrame(devices, columns=["Lot_ID", "Dev#", "SN", "SKU"])

# Create Excel file with renamed sheet
excel_path = os.path.join(destination_folder, "Devices.xlsx")
with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Devices", index=False)

print(f"SECTION 2 complete: Devices.xlsx created at {excel_path}")

# -------------------------
# SECTION 3 - Run the FIX script
# -------------------------
try:
    python_executable = r"C:\Users\crathod\Documents\Datasheet Automation\env\Scripts\python.exe"
    subprocess.run([python_executable, fix_script_path], check=True)
    print("SECTION 3 complete: FIX script executed.")
except subprocess.CalledProcessError as e:
    print(f"Error running FIX script: {e}")
except FileNotFoundError:
    print(f"FIX script not found at {fix_script_path}")

print("All sections complete.")
