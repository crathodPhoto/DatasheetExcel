import os
import pandas as pd

# Define the folder path where the .txt files are located
folder_path = r'C:\Users\crathod\Documents\Datasheet Automation\Script Output\Other'  # <-- Change this to your target folder

# Mapping of filename keywords to the phrase to search for
criteria = {
    "LIV_vs_Temp": "LIV Sweep vs Temperature",
    "SpecWidth": "Mode Spacing vs I &T",
    "WLT_SMSR": "SMSR vs I &T",
    "WLT_Wave": "Peak Wavelength vs I &T"
}

# List to store results
results = []

# Iterate over all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        file_path = os.path.join(folder_path, filename)

        # Determine which phrase to search for based on filename
        for key, phrase in criteria.items():
            if key in filename:
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                    
                    # Find occurrences of the phrase in the file
                    phrase_indices = [i for i, line in enumerate(lines) if phrase in line]
                    count = len(phrase_indices)

                    results.append({
                        'Filename': filename,
                        'Keyword': key,
                        'Phrase': phrase,
                        'Count': count
                    })

                    # If phrase occurs more than once, modify the file
                    if count > 1:
                        last_occurrence = phrase_indices[-1]  # Get the last occurrence index
                        modified_lines = lines[last_occurrence:]  # Keep only lines from this point onward

                        # Overwrite the original file with the modified content
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.writelines(modified_lines)

                        print(f"Modified {filename}: Retained lines from {last_occurrence + 1} onwards.")

                except Exception as e:
                    print(f"Error processing {filename}: {e}")
                
                # If a file matches one criterion, we assume it doesn't match others.
                break

# Create a DataFrame from the results
df = pd.DataFrame(results)

# Define the output Excel file path
output_excel = os.path.join(folder_path, "B.xlsx")

# Write the DataFrame to an Excel file
df.to_excel(output_excel, index=False)

print(f"Results have been written to {output_excel}")
print("Processing complete.")
