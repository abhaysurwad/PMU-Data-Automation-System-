import pandas as pd
import os
import re
from datetime import datetime

# =====================================================
# Dynamic Base Folder
# =====================================================

# EXE/PY jis folder me run hogi wahi base folder hoga
base_folder = os.getcwd()

# Input Folder
input_folder = os.path.join(base_folder, "PMU_Input")

# Output Folder
output_folder = os.path.join(base_folder, "PMU_Output")

# Output folder automatically create hoga
os.makedirs(output_folder, exist_ok=True)

# =====================================================
# Date Wise Output File
# =====================================================

today = datetime.now().strftime("%Y-%m-%d")

output_file = os.path.join(
    output_folder,
    f"PMU_Validity_Report_{today}.xlsx"
)

# =====================================================
# Empty List
# =====================================================

all_summary = []

# Serial Number
s_no = 1

# =====================================================
# Read All PMU Files
# =====================================================

for file in os.listdir(input_folder):

    if file.endswith(".log") or file.endswith(".txt"):

        file_path = os.path.join(input_folder, file)

        try:

            # Read File
            df = pd.read_csv(file_path)

            # Remove Extra Spaces
            df.columns = df.columns.str.strip()

            # Extract PMU ID
            # Example:
            # PmuStats_15001_15002.log
            match = re.search(r'PmuStats_(\d+)_(\d+)', file)

            if match:

                # Second PMU ID
                pmu_id = match.group(2)

                # Find Valid Count Column
                validity_column = [
                    col for col in df.columns
                    if "Valid Count" in col
                ][0]

                # Total Validity
                total_validity = df[validity_column].sum()

                # Validity Calculation
                validity = total_validity / 1440

                # Store Data
                all_summary.append({
                    "S.NO": s_no,
                    "PMU ID": pmu_id,
                    "Validity": round(validity, 2)
                })

                s_no += 1

        except Exception as e:

            print(f"Error in file {file}: {e}")

# =====================================================
# Create Final DataFrame
# =====================================================

final_report = pd.DataFrame(all_summary)

# =====================================================
# Save Excel File
# =====================================================

final_report.to_excel(output_file, index=False)

# =====================================================
# Success Message
# =====================================================

print("\n===================================")
print(" Excel File Saved Successfully ")
print("===================================")
print("Saved At:")
print(output_file)
print("===================================")