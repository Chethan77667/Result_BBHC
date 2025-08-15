import pandas as pd
import numpy as np

# Load the Excel file
file_path = "F:\\Result Analysis\\I Sem BCA Dec 2024-Jan 2025 updated one new.xlsx"
df = pd.read_excel(file_path, sheet_name='Sheet2', header=None)

subject_code = input("Enter the subject code (e.g., BCACACS101): ").strip()
results = []
usn = None
name = None

# Iterate through the rows
for i in range(len(df)):
    row = df.iloc[i]

    # Update USN if found
    if str(row[4]).strip().upper() == 'USN':
        usn = row[5]

    # Check if subject code is in the row
    if subject_code in row.values:
        col_idx = list(row).index(subject_code)

        # Try to extract marks from next row
        try:
            mark_entry = df.iloc[i + 1, col_idx]
            if isinstance(mark_entry, str) and '+' in mark_entry:
                parts = mark_entry.replace('-', '0').split('+')
                parts = [float(p.strip()) for p in parts if p.strip().replace('.', '', 1).isdigit()]
                total_marks = sum(parts)
            else:
                total_marks = np.nan
        except:
            total_marks = np.nan

    # Once CGPA is reached, backtrack for name and finalize record
    if str(row[4]).strip().upper() == 'CGPA':
        for j in range(i - 1, max(i - 10, 0), -1):
            candidate_row = df.iloc[j]
            for col in range(4, 2, -1):
                val = candidate_row[col]
                if isinstance(val, str) and val.strip():
                    val = val.strip()
                    if not val.upper().startswith("USN") and val.upper() not in ["SGPA", "RESULT", "CLASS", "TERM GRADE", "M.C.NO"]:
                        name = val
                        break
            if name:
                break

        try:
            if total_marks is None or pd.isna(total_marks):
                total_marks = "Not found"
        except NameError:
            total_marks = "Not found"

        #print(f"{name} || {usn} || {total_marks}")
        if name and usn and not pd.isna(total_marks):
            results.append((usn, name.upper(), total_marks))

            # Reset for next student
            name = None
            usn = None
            total_marks = None

# Create a DataFrame and sort by Total Marks descending
df_results = pd.DataFrame(results, columns=["USN", "Name", "Total Marks"])

# Separate rows with numeric total marks for sorting
df_numeric = df_results[pd.to_numeric(df_results["Total Marks"], errors="coerce").notna()]
df_non_numeric = df_results[pd.to_numeric(df_results["Total Marks"], errors="coerce").isna()]

# Convert to float for sorting
df_numeric["Total Marks"] = df_numeric["Total Marks"].astype(float)
df_numeric = df_numeric.sort_values(by="Total Marks", ascending=False)

# Combine them back (numeric sorted first, then non-numeric like "Not found")
df_results = pd.concat([df_numeric, df_non_numeric], ignore_index=True)

# Save to Excel
output_file = input("Enter the name of the Excel file to save (without .xlsx): ")
df_results.to_excel(f"{output_file}.xlsx", index=False, engine='openpyxl')

print(f"✅ Saved to {output_file}.xlsx (sorted by Total Marks)")
