import pandas as pd

def extract_subject_dict(file):
    df_subjects = pd.read_excel(file, sheet_name='Sheet1', header=None)
    subjects_dict = {}
    for i in range(len(df_subjects)):
        row = df_subjects.iloc[i]
        subject_code = str(row[3]).strip()
        subject_name = str(row[4]).strip()
        if len(subject_code) == 10 and subject_code.isalnum():
            subjects_dict[subject_code.upper()] = subject_name
    return subjects_dict

def extract_usns(df):
    student_blocks = []
    for i in range(len(df)):
        if str(df.iloc[i, 4]).strip().upper() == 'USN':
            for col in range(5, len(df.columns)):
                val = df.iloc[i, col]
                if isinstance(val, str) and val.strip().upper().startswith('U'):
                    student_blocks.append((val.strip().upper(), i))
                    break
    return student_blocks

def get_result_for_usn_subject(block, input_usn, subject_code, subjects_dict):
    subject_code = subject_code.strip().upper()
    for i in range(len(block)):
        row = block.iloc[i]
        if subject_code in row.values:
            col_idx = list(row).index(subject_code)
            try:
                result_status = block.iloc[i + 10, col_idx]
                marks_scored = block.iloc[i + 7, col_idx]
            except:
                result_status = ""
                marks_scored = ""
            return f"{marks_scored}-{result_status}"
    return "N/A"

def process_result_file(input_file, output_file):
    df = pd.read_excel(input_file, sheet_name='Sheet2', header=None)
    subjects_dict = extract_subject_dict(input_file)
    student_blocks = extract_usns(df)

    if not student_blocks:
        raise ValueError("No student records found in the file.")

    full_data = []
    for usn, start_idx in student_blocks:
        block = df.iloc[start_idx-5:start_idx + 12]
        student_row = {"USN": usn, "Name": "Not Found"}

        # Process subject results
        for code in subjects_dict:
            entry = get_result_for_usn_subject(block, usn, code, subjects_dict)
            student_row[f"{code} - {subjects_dict[code]}"] = entry
            # Try to get name
            if student_row["Name"] == "Not Found":
                for i in range(len(block) - 1):
                    if str(block.iloc[i, 4]).strip().upper() == 'USN':
                        val = str(block.iloc[i + 1, 4]).strip()
                        if val and not val.upper().startswith("U0") and len(val.split()) >= 1:
                            student_row["Name"] = val
                        break

        full_data.append(student_row)

    final_df = pd.DataFrame(full_data)
    final_df.insert(0, "Sl. No", range(1, len(final_df) + 1))
    # Reorder columns
    subject_columns = [f"{code} - {subjects_dict[code]}" for code in subjects_dict]
    final_columns = ["Sl. No", "Name", "USN"] + subject_columns
    final_df = final_df[final_columns]
    final_df.to_excel(output_file, index=False)
    print(f"✅ Output saved to {output_file}")

# Usage example:
process_result_file("I_BCA_I_sem_re_exam_ledger_Dec-Jan_2025.xlsx", 'output_file3.xlsx')