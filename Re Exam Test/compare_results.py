import pandas as pd
import re

def extract_student_data(file):
    df = pd.read_excel(file)
    df = df.fillna('N/A')
    students = {}
    usn_pattern = re.compile(r'^U05[A-Z]{2}23', re.IGNORECASE)
    for _, row in df.iterrows():
        usn = str(row['USN']).strip().upper()
        if usn_pattern.match(usn):
            students[usn] = row.to_dict()
    return students, list(df.columns)

def parse_mark_result(cell):
    if isinstance(cell, str) and ('-Pass' in cell or '-Fail' in cell):
        try:
            mark, result = cell.split('-')
            mark = int(mark) if mark.isdigit() else None
            return mark, result.strip()
        except:
            return None, None
    return None, None

def get_subject_columns(columns):
    # Exclude Sl. No, Name, USN
    return [col for col in columns if col not in ['Sl. No', 'Name', 'USN']]

def calculate_total_and_percentage(student, subject_cols, max_total):
    total = 0
    for col in subject_cols:
        mark, result = parse_mark_result(student.get(col, 'N/A'))
        if mark is not None:
            total += mark
    percentage = (total / max_total) * 100 if max_total > 0 else 0
    return total, round(percentage, 2)

def get_overall_result(student, subject_cols):
    for col in subject_cols:
        cell = student.get(col, 'N/A')
        result = None
        if isinstance(cell, str):
            if '-' in cell:
                parts = cell.split('-')
                result = parts[-1].strip().lower()
            else:
                result = cell.strip().lower()
        else:
            result = str(cell).strip().lower()
        if result == 'n/a':
            continue  # Ignore N/A subjects
        if result != 'pass':
            return 'Fail'
    return 'Pass'

def compare_results(regular_file, reexam_file, output_txt):
    reg_students, columns = extract_student_data(regular_file)
    re_students, _ = extract_student_data(reexam_file)
    subject_cols = get_subject_columns(columns)

    # Prompt user for max total
    while True:
        try:
            max_total = float(input(f"Enter the maximum total marks for all subjects (e.g., 600): "))
            if max_total > 0:
                break
            else:
                print("Max total must be positive.")
        except ValueError:
            print("Please enter a valid number.")

    written_usns = set()
    with open(output_txt, 'w', encoding='utf-8') as f:
        for usn, re_stud in re_students.items():
            if usn not in reg_students:
                continue
            reg_stud = reg_students[usn]
            changed_subjects = []
            merged_stud = reg_stud.copy()
            for col in subject_cols:
                # If subject present in re-exam, use re-exam marks (even if fail/pass)
                if col in re_stud and re_stud[col] != 'N/A':
                    if reg_stud.get(col, 'N/A') != re_stud[col]:
                        changed_subjects.append((col, reg_stud.get(col, 'N/A'), re_stud[col]))
                    merged_stud[col] = re_stud[col]
            before_total, before_pct = calculate_total_and_percentage(reg_stud, subject_cols, max_total)
            after_total, after_pct = calculate_total_and_percentage(merged_stud, subject_cols, max_total)
            before_result = get_overall_result(reg_stud, subject_cols)
            after_result = get_overall_result(merged_stud, subject_cols)
            name = reg_stud.get('Name', '')
            # Debug: print all subject results for a specific USN after merging
            if usn == "U05BB23M0041":
                print(f"Debug: Subject results for {usn} after merging:")
                for col in subject_cols:
                    print(f"  {col}: {merged_stud.get(col, 'N/A')}")
            f.write(f'USN: {usn}\nName: {name}\n')
            f.write(f'Before Result: {before_result}\tCurrent Result: {after_result}\n')
            f.write(f'Before Total: {before_total}\tCurrent Total: {after_total}\n')
            f.write(f'Before Percentage: {before_pct}\tCurrent Percentage: {after_pct}\n')
            if changed_subjects:
                f.write('Changed Subjects:\n')
                for col, before, after in changed_subjects:
                    f.write(f'  {col}: Before: {before}\tCurrent: {after}\n')
            else:
                f.write('No subject marks changed.\n')
            f.write('-'*40 + '\n')
            written_usns.add(usn)
    # Print USNs and names of students in re-exam file but not written to text file
    skipped = []
    for usn, re_stud in re_students.items():
        if usn not in written_usns:
            name = re_stud.get('Name', '')
            skipped.append((usn, name))
    if skipped:
        print("USNs and names in re-exam file but not written to text file:")
        for usn, name in skipped:
            print(f"  {usn}: {name}")
    print(f'Results written to {output_txt}')

# Example usage:
compare_results('output_file1.xlsx', 'output_file2.xlsx', 're_exam_comparison.txt') 