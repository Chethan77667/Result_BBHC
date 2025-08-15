import pandas as pd
import re

def parse_reexam_txt(txt_path):
    students = {}
    with open(txt_path, 'r', encoding='utf-8') as f:
        content = f.read()
    blocks = content.strip().split('-' * 40)
    for block in blocks:
        lines = [line.strip() for line in block.strip().split('\n') if line.strip()]
        if not lines:
            continue
        usn = None
        name = None
        before_result = after_result = None
        before_pct = after_pct = None
        before_total = after_total = None
        changed_subjects = {}
        for line in lines:
            if line.startswith('USN:'):
                usn = line.split('USN:')[1].strip()
            elif line.startswith('Name:'):
                name = line.split('Name:')[1].strip()
            elif line.startswith('Before Result:'):
                parts = line.split('\t')
                before_result = parts[0].split(':')[1].strip()
                after_result = parts[1].split(':')[1].strip()
            elif line.startswith('Before Total:'):
                parts = line.split('\t')
                before_total = parts[0].split(':')[1].strip()
                after_total = parts[1].split(':')[1].strip()
            elif line.startswith('Before Percentage:'):
                parts = line.split('\t')
                before_pct = float(parts[0].split(':')[1].strip())
                after_pct = float(parts[1].split(':')[1].strip())
            elif line.startswith('Changed Subjects:'):
                continue
            elif ':' in line and 'Before:' in line and 'Current:' in line:
                # Subject line
                subject, rest = line.split(': Before:')
                before, after = rest.split('Current:')
                # Only keep Pass/Fail part
                after_val = after.strip()
                if '-' in after_val:
                    after_val = after_val.split('-')[1].strip()
                changed_subjects[subject.strip()] = after_val
        if usn:
            students[usn] = {
                'Result': after_result,
                'Total Marks': after_total,
                'Percentage': after_pct,
                'ChangedSubjects': changed_subjects
            }
    return students

def update_master_excel(master_path, txt_path, output_path):
    updates = parse_reexam_txt(txt_path)
    df = pd.read_excel(master_path)
    for idx, row in df.iterrows():
        usn = str(row['USN']).strip().upper()
        if usn in updates:
            # Update Result, Total Marks, and Percentage
            if updates[usn]['Result']:
                df.at[idx, 'Result'] = str(updates[usn]['Result']).upper()
            if updates[usn]['Total Marks']:
                df.at[idx, 'Total Marks'] = updates[usn]['Total Marks']
            if updates[usn]['Percentage'] is not None:
                df.at[idx, 'Percentage'] = updates[usn]['Percentage']
            # Update changed subject columns to only Pass/Fail
            for subj, new_val in updates[usn]['ChangedSubjects'].items():
                if subj in df.columns:
                    df.at[idx, subj] = new_val
    # Fill all NaN with 'N/A' before saving
    df = df.fillna('N/A')
    df.to_excel(output_path, index=False)
    print(f'Updated master file saved as {output_path}')

# Example usage:
update_master_excel("F:\Cursor Projects\Internals Lookup\SEM RESULTS\II BCOM\I SEM.xlsx", 're_exam_comparison.txt', 'updated_master.xlsx') 