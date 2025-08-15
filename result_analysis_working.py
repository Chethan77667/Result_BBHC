import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Sem Result Analysis", layout="wide")
st.title("📊 Semester Result Analysis Tool")

uploaded_file = st.file_uploader("📁 Upload the Excel file", type=["xlsx"])
output_filename = st.text_input("💾 Enter output file name (without extension):")

# ======= Extract Functions =======
@st.cache_data
def extract_subject_dict(file) -> dict:
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
    usns = set()
    for i in range(len(df)):
        row = df.iloc[i]
        if str(row[4]).strip().upper() == 'USN':
            for col in range(5, len(row)):
                val = row[col]
                if isinstance(val, str) and val.strip().upper().startswith('U'):
                    usns.add(val.strip().upper())
    return sorted(usns)

def get_result_for_usn_subject(df, input_usn, subject_code, subjects_dict):
    input_usn = input_usn.strip().upper()
    subject_code = subject_code.strip().upper()
    for i in range(len(df)):
        row = df.iloc[i]
        if subject_code in row.values:
            col_idx = list(row).index(subject_code)
            try:
                result_status = df.iloc[i + 10, col_idx]
            except:
                result_status = ""

            usn_row_index = None
            for j in range(i, min(i + 10, len(df))):
                if str(df.iloc[j, 4]).strip().upper() == 'USN':
                    usn_row_index = j
                    break

            if usn_row_index is not None:
                usn_row = df.iloc[usn_row_index]
                for c in range(len(usn_row)):
                    if str(usn_row[c]).strip().upper() == input_usn:
                        name = "Not Found"
                        for k in range(usn_row_index, min(usn_row_index + 20, len(df))):
                            if str(df.iloc[k, 4]).strip().upper() == 'CGPA':
                                for m in range(k - 1, max(k - 10, 0), -1):
                                    name_row = df.iloc[m]
                                    for n in range(4, 2, -1):
                                        val = name_row[n]
                                        if isinstance(val, str) and val.strip():
                                            val = val.strip()
                                            if not val.upper().startswith("USN") and val.upper() not in ["SGPA", "RESULT", "CLASS", "TERM GRADE", "M.C.NO"]:
                                                name = val
                                                break
                                    if name != "Not Found":
                                        break
                                break
                        return {
                            "Name": name.upper(),
                            "USN": input_usn,
                            f"{subject_code} - {subjects_dict[subject_code]}": result_status
                        }
    return {
        "Name": "Not Found",
        "USN": input_usn,
        f"{subject_code} - {subjects_dict[subject_code]}": "N/A"
    }

def get_gpa_result_total_percent(df, input_usn):
    input_usn = input_usn.strip().upper()
    sgpa = cgpa = result_status = total_marks = max_total = percentage = "Not Found"

    for i in range(len(df)):
        row = df.iloc[i]
        if str(row[4]).strip().upper() == 'USN':
            for col in range(5, len(row)):
                val = row[col]
                if isinstance(val, str) and val.strip().upper() == input_usn:
                    for j in range(i + 1, min(i + 40, len(df))):
                        label = str(df.iloc[j, 4]).strip().upper()
                        label2 = str(df.iloc[j, 6]).strip().upper()
                        if label == 'SGPA':
                            sgpa = df.iloc[j, 5]
                        elif label == 'CGPA':
                            cgpa = df.iloc[j, 5]
                        elif label.startswith("RESULT:"):
                            result_status = label.split(":")[-1].strip().upper()
                        if label2 == 'TOTAL':
                            total_marks = df.iloc[j, 16]
                        elif label2 == 'MAX. TOTAL':
                            max_total = df.iloc[j, 16]
                        if isinstance(total_marks, (int, float)) and isinstance(max_total, (int, float)) and max_total != 0:
                            percentage = round((total_marks / max_total) * 100, 2)
                        if all(x != "Not Found" for x in [sgpa, cgpa, result_status, total_marks, max_total, percentage]):
                            return {
                                "USN": input_usn,
                                "SGPA": sgpa,
                                "CGPA": cgpa,
                                "Result": result_status,
                                "Total Marks": total_marks,
                                "Percentage": percentage
                            }
                    break
    return {
        "USN": input_usn,
        "SGPA": sgpa,
        "CGPA": cgpa,
        "Result": result_status,
        "Total Marks": total_marks,
        "Percentage": percentage
    }

# ======= MAIN LOGIC =======
if uploaded_file and output_filename:
    if st.button("🔄 Reset"):
        st.session_state.clear()
        st.experimental_rerun()
        
    if 'final_df' not in st.session_state:
        df = pd.read_excel(uploaded_file, sheet_name='Sheet2', header=None)
        subjects_dict = extract_subject_dict(uploaded_file)
        usn_list = extract_usns(df)

        full_data = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        total = len(usn_list)

        for idx, usn in enumerate(usn_list, start=1):
            student_row = {"USN": usn, "Name": "Not Found"}
            for code in subjects_dict:
                result = get_result_for_usn_subject(df, usn, code, subjects_dict)
                student_row[f"{code} - {subjects_dict[code]}"] = result.get(f"{code} - {subjects_dict[code]}", "")
                if student_row["Name"] == "Not Found" and result["Name"] != "Not Found":
                    student_row["Name"] = result["Name"]

            gpa_data = get_gpa_result_total_percent(df, usn)
            student_row["SGPA"] = gpa_data["SGPA"]
            student_row["CGPA"] = gpa_data["CGPA"]
            student_row["Result"] = gpa_data["Result"]
            student_row["Total Marks"] = gpa_data["Total Marks"]
            student_row["Percentage"] = gpa_data["Percentage"]

            full_data.append(student_row)

            percent_complete = int((idx / total) * 100)
            progress_bar.progress(percent_complete / 100)
            status_text.text(f"Processing {idx}/{total} students... ({percent_complete}%)")

        final_df = pd.DataFrame(full_data)
        final_df.insert(0, "Sl. No", range(1, len(final_df) + 1))
        subject_columns = [f"{code} - {subjects_dict[code]}" for code in subjects_dict]
        final_columns = ["Sl. No", "Name", "USN"] + subject_columns + ["Result", "Total Marks", "Percentage", "CGPA", "SGPA"]
        final_df = final_df[final_columns]

        final_df = final_df.sort_values(by="Percentage", ascending=False, na_position='last')

        st.session_state['final_df'] = final_df
        st.session_state['subjects_dict'] = subjects_dict

    st.success(f"✅ Processed {len(st.session_state['final_df'])} student records.")

    tab1, tab2 = st.tabs(["📋 View Results", "📥 Download File"])
    with tab1:
        st.dataframe(st.session_state['final_df'], use_container_width=True)

    with tab2:
        buffer = BytesIO()
        st.session_state['final_df'].to_excel(buffer, index=False, engine='openpyxl')
        st.download_button(
            label="Download Excel File",
            data=buffer.getvalue(),
            file_name=f"{output_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )