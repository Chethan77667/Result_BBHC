import streamlit as st
import pandas as pd
import io
import os
from typing import Dict, List, Tuple, Optional, Any
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# === PAGE CONFIGURATION ===
st.set_page_config(
    page_title="Semester Result Analysis",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stAlert {
        margin-top: 1rem;
    }
    .block-container {
        padding-top: 2rem;
    }
    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: white;
        text-align: center;
        padding: 10px 0;
        border-top: 1px solid #e6e6e6;
        z-index: 999;
    }
    </style>
    """, unsafe_allow_html=True)

# === FUNCTIONS ===

def extract_subject_dict(file) -> Dict[str, str]:
    """
    Extract subject codes and names from the Excel file.
    
    Args:
        file: Uploaded Excel file object
    
    Returns:
        Dictionary mapping subject codes to subject names
    """
    try:
        df_subjects = pd.read_excel(file, sheet_name='Sheet1', header=None)
        subjects_dict = {}
        for i in range(len(df_subjects)):
            row = df_subjects.iloc[i]
            subject_code = str(row[3]).strip()
            subject_name = str(row[4]).strip()
            if len(subject_code) == 10 and subject_code.isalnum():
                subjects_dict[subject_code.upper()] = subject_name
        return subjects_dict
    except Exception as e:
        logger.error(f"Error extracting subjects: {str(e)}")
        raise Exception("Failed to extract subjects from the file. Please check the file format.")

def extract_usns(df: pd.DataFrame) -> List[Tuple[str, int]]:
    """
    Extract USNs and their starting positions from the dataframe.
    
    Args:
        df: Input DataFrame containing result data
    
    Returns:
        List of tuples containing (USN, start_row_index)
    """
    student_blocks = []
    try:
        for i in range(len(df)):
            if str(df.iloc[i, 4]).strip().upper() == 'USN':
                for col in range(5, len(df.columns)):
                    val = df.iloc[i, col]
                    if isinstance(val, str) and val.strip().upper().startswith('U'):
                        student_blocks.append((val.strip().upper(), i))
                        break
        return student_blocks
    except Exception as e:
        logger.error(f"Error extracting USNs: {str(e)}")
        raise Exception("Failed to extract student USNs. Please check the file format.")

def get_result_for_usn_subject(block, input_usn, subject_code, subjects_dict):
    subject_code = subject_code.strip().upper()
    for i in range(len(block)):
        row = block.iloc[i]
        if subject_code in row.values:
            col_idx = list(row).index(subject_code)
            try:
                result_status = block.iloc[i + 10, col_idx]
                marks_scored = block.iloc[i + 7, col_idx]
                max_marks = block.iloc[i + 8, col_idx]
            except:
                result_status = ""
                marks_scored = ""
                max_marks = ""

            name = "Not Found"

            for i in range(len(block) - 1):
                if str(block.iloc[i, 4]).strip().upper() == 'USN':
                    val = str(block.iloc[i + 1, 4]).strip()
                    if val and not val.upper().startswith("U0") and len(val.split()) >= 1:
                        name = val
                    break

            return {
                "Name": name.upper(),
                "USN": input_usn,
                f"{subject_code} - {subjects_dict[subject_code]}": result_status,
                f"{subject_code}_marks": marks_scored,
                f"{subject_code}_max": max_marks
            }

    return {
        "Name": "Not Found",
        "USN": input_usn,
        f"{subject_code} - {subjects_dict[subject_code]}": "N/A",
        f"{subject_code}_marks": "N/A",
        f"{subject_code}_max": "N/A"
    }

def get_gpa_result_total_percent(block, input_usn):
    sgpa = cgpa = result_status = "Not Found"
    for i in range(len(block)):
        row = block.iloc[i]
        label = str(row[4]).strip().upper()
        if label == 'SGPA':
            sgpa = row[5]
        elif label == 'CGPA':
            cgpa = row[5]
        elif label.startswith("RESULT:"):
            result_status = label.split(":")[-1].strip().upper()
    return {
        "USN": input_usn,
        "SGPA": sgpa,
        "CGPA": cgpa,
        "Result": result_status,
    }

def get_total_and_percentage(block):
    total_marks = max_total = percentage = None
    for i in range(len(block)):
        row = block.iloc[i]
        if "Total" in row.values:
            try:
                col_idx = list(row).index("Total")
                total_marks = block.iloc[i + 7, col_idx]
                max_total = block.iloc[i + 8, col_idx]
                if isinstance(total_marks, (int, float)) and isinstance(max_total, (int, float)) and max_total != 0:
                    percentage = round((total_marks / max_total) * 100, 2)
            except Exception as e:
                print(f"Error extracting totals: {e}")
            break
    return {
        "Total Marks": total_marks,
        "Max Marks": max_total,
        "Percentage": percentage
    }

def clean_to_int(val):
    if isinstance(val, float) and val.is_integer():
        return int(val)
    return val

@st.cache_data(show_spinner="📊 Processing result sheet...")
def process_result_file(uploaded_file, output_filename: str) -> pd.DataFrame:
    """
    Process the uploaded result file and generate analysis.
    
    Args:
        uploaded_file: Uploaded Excel file
        output_filename: Name for the output file
    
    Returns:
        DataFrame containing processed results
    """
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Sheet2', header=None)
        subjects_dict = extract_subject_dict(uploaded_file)
        student_blocks = extract_usns(df)

        if not student_blocks:
            raise ValueError("No student records found in the file.")

        full_data = []
        progress_text = st.empty()
        progress_bar = st.progress(0)

        for idx, (usn, start_idx) in enumerate(student_blocks, start=1):
            progress_text.text(f"Processing student {idx}/{len(student_blocks)}: {usn}")
            block = df.iloc[start_idx-5:start_idx + 12]
            student_row = {"USN": usn, "Name": "Not Found"}

            # Process subject results
            for code in subjects_dict:
                result = get_result_for_usn_subject(block, usn, code, subjects_dict)
                student_row[f"{code} - {subjects_dict[code]}"] = result.get(f"{code} - {subjects_dict[code]}", "")
                if student_row["Name"] == "Not Found" and result["Name"] != "Not Found":
                    student_row["Name"] = result["Name"]

            # Process GPA and totals
            gpa_data = get_gpa_result_total_percent(block, usn)
            total_data = get_total_and_percentage(block)
            
            student_row.update({
                "SGPA": gpa_data["SGPA"],
                "CGPA": gpa_data["CGPA"],
                "Result": gpa_data["Result"],
                "Total Marks": clean_to_int(total_data["Total Marks"]),
                "Max Marks": clean_to_int(total_data["Max Marks"]),
                "Percentage": total_data["Percentage"]
            })

            full_data.append(student_row)
            progress_bar.progress(idx / len(student_blocks))

        progress_text.empty()
        progress_bar.empty()

        final_df = pd.DataFrame(full_data)
        
        # First sort by percentage
        final_df = final_df.sort_values(by="Percentage", ascending=False)
        # Then add Sl. No starting from 1
        final_df.insert(0, "Sl. No", range(1, len(final_df) + 1))

        # Organize columns
        subject_columns = [f"{code} - {subjects_dict[code]}" for code in subjects_dict]
        final_columns = ["Sl. No", "Name", "USN"] + subject_columns + [
            "Result", "Total Marks", "Max Marks", "Percentage", "CGPA", "SGPA"
        ]
        final_df = final_df[final_columns]

        # Clean and format numeric columns
        final_df["Total Marks"] = pd.to_numeric(final_df["Total Marks"], errors="coerce").fillna(0).astype(int)
        final_df["Max Marks"] = pd.to_numeric(final_df["Max Marks"], errors="coerce").fillna(0).astype(int)
        final_df["Percentage"] = final_df["Percentage"].apply(lambda x: 0 if pd.isna(x) else round(x, 2) if isinstance(x, (float, int)) else 0.0)

        return final_df

    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        st.error(f"Error processing file: {str(e)}")
        return None

# === MAIN APP ===

def main():
    st.title("📚 Semester Result Analysis")
    st.markdown("Upload your result sheet to view and save results.")

    # Add footer
    st.markdown("""
    <div class="footer">
        <p style='font-size: 14px; color: #666; margin: 0;'>
            Designed and Developed by<br>
            <strong>Mr. Pranam B</strong><br>
            Assistant Professor<br>
            Department of Computer Applications
        </p>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.header("ℹ️ Instructions")
        st.markdown("""
        1. Upload an Excel file (.xlsx format)
        2. Enter a name for the output file
        3. View the processed results
        4. Search for individual student results
        5. Save the results
        """)
        
        st.markdown("---")
        st.markdown("### 📌 File Format Requirements")
        st.markdown("""
        - Sheet1: Subject details
        - Sheet2: Result data
        - Required columns: USN, Name, Subjects
        """)

    uploaded_file = st.file_uploader("📂 Upload the Excel file (.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        output_filename = st.text_input(
            "📝 Enter output file name (without extension):",
            value="",
            help="This will be used to save the processed results"
        )

        if output_filename:
            final_df = process_result_file(uploaded_file, output_filename)
            
            if final_df is not None:
                st.subheader("📋 Results")
                st.dataframe(final_df, hide_index=True)

                if st.button("💾 Save Result Excel"):
                    try:
                        save_path = os.path.join(os.getcwd(), f"{output_filename}.xlsx")
                        final_df.to_excel(save_path, index=False)
                        st.success(f"✅ File saved successfully at: {save_path}")
                    except Exception as e:
                        st.error(f"Error saving file: {str(e)}")

                # Individual student search
                st.markdown("---")
                search_usn = st.text_input(
                    "🔍 Search by USN:",
                    help="Enter a USN to view detailed individual result"
                ).strip().upper()

                if search_usn:
                    student_result = final_df[final_df["USN"] == search_usn]
                    
                    if not student_result.empty:
                        st.subheader(f"🎯 Individual Result Card")
                        student = student_result.iloc[0]
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown(f"**👤 Name:** {student['Name']}")
                            st.markdown(f"**🆔 USN:** {student['USN']}")
                            st.markdown(f"**📊 Total Marks:** {student['Total Marks']}/{student['Max Marks']}")
                        
                        with col2:
                            st.markdown(f"**🎯 Result:** {student['Result']}")
                            percentage = 0 if pd.isna(student['Percentage']) else student['Percentage']
                            st.markdown(f"**📈 Percentage:** {percentage}%")
                            cgpa = "" if pd.isna(student['CGPA']) else student['CGPA']
                            sgpa = "" if pd.isna(student['SGPA']) else student['SGPA']
                            st.markdown(f"**🎓 CGPA:** {cgpa}")
                            st.markdown(f"**📚 SGPA:** {sgpa}")
                        
                        st.markdown("### 📘 Subject-wise Performance")
                        subject_data = []
                        # Extract subjects dictionary from the uploaded file
                        subjects_dict = extract_subject_dict(uploaded_file)
                        
                        # Get the student's block from the original DataFrame
                        original_df = pd.read_excel(uploaded_file, sheet_name='Sheet2', header=None)
                        student_blocks = extract_usns(original_df)
                        student_block = None
                        for usn, start_idx in student_blocks:
                            if usn == student['USN']:
                                student_block = original_df.iloc[start_idx-5:start_idx + 12]
                                break
                        
                        if student_block is not None:
                            for col in final_df.columns:
                                if "-" in col and student[col] != "N/A":
                                    subject_code = col.split(" - ")[0]
                                    subject_name = col.split(" - ")[1]
                                    result = get_result_for_usn_subject(student_block, student['USN'], subject_code, subjects_dict)
                                    marks_scored = result.get(f"{subject_code}_marks", "N/A")
                                    max_marks = result.get(f"{subject_code}_max", "N/A")
                                    subject_data.append({
                                        "Subject": col,
                                        "Marks": f"{marks_scored}/{max_marks}",
                                        "Result": student[col]
                                    })
                        
                        if subject_data:
                            subject_df = pd.DataFrame(subject_data)
                            subject_df.index = subject_df.index + 1  # Start index from 1
                            st.table(subject_df)
                    else:
                        st.warning("🚫 No record found for this USN.")

if __name__ == "__main__":
    main()
