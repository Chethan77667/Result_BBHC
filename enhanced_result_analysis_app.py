import streamlit as st
import pandas as pd
import io
import os
import json
import glob
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# === PAGE CONFIGURATION ===
st.set_page_config(
    page_title="Enhanced Result Analysis",
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
    .file-info {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)

# === CONSTANTS AND PATHS ===
UPLOAD_DIR = "uploaded_files"
PROCESSED_DIR = "processed_results"
METADATA_FILE = "file_metadata.json"

# Create directories if they don't exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

# === UTILITY FUNCTIONS ===

def load_metadata():
    """Load file metadata from JSON file."""
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_metadata(metadata):
    """Save file metadata to JSON file."""
    with open(METADATA_FILE, 'w') as f:
        json.dump(metadata, f, indent=2)

def auto_save_file(uploaded_file, file_type="upload"):
    """Auto-save uploaded file with timestamp."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{timestamp}_{uploaded_file.name}"
    
    if file_type == "upload":
        save_path = os.path.join(UPLOAD_DIR, filename)
    else:
        save_path = os.path.join(PROCESSED_DIR, filename)
    
    with open(save_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    
    return filename, save_path

def get_semester_year_from_filename(filename):
    """Extract semester and year information from filename."""
    filename_lower = filename.lower()
    
    # Extract semester
    if 'i sem' in filename_lower or '1st sem' in filename_lower:
        semester = "I"
    elif 'ii sem' in filename_lower or '2nd sem' in filename_lower:
        semester = "II"
    elif 'iii sem' in filename_lower or '3rd sem' in filename_lower:
        semester = "III"
    elif 'iv sem' in filename_lower or '4th sem' in filename_lower:
        semester = "IV"
    elif 'v sem' in filename_lower or '5th sem' in filename_lower:
        semester = "V"
    elif 'vi sem' in filename_lower or '6th sem' in filename_lower:
        semester = "VI"
    else:
        semester = "Unknown"
    
    # Extract year
    if '2023' in filename:
        year = "2023-24"
    elif '2024' in filename:
        year = "2024-25"
    elif '2025' in filename:
        year = "2025-26"
    else:
        year = "Unknown"
    
    return semester, year

def search_files(query, file_type="all"):
    """Search for files based on query."""
    metadata = load_metadata()
    results = []
    
    # Search in metadata (processed files)
    for file_info in metadata.values():
        if file_type != "all" and file_info.get("type") != file_type:
            continue
            
        filename = file_info.get("filename", "")
        original_file = file_info.get("original_file", "")
        
        # Search in both filename and original filename
        if (query.lower() in filename.lower() or 
            query.lower() in original_file.lower()):
            results.append(file_info)
    
    # Search for existing files in current directory
    if file_type in ["all", "existing"]:
        current_dir_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
        for filename in current_dir_files:
            if query.lower() in filename.lower():
                # Create a file info entry for existing files
                file_info = {
                    "filename": filename,
                    "original_file": filename,
                    "type": "existing",
                    "upload_date": "Existing file",
                    "file_path": os.path.join('.', filename),
                    "student_count": "Unknown"
                }
                results.append(file_info)
    
    # Search for uploaded files in uploaded_files directory
    if file_type in ["all", "uploaded"] and os.path.exists(UPLOAD_DIR):
        uploaded_files = [f for f in os.listdir(UPLOAD_DIR) if f.endswith('.xlsx')]
        for filename in uploaded_files:
            if query.lower() in filename.lower():
                # Create a file info entry for uploaded files
                file_info = {
                    "filename": filename,
                    "original_file": filename,
                    "type": "uploaded",
                    "upload_date": "Uploaded file",
                    "file_path": os.path.join(UPLOAD_DIR, filename),
                    "student_count": "Unknown"
                }
                results.append(file_info)
    
    return results

def update_result_with_reexam(result_file_path, reexam_file):
    """Update result file with re-exam data."""
    try:
        # Load original result file
        result_df = pd.read_excel(result_file_path)
        
        # Load re-exam file from the uploaded file object
        reexam_df = pd.read_excel(reexam_file)
        
        # Extract USNs from re-exam file
        reexam_usns = set()
        for col in reexam_df.columns:
            if 'USN' in str(col).upper():
                usns = reexam_df[col].dropna().astype(str)
                reexam_usns.update([usn.strip().upper() for usn in usns if usn.strip().startswith('U')])
        
        # Update results for students in re-exam
        updated_count = 0
        for idx, row in result_df.iterrows():
            usn = str(row['USN']).strip().upper()
            if usn in reexam_usns:
                # Find corresponding row in re-exam file
                for _, reexam_row in reexam_df.iterrows():
                    for col in reexam_df.columns:
                        if 'USN' in str(col).upper():
                            reexam_usn = str(reexam_row[col]).strip().upper()
                            if reexam_usn == usn:
                                # Update subject results
                                for subject_col in result_df.columns:
                                    if subject_col not in ['Sl. No', 'Name', 'USN', 'Result', 'Total Marks', 'Max Marks', 'Percentage', 'CGPA', 'SGPA']:
                                        # Find corresponding subject in re-exam file
                                        for reexam_col in reexam_df.columns:
                                            if subject_col.split(' - ')[0] in str(reexam_col):
                                                new_result = reexam_row[reexam_col]
                                                if pd.notna(new_result):
                                                    result_df.at[idx, subject_col] = new_result
                                                break
                                
                                # Recalculate totals and percentages
                                total_marks = 0
                                max_marks = 0
                                for subject_col in result_df.columns:
                                    if subject_col not in ['Sl. No', 'Name', 'USN', 'Result', 'Total Marks', 'Max Marks', 'Percentage', 'CGPA', 'SGPA']:
                                        cell_value = result_df.at[idx, subject_col]
                                        if isinstance(cell_value, str) and '-' in cell_value:
                                            try:
                                                marks = int(cell_value.split('-')[0])
                                                total_marks += marks
                                                max_marks += 100  # Assuming max marks per subject is 100
                                            except:
                                                pass
                                
                                if max_marks > 0:
                                    percentage = round((total_marks / max_marks) * 100, 2)
                                    result_df.at[idx, 'Total Marks'] = total_marks
                                    result_df.at[idx, 'Max Marks'] = max_marks
                                    result_df.at[idx, 'Percentage'] = percentage
                                
                                updated_count += 1
                                break
                
        return result_df, updated_count
        
    except Exception as e:
        st.error(f"Error updating result file: {str(e)}")
        return None, 0

# === EXISTING FUNCTIONS ===

def extract_subject_dict(file, sheet_name='Sheet1') -> Dict[str, str]:
    """Extract subject codes and names from the Excel file."""
    try:
        df_subjects = pd.read_excel(file, sheet_name=sheet_name, header=None)
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
    """Extract USNs and their starting positions from the dataframe."""
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
def process_result_file(uploaded_file, output_filename: str, subject_sheet='Sheet1', result_sheet='Sheet2') -> pd.DataFrame:
    """Process the uploaded result file and generate analysis."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=result_sheet, header=None)
        subjects_dict = extract_subject_dict(uploaded_file, subject_sheet)
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
    st.title("📚 Enhanced Semester Result Analysis")
    
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

    # Navigation tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Process Results", "🔍 Search Files", "🔄 Re-exam Update", "ℹ️ Instructions"])

    with tab1:
        st.header("📊 Process New Results")
        
        uploaded_file = st.file_uploader("📂 Upload the Excel file (.xlsx)", type=["xlsx"], key="upload_tab1")
        
        if uploaded_file:
            # Auto-save uploaded file
            saved_filename, saved_path = auto_save_file(uploaded_file, "upload")
            
            # Extract semester and year info
            semester, year = get_semester_year_from_filename(uploaded_file.name)
            
            # Get available sheets from the Excel file
            try:
                excel_file = pd.ExcelFile(uploaded_file)
                available_sheets = excel_file.sheet_names
                st.info(f"📋 Available worksheets: {', '.join(available_sheets)}")
                
                # Sheet selection
                col1, col2 = st.columns(2)
                with col1:
                    subject_sheet = st.selectbox(
                        "📚 Select Subject Details Sheet:",
                        available_sheets,
                        index=0 if 'Sheet1' in available_sheets else 0,
                        help="Sheet containing subject codes and names"
                    )
                with col2:
                    result_sheet = st.selectbox(
                        "📊 Select Result Data Sheet:",
                        available_sheets,
                        index=1 if len(available_sheets) > 1 and 'Sheet2' in available_sheets else 0,
                        help="Sheet containing student results"
                    )
            except Exception as e:
                st.error(f"Error reading Excel file: {str(e)}")
                return
            
            output_filename = st.text_input(
                "📝 Enter output file name (without extension):",
                value=f"{semester}_Sem_{year}_Results",
                help="This will be used to save the processed results"
            )

            if output_filename:
                final_df = process_result_file(uploaded_file, output_filename, subject_sheet, result_sheet)
                
                if final_df is not None:
                    # Auto-save processed results
                    processed_filename = f"{output_filename}.xlsx"
                    processed_path = os.path.join(PROCESSED_DIR, processed_filename)
                    final_df.to_excel(processed_path, index=False)
                    
                    # Save metadata
                    metadata = load_metadata()
                    file_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    metadata[file_id] = {
                        "filename": processed_filename,
                        "original_file": saved_filename,
                        "type": "processed",
                        "semester": semester,
                        "year": year,
                        "upload_date": datetime.now().isoformat(),
                        "student_count": len(final_df),
                        "file_path": processed_path
                    }
                    save_metadata(metadata)
                    
                    st.success(f"✅ File auto-saved as: {processed_filename}")
                    
                    st.subheader("📋 Results")
                    # Display table header
                    cols = st.columns([1, 2, 2, 2, 2, 2, 1])
                    headers = ["Sl. No", "Name", "USN", "Result", "Total Marks", "Percentage", "View"]
                    for col, header in zip(cols, headers):
                        col.markdown(f"**{header}**")
                    
                    # Display each row with a button
                    for idx, row in final_df.iterrows():
                        cols = st.columns([1, 2, 2, 2, 2, 2, 1])
                        cols[0].write(row["Sl. No"])
                        cols[1].write(row["Name"])
                        cols[2].write(row["USN"])
                        cols[3].write(row["Result"])
                        cols[4].write(f"{row['Total Marks']}/{row['Max Marks']}")
                        cols[5].write(row["Percentage"])
                        if cols[6].button("View", key=f"view_{row['USN']}_{idx}"):
                            st.session_state.selected_usn = row['USN']
                            st.session_state.page = 'student'
                            st.experimental_rerun()
                    
                    # Download button
                    buffer = io.BytesIO()
                    final_df.to_excel(buffer, index=False)
                    st.download_button(
                        label="💾 Download Processed Results",
                        data=buffer.getvalue(),
                        file_name=processed_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_processed_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    )

    with tab2:
        st.header("🔍 Search Saved Files")
        
        # Debug: Show available files
        if st.checkbox("🔧 Debug: Show all available files"):
            st.subheader("📂 Files in uploaded_files directory:")
            if os.path.exists(UPLOAD_DIR):
                uploaded_files = [f for f in os.listdir(UPLOAD_DIR) if f.endswith('.xlsx')]
                if uploaded_files:
                    for file in uploaded_files:
                        st.write(f"• {file}")
                else:
                    st.write("No Excel files found in uploaded_files directory")
            else:
                st.write("uploaded_files directory does not exist")
            
            st.subheader("📂 Files in current directory:")
            current_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
            if current_files:
                for file in current_files:
                    st.write(f"• {file}")
            else:
                st.write("No Excel files found in current directory")
        
        # Search functionality
        search_query = st.text_input("🔍 Search files by name:", placeholder="Enter filename, semester, or year...")
        file_type_filter = st.selectbox("📁 File type:", ["All", "Uploaded", "Processed", "Existing"])
        
        if search_query:
            file_type = "all"
            if file_type_filter == "Uploaded":
                file_type = "upload"
            elif file_type_filter == "Processed":
                file_type = "processed"
            elif file_type_filter == "Existing":
                file_type = "existing"
            
            search_results = search_files(search_query, file_type)
            
            if search_results:
                st.success(f"Found {len(search_results)} file(s)")
                
                for file_info in search_results:
                    with st.expander(f"📄 {file_info['filename']}"):
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            st.markdown(f"**📅 Upload Date:** {file_info.get('upload_date', 'Unknown')}")
                            st.markdown(f"**📊 Type:** {file_info.get('type', 'Unknown')}")
                            if file_info.get('original_file'):
                                st.markdown(f"**📁 Original File:** {file_info['original_file']}")
                            if file_info.get('semester') and file_info.get('year'):
                                st.markdown(f"**🎓 Semester:** {file_info['semester']} ({file_info['year']})")
                            if file_info.get('student_count'):
                                st.markdown(f"**👥 Students:** {file_info['student_count']}")
                            if file_info.get('type') == 'uploaded':
                                st.markdown(f"**📂 Location:** {UPLOAD_DIR}/")
                            elif file_info.get('type') == 'existing':
                                st.markdown(f"**📂 Location:** Current Directory")
                        
                        with col2:
                             # Download button
                             if os.path.exists(file_info.get('file_path', '')):
                                 with open(file_info['file_path'], 'rb') as f:
                                     st.download_button(
                                         label="📥 Download",
                                         data=f.read(),
                                         file_name=file_info['filename'],
                                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                         key=f"download_search_{file_info['filename']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                                     )
                             
                             # Print option
                             if st.button("🖨️ Print", key=f"print_{file_info['filename']}"):
                                 st.info("File opened in new tab for printing")
            else:
                st.warning("No files found matching your search criteria.")
        else:
            st.info("Enter a search query to find saved files.")

    with tab3:
        st.header("🔄 Re-exam Update")
        

        
        st.markdown("""
        **Instructions:**
        1. Upload the re-exam ledger file
        2. Select the original result file to update
        3. The system will automatically update the results based on USN matching
        4. Totals and percentages will be recalculated
        5. **The original file will be overwritten with the updated data**
        6. **No copy files are created - the original file is modified in place**
        """)
        
        # File upload
        reexam_file = st.file_uploader("📂 Upload Re-exam Ledger (.xlsx)", type=["xlsx"], key="reexam_file")
        
        # Check if we have the file ready for processing
        if reexam_file:
            # Process re-exam file directly without auto-saving
            reexam_filename = reexam_file.name
            
            # Get available files from uploaded_files directory only
            uploaded_files = []
            
            # Get uploaded files from uploaded_files directory
            if os.path.exists(UPLOAD_DIR):
                uploaded_file_list = [f for f in os.listdir(UPLOAD_DIR) if f.endswith('.xlsx')]
                for filename in uploaded_file_list:
                    file_info = {
                        "filename": filename,
                        "original_file": filename,
                        "type": "uploaded",
                        "upload_date": "Uploaded file",
                        "file_path": os.path.join(UPLOAD_DIR, filename),
                        "student_count": "Unknown"
                    }
                    uploaded_files.append(file_info)
            
            # Use only uploaded files
            all_files = uploaded_files
            
            if all_files:
                st.subheader("📂 Select Original Result File to Update")
                
                # Create a selection interface
                selected_file = st.selectbox(
                    "Choose the original result file:",
                    options=all_files,
                    format_func=lambda x: f"{x['filename']} ({x.get('type', 'Unknown')}) - {x.get('semester', 'Unknown')} Sem, {x.get('year', 'Unknown')} - {x.get('student_count', 'Unknown')} students",
                    help="Select the file that you want to update with re-exam data"
                )
                
                if selected_file:
                    st.info(f"Selected: {selected_file['filename']}")
                    
                    # Show file details
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(f"**📅 Upload Date:** {selected_file.get('upload_date', 'Unknown')}")
                        st.markdown(f"**🎓 Semester:** {selected_file.get('semester', 'Unknown')} ({selected_file.get('year', 'Unknown')})")
                    with col2:
                        st.markdown(f"**👥 Students:** {selected_file.get('student_count', 'Unknown')}")
                        st.markdown(f"**📁 Original File:** {selected_file.get('original_file', 'Unknown')}")
                    
                    # Check for automatic filename matching
                    original_filename = selected_file['filename']
                    reexam_filename_clean = reexam_filename.replace(f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_", "")
                    
                    # Show automatic matching info
                    if original_filename.lower() in reexam_filename_clean.lower() or reexam_filename_clean.lower() in original_filename.lower():
                        st.success("✅ **Automatic Match Found!** Files appear to be related based on filename.")
                        st.info(f"Original: {original_filename} | Re-exam: {reexam_filename_clean}")
                    else:
                        st.warning("⚠️ **Manual Selection** - Files don't appear to match automatically. Please verify selection.")
                    
                    # Confirmation dialog
                    st.subheader("🔄 Update Confirmation")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        confirm_update = st.checkbox(
                            "✅ I confirm that I want to update this file with re-exam data",
                            help="Check this box to confirm the update operation"
                        )
                    with col2:
                        if confirm_update:
                            st.success("✅ Confirmation received!")
                        else:
                            st.warning("⚠️ Please confirm the update operation")
                    
                    # Update button (only enabled after confirmation)
                    if confirm_update:
                        if st.button("🔄 Update Results with Re-exam Data", type="primary"):
                            with st.spinner("Updating results with re-exam data..."):
                                try:
                                    # Get the original result file path
                                    original_file_path = selected_file.get('file_path')
                                    
                                    if os.path.exists(original_file_path):
                                        # Update results - pass the file object directly
                                        updated_df, updated_count = update_result_with_reexam(original_file_path, reexam_file)
                                        
                                        if updated_df is not None:
                                            # Update the original file directly (overwrite it)
                                            original_file_path = selected_file.get('file_path')
                                            updated_df.to_excel(original_file_path, index=False)
                                            
                                            # Update metadata for the original file
                                            metadata = load_metadata()
                                            
                                            # Find and update the existing metadata entry for this file
                                            file_updated = False
                                            for file_id, file_info in metadata.items():
                                                if file_info.get('file_path') == original_file_path:
                                                    # Update existing metadata
                                                    file_info['reexam_file'] = reexam_filename
                                                    file_info['type'] = 'reexam_updated'
                                                    file_info['last_updated'] = datetime.now().isoformat()
                                                    file_info['updated_count'] = updated_count
                                                    file_info['total_students'] = len(updated_df)
                                                    file_updated = True
                                                    break
                                            
                                            # If no existing metadata found, create a new entry
                                            if not file_updated:
                                                file_id = f"reexam_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                                                metadata[file_id] = {
                                                    "filename": selected_file['filename'],
                                                    "original_file": selected_file['filename'],
                                                    "reexam_file": reexam_filename,
                                                    "type": "reexam_updated",
                                                    "upload_date": selected_file.get('upload_date', 'Unknown'),
                                                    "last_updated": datetime.now().isoformat(),
                                                    "student_count": len(updated_df),
                                                    "updated_count": updated_count,
                                                    "file_path": original_file_path
                                                }
                                            
                                            save_metadata(metadata)
                                            
                                            # Success message box
                                            st.success(f"✅ **Original File Updated Successfully!**")
                                            st.info(f"""
                                            **📊 Update Summary:**
                                            - **File Updated:** {selected_file['filename']}
                                            - **Re-exam File:** {reexam_filename}
                                            - **Students Updated:** {updated_count} out of {len(updated_df)} total students
                                            - **File Location:** {os.path.dirname(original_file_path)}/
                                            - **⚠️ Note:** Original file has been modified with re-exam data
                                            """)
                                            
                                            # Show updated results
                                            st.subheader("📋 Updated Results (Original File Modified)")
                                            st.dataframe(updated_df)
                                            
                                            # Download updated file (same filename as original)
                                            buffer = io.BytesIO()
                                            updated_df.to_excel(buffer, index=False)
                                            st.download_button(
                                                label="📥 Download Updated File",
                                                data=buffer.getvalue(),
                                                file_name=selected_file['filename'],
                                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                key=f"download_reexam_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                                            )
                                        else:
                                            st.error("Failed to update results. Please check the file formats.")
                                    else:
                                        st.error(f"Original result file not found: {original_file_path}")
                                except Exception as e:
                                    st.error(f"Error updating results: {str(e)}")
            else:
                st.warning("No uploaded files found in the uploaded_files directory. Please upload some files first in the 'Process Results' tab.")

    with tab4:
        st.header("ℹ️ Instructions")
        st.markdown("""
        ### 📊 Process Results Tab
        1. **Upload Excel file** (.xlsx format)
        2. **Enter output filename** (auto-generated based on semester/year)
        3. **View processed results** with individual student details
        4. **Auto-save** - Files are automatically saved for future access
        
        ### 🔍 Search Files Tab
        1. **Search by filename** - Find saved files by name, semester, or year
        2. **Filter by type** - Uploaded files, processed results, or re-exam updates
        3. **Download/Print** - Access saved files with download and print options
        
        ### 🔄 Re-exam Update Tab
        1. **Upload original result file** (processed Excel file)
        2. **Upload re-exam ledger** (new Excel file with updated marks)
        3. **Automatic update** - System matches USNs and updates results directly in the original file
        4. **Recalculate totals** - Totals and percentages are automatically recalculated
        5. **⚠️ Important:** Original file is modified in place, no separate copy is created
        
        ### 📌 File Format Requirements
        - **Sheet1**: Subject details with subject codes and names
        - **Sheet2**: Result data with USN, Name, and subject marks
        - **Required columns**: USN, Name, Subjects, Result, Total Marks, Percentage
        
        ### 💾 Auto-Save Features
        - **Uploaded files** are automatically saved with timestamps
        - **Processed results** are saved with metadata (semester, year, student count)
        - **Re-exam updates** modify original files directly (no copies created)
        - **Search functionality** allows easy access to all saved files
        """)

if __name__ == "__main__":
    main()
