# Enhanced Semester Result Analysis App

## Overview
This enhanced Streamlit application provides a comprehensive solution for processing, managing, and updating semester results with advanced features including auto-save, search functionality, and re-exam updates.

## Features

### 📊 Process Results Tab
- **File Upload**: Upload Excel files (.xlsx format) with semester results
- **Auto-Save**: Files are automatically saved with timestamps
- **Smart Naming**: Output filenames are auto-generated based on semester and year
- **Real-time Processing**: View results as they are processed with progress indicators
- **Individual Student View**: Click "View" to see detailed student information
- **Download Results**: Download processed results as Excel files

### 🔍 Search Files Tab
- **Search Functionality**: Search saved files by filename, semester, or year
- **File Type Filtering**: Filter by uploaded files, processed results, or re-exam updates
- **File Metadata**: View upload date, semester, year, and student count
- **Download/Print Options**: Access saved files with download and print capabilities
- **Organized Display**: Files are displayed in expandable sections with detailed information

### 🔄 Re-exam Update Tab
- **Original File Upload**: Upload previously processed result files
- **Re-exam Ledger Upload**: Upload new Excel files with updated marks
- **Automatic USN Matching**: System automatically matches students by USN
- **Result Updates**: Subject marks are updated based on re-exam data
- **Automatic Recalculation**: Totals and percentages are recalculated automatically
- **Updated File Generation**: New files are created with updated results

### 💾 Auto-Save System
- **Uploaded Files**: Automatically saved in `uploaded_files/` directory
- **Processed Results**: Saved in `processed_results/` directory with metadata
- **Metadata Tracking**: JSON file tracks all file information and relationships
- **Timestamp Organization**: Files are organized with timestamps for easy identification

## File Structure
```
Result Analysis/
├── enhanced_result_analysis_app.py    # Main application
├── enhanced_requirements.txt          # Dependencies
├── README_Enhanced_App.md            # This file
├── uploaded_files/                   # Auto-saved uploaded files
├── processed_results/                # Auto-saved processed results
└── file_metadata.json               # File tracking metadata
```

## Installation and Setup

1. **Install Dependencies**:
   ```bash
   pip install -r enhanced_requirements.txt
   ```

2. **Run the Application**:
   ```bash
   streamlit run enhanced_result_analysis_app.py
   ```

3. **Access the App**:
   - Open your web browser
   - Navigate to `http://localhost:8501`

## File Format Requirements

### Input Excel Files
- **Sheet1**: Subject details with subject codes and names
- **Sheet2**: Result data with USN, Name, and subject marks
- **Required Structure**: 
  - Column 3: Subject codes (10-character alphanumeric)
  - Column 4: Subject names
  - USN column: Student USNs starting with 'U'
  - Result columns: Subject-wise marks and results

### Output Files
- **Processed Results**: Excel files with organized student data
- **Columns**: Sl. No, Name, USN, Subject Results, Total Marks, Percentage, CGPA, SGPA
- **Sorting**: Results sorted by percentage (highest to lowest)

## Usage Instructions

### Processing New Results
1. Go to "📊 Process Results" tab
2. Upload Excel file with semester results
3. Review auto-generated filename or enter custom name
4. Wait for processing to complete
5. View results and download processed file

### Searching Saved Files
1. Go to "🔍 Search Files" tab
2. Enter search query (filename, semester, year)
3. Select file type filter if needed
4. Browse search results
5. Download or print files as needed

### Updating with Re-exam Data
1. Go to "🔄 Re-exam Update" tab
2. Upload original processed result file
3. Upload re-exam ledger file
4. Click "Update Results with Re-exam Data"
5. Review updated results and download

## Technical Details

### Auto-Save System
- Files are saved with timestamps: `YYYYMMDD_HHMMSS_filename.xlsx`
- Metadata includes: filename, type, semester, year, upload date, student count
- JSON metadata file maintains file relationships and information

### Search Functionality
- Case-insensitive search by filename
- Filtering by file type (uploaded, processed, re-exam updated)
- Real-time search results with file details

### Re-exam Update Process
- USN-based matching between original and re-exam files
- Subject-wise mark updates
- Automatic total and percentage recalculation
- New file generation with update information

## Error Handling
- File format validation
- Missing data handling
- USN matching validation
- Progress indicators for long operations
- User-friendly error messages

## Browser Compatibility
- Modern web browsers (Chrome, Firefox, Safari, Edge)
- Responsive design for different screen sizes
- Mobile-friendly interface

## Support
For technical support or questions, contact:
- **Developer**: Mr. Pranam B
- **Department**: Computer Applications
- **Role**: Assistant Professor

## Version History
- **v2.0**: Enhanced version with auto-save, search, and re-exam features
- **v1.0**: Basic result processing functionality
