# Semester Result Analyzer

A professional Flask web application for analyzing semester results from Excel files.

## Features

- **Professional UI**: Built with Tailwind CSS for a modern, responsive design
- **Excel Processing**: Automatically processes Excel files from year-based folders
- **Batch Processing**: Processes all semester result files with a single click
- **Real-time Feedback**: Shows processing status and results
- **Error Handling**: Comprehensive error reporting and logging

## Project Structure

```
Sem Result/
├── app.py                 # Main Flask application
├── excelprocess.py        # Excel processing script
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Main UI template
├── 2023/                 # Year folders containing result files
│   ├── BBA/
│   │   ├── Sem 1.xlsx
│   │   ├── Sem 2.xlsx
│   │   └── Sem 3.xlsx
│   ├── BCA/
│   └── BCom/
└── README.md
```

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the Flask application:**
   ```bash
   python app.py
   ```

3. **Access the application:**
   Open your browser and go to `http://localhost:5000`

## Usage

1. **Start the Application**: Run `python app.py` to start the Flask server
2. **Process Excel Files**: Click the "Process All Excel Files" button on the main page
3. **View Results**: The application will show processing status and results
4. **Check Output**: Processed files will be saved with `_result.xlsx` suffix

## How It Works

1. The application scans for year folders (e.g., 2023, 2024, etc.)
2. Within each year folder, it looks for course folders (BBA, BCA, BCom, etc.)
3. It processes all Excel files in these folders using the `excelprocess.py` script
4. Results are displayed in real-time with success/error feedback
5. Processed files are saved alongside the original files

## Excel File Format

The application expects Excel files with:
- **Sheet1**: Contains subject codes and names
- **Sheet2**: Contains student results and grades

## Next Steps

This is the initial version with Excel processing functionality. Future enhancements will include:
- USN-based result lookup
- Individual student analysis
- Performance analytics and charts
- Export functionality

## Requirements

- Python 3.7+
- Flask 2.3.3
- pandas 2.1.1
- openpyxl 3.1.2

## Support

For issues or questions, please check the application logs or contact the development team.


