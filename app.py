from flask import Flask, render_template, request, jsonify, flash, redirect, url_for, send_file
import os
import subprocess
import sys
from pathlib import Path
import logging
import pandas as pd
import re
from werkzeug.utils import secure_filename

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this to a random secret key

# Get the current directory (where the app is running)
BASE_DIR = Path(__file__).parent
STUDENTS_DATA_FILE = BASE_DIR / "students_data.xlsx"
FILE_MANAGER_ALLOWED_ROOTS = {"2022", "2023", "2024", "Re-Exam"}

def normalize_relative_path(relative_path: str) -> str:
    if not relative_path:
        return ""
    normalized = relative_path.replace("\\", "/").strip()
    # Remove leading "./" or ".\" occurrences
    while normalized.startswith("./"):
        normalized = normalized[2:]
    normalized = normalized.strip("/")
    return normalized

def get_root_segment(relative_path: str) -> str:
    normalized = normalize_relative_path(relative_path)
    if not normalized:
        return ""
    return normalized.split("/", 1)[0]

def ensure_allowed_root(relative_path: str, allow_root_level=False):
    """
    Check if the root segment is in allowed roots.
    If allow_root_level is True, allow any root-level folder (for operations like delete/rename).
    """
    root_segment = get_root_segment(relative_path)
    if allow_root_level:
        # Allow any root-level folder for operations like delete/rename
        return
    if root_segment and root_segment not in FILE_MANAGER_ALLOWED_ROOTS:
        raise ValueError("Access to this location is restricted to permitted folders")

def ensure_allowed_absolute_path(abs_path: Path):
    relative = to_relative_path(abs_path)
    if not relative:
        return
    root_segment = relative.split("/", 1)[0]
    if root_segment and root_segment not in FILE_MANAGER_ALLOWED_ROOTS:
        raise ValueError("Access to this location is restricted to permitted folders")

def get_safe_path(relative_path: str = "", allow_root_level=False) -> Path:
    """
    Resolve a relative path against BASE_DIR ensuring it stays within the project directory.
    Raises ValueError if the resolved path escapes BASE_DIR.
    If allow_root_level is True, allows root-level folders without restriction.
    """
    try:
        ensure_allowed_root(relative_path, allow_root_level=allow_root_level)
        base_resolved = BASE_DIR.resolve()
        sanitized = (relative_path or "").strip()
        if sanitized in (".", "./"):
            sanitized = ""
        target_path = (base_resolved / Path(sanitized)).resolve()
        if not str(target_path).startswith(str(base_resolved)):
            raise ValueError("Path escapes project directory")
        if not allow_root_level:
            ensure_allowed_absolute_path(target_path)
        return target_path
    except Exception as exc:
        raise ValueError(str(exc))

def to_relative_path(abs_path: Path) -> str:
    """
    Convert an absolute path within BASE_DIR to a POSIX-style relative path string.
    """
    rel_path = abs_path.resolve().relative_to(BASE_DIR.resolve())
    rel_str = str(rel_path).replace("\\", "/")
    return "" if rel_str in (".", "") else rel_str

def summarize_directory(path: Path) -> dict:
    """
    Produce a serialisable summary of a directory entry for the file manager UI.
    """
    entry = {
        "name": path.name,
        "path": to_relative_path(path),
        "type": "directory" if path.is_dir() else "file",
        "has_subdirectories": False,
        "has_files": False
    }

    if path.is_dir():
        try:
            for child in path.iterdir():
                if child.name.startswith("."):
                    continue
                if child.is_dir():
                    entry["has_subdirectories"] = True
                else:
                    entry["has_files"] = True
                if entry["has_subdirectories"] and entry["has_files"]:
                    break
        except PermissionError:
            entry["error"] = "Permission denied"
    else:
        entry["has_files"] = True

    return entry

def format_file_size(num_bytes: int) -> str:
    if num_bytes is None:
        return "Unknown"
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if num_bytes < 1024.0:
            return f"{num_bytes:3.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} PB"

def format_timestamp(epoch_seconds: float) -> str:
    try:
        from datetime import datetime
        dt = datetime.fromtimestamp(epoch_seconds)
        return dt.strftime("%d %b %Y %I:%M %p")
    except Exception:
        return "Unknown"

def build_breadcrumbs(relative_path: str) -> list:
    """
    Build breadcrumb navigation data for the given relative path.
    """
    breadcrumbs = [{"name": "Home", "path": ""}]
    if not relative_path:
        return breadcrumbs

    clean_parts = [part for part in relative_path.split("/") if part not in ("", ".")]
    current_parts = []
    for part in clean_parts:
        current_parts.append(part)
        breadcrumbs.append({
            "name": part,
            "path": "/".join(current_parts)
        })
    return breadcrumbs

@app.route('/api/files', methods=['GET'])
def api_list_files():
    """
    List directory contents for the file manager UI.
    """
    relative_path = request.args.get('path', '').strip()
    try:
        target_path = get_safe_path(relative_path)
    except ValueError as exc:
        return jsonify({'success': False, 'error': str(exc)}), 400

    if not target_path.exists():
        return jsonify({'success': False, 'error': 'Path not found'}), 404

    if target_path.is_file():
        file_entry = summarize_directory(target_path)
        file_entry.pop("has_subdirectories", None)
        return jsonify({
            'success': True,
            'entry': file_entry,
            'breadcrumbs': build_breadcrumbs(to_relative_path(target_path.parent))
        })

    children = []
    try:
        for child in sorted(target_path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower())):
            if child.name.startswith("."):
                continue
            # In root directory: only show folders that match year format (4 digits like 2022, 2023, 2024)
            # or are in the allowed roots list (like "Re-Exam")
            if relative_path in ("", "/"):
                if child.is_dir():
                    # Check if folder name is a 4-digit year (2020-2099) or in allowed roots
                    is_year_format = (len(child.name) == 4 and child.name.isdigit() and 
                                     child.name.startswith(('20', '21')))
                    is_allowed_root = child.name in FILE_MANAGER_ALLOWED_ROOTS
                    if not (is_year_format or is_allowed_root):
                        continue
                else:
                    # Don't show files in root directory
                    continue
            children.append(summarize_directory(child))
    except PermissionError:
        return jsonify({'success': False, 'error': 'Permission denied'}), 403

    response_entry = {
        'name': target_path.name if relative_path else '',
        'path': to_relative_path(target_path),
        'type': 'directory',
        'children': children,
        'is_root': relative_path in ('', '/')
    }

    return jsonify({
        'success': True,
        'entry': response_entry,
        'breadcrumbs': build_breadcrumbs(relative_path)
    })

@app.route('/api/folders', methods=['POST'])
def api_create_folder():
    """
    Create a new folder under the specified parent directory.
    """
    data = request.get_json(silent=True) or {}
    parent_path = data.get('parent_path', '').strip()
    folder_name = (data.get('folder_name') or '').strip()

    if not folder_name:
        return jsonify({'success': False, 'error': 'Folder name is required'}), 400

    if any(sep in folder_name for sep in ('/', '\\')):
        return jsonify({'success': False, 'error': 'Folder name cannot contain path separators'}), 400

    try:
        parent_abs = get_safe_path(parent_path)
    except ValueError as exc:
        return jsonify({'success': False, 'error': str(exc)}), 400

    if not parent_abs.exists() or not parent_abs.is_dir():
        return jsonify({'success': False, 'error': 'Parent directory not found'}), 404

    new_folder_path = parent_abs / folder_name
    
    # Allow creating folders in root directory without restriction
    # Only check restrictions if creating inside a subdirectory
    is_root_directory = (not parent_path or parent_path.strip() == '')
    if not is_root_directory:
        try:
            ensure_allowed_absolute_path(new_folder_path)
        except ValueError as exc:
            return jsonify({'success': False, 'error': str(exc)}), 403
    
    if new_folder_path.exists():
        return jsonify({'success': False, 'error': 'A file or folder with the same name already exists'}), 409

    try:
        new_folder_path.mkdir(parents=False, exist_ok=False)
    except Exception as exc:
        logger.error(f"Error creating folder {new_folder_path}: {exc}")
        return jsonify({'success': False, 'error': f'Unable to create folder: {exc}'}), 500

    return jsonify({
        'success': True,
        'message': 'Folder created successfully',
        'folder': summarize_directory(new_folder_path)
    }), 201

@app.route('/api/upload', methods=['POST'])
def api_upload_files():
    """
    Upload one or more files to a specified directory.
    """
    target_relative_path = request.form.get('path', '').strip()
    try:
        target_directory = get_safe_path(target_relative_path)
    except ValueError as exc:
        return jsonify({'success': False, 'error': str(exc)}), 400

    if not target_directory.exists() or not target_directory.is_dir():
        return jsonify({'success': False, 'error': 'Target directory not found'}), 404

    uploaded_files = request.files.getlist('files')
    if not uploaded_files:
        return jsonify({'success': False, 'error': 'No files uploaded'}), 400

    saved_files = []
    for file_storage in uploaded_files:
        filename = secure_filename(file_storage.filename)
        if not filename:
            continue

        destination = target_directory / filename
        try:
            ensure_allowed_absolute_path(destination)
        except ValueError as exc:
            return jsonify({'success': False, 'error': str(exc)}), 403
        counter = 1
        while destination.exists():
            stem = destination.stem
            suffix = destination.suffix
            new_name = f"{stem}_{counter}{suffix}"
            destination = target_directory / new_name
            counter += 1

        try:
            file_storage.save(destination)
            saved_files.append(summarize_directory(destination))
        except Exception as exc:
            logger.error(f"Error saving uploaded file to {destination}: {exc}")
            return jsonify({'success': False, 'error': f'Error saving file: {exc}'}), 500

    if not saved_files:
        return jsonify({'success': False, 'error': 'No valid files uploaded'}), 400

    return jsonify({
        'success': True,
        'message': f'Uploaded {len(saved_files)} file(s) successfully',
        'files': saved_files
    }), 201

@app.route('/api/rename', methods=['POST'])
def api_rename_entry():
    """
    Rename an existing file or folder within BASE_DIR.
    """
    data = request.get_json(silent=True) or {}
    current_path = data.get('path', '').strip()
    new_name = (data.get('new_name') or '').strip()

    if not current_path:
        return jsonify({'success': False, 'error': 'Path is required'}), 400

    if not new_name:
        return jsonify({'success': False, 'error': 'New name is required'}), 400

    if any(sep in new_name for sep in ('/', '\\')):
        return jsonify({'success': False, 'error': 'New name cannot contain path separators'}), 400

    # Get the absolute path - allow root-level folders
    try:
        current_abs = get_safe_path(current_path, allow_root_level=True)
    except ValueError as exc:
        return jsonify({'success': False, 'error': str(exc)}), 400

    if not current_abs.exists():
        return jsonify({'success': False, 'error': 'Specified path not found'}), 404

    if current_abs == BASE_DIR.resolve():
        return jsonify({'success': False, 'error': 'Cannot rename the root directory'}), 403

    # Check if this is a root-level folder (parent is BASE_DIR)
    is_root_level = (current_abs.parent == BASE_DIR.resolve())
    
    new_abs = current_abs.parent / new_name
    
    # Allow renaming root-level folders without restriction
    # Only check restrictions if renaming inside a subdirectory
    if not is_root_level:
        try:
            ensure_allowed_absolute_path(new_abs)
        except ValueError as exc:
            return jsonify({'success': False, 'error': str(exc)}), 403
    if new_abs.exists():
        return jsonify({'success': False, 'error': 'A file or folder with the new name already exists'}), 409

    try:
        current_abs.rename(new_abs)
    except Exception as exc:
        logger.error(f"Error renaming {current_abs} to {new_abs}: {exc}")
        return jsonify({'success': False, 'error': f'Unable to rename: {exc}'}), 500

    return jsonify({
        'success': True,
        'message': 'Renamed successfully',
        'entry': summarize_directory(new_abs)
    })

@app.route('/api/delete', methods=['POST'])
def api_delete_entry():
    """
    Delete a file or folder within BASE_DIR.
    """
    data = request.get_json(silent=True) or {}
    path_to_delete = data.get('path', '').strip()

    if not path_to_delete:
        return jsonify({'success': False, 'error': 'Path is required'}), 400

    # Get the absolute path - allow root-level folders
    try:
        target_abs = get_safe_path(path_to_delete, allow_root_level=True)
    except ValueError as exc:
        return jsonify({'success': False, 'error': str(exc)}), 400

    if not target_abs.exists():
        return jsonify({'success': False, 'error': 'Path not found'}), 404

    if target_abs == BASE_DIR.resolve():
        return jsonify({'success': False, 'error': 'Cannot delete the root directory'}), 403

    try:
        if target_abs.is_dir():
            import shutil
            shutil.rmtree(target_abs)
        else:
            target_abs.unlink()
    except Exception as exc:
        logger.error(f"Error deleting {target_abs}: {exc}")
        return jsonify({'success': False, 'error': f'Unable to delete: {exc}'}), 500

    return jsonify({
        'success': True,
        'message': 'Deleted successfully'
    })

@app.route('/files/open')
def open_file():
    """
    Stream a file from the project directory to the browser.
    Opens files inline (no download dialog) using proper MIME types.
    """
    relative_path = request.args.get('path', '').strip()
    if not relative_path:
        return jsonify({'success': False, 'error': 'File path is required'}), 400

    try:
        file_path = get_safe_path(relative_path)
    except ValueError as exc:
        return jsonify({'success': False, 'error': str(exc)}), 400

    if not file_path.exists() or not file_path.is_file():
        return jsonify({'success': False, 'error': 'File not found'}), 404

    # Determine MIME type based on file extension
    mime_type = None
    suffix = file_path.suffix.lower()
    mime_map = {
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.xls': 'application/vnd.ms-excel',
        '.pdf': 'application/pdf',
        '.txt': 'text/plain',
        '.csv': 'text/csv',
        '.html': 'text/html',
        '.json': 'application/json',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
    }
    mime_type = mime_map.get(suffix, 'application/octet-stream')

    # Use send_file with as_attachment=False and explicit Content-Disposition: inline
    response = send_file(
        str(file_path),
        mimetype=mime_type,
        as_attachment=False
    )
    
    # Explicitly set Content-Disposition to inline to prevent download dialog
    response.headers['Content-Disposition'] = f'inline; filename="{file_path.name}"'
    
    return response

@app.route('/file-viewer')
def file_viewer():
    """
    Render a simple preview of an Excel file (xlsx/xls) from allowed directories.
    """
    relative_path = request.args.get('path', '').strip()
    if not relative_path:
        return render_template('file_viewer.html', error="File path is required"), 400

    try:
        file_path = get_safe_path(relative_path)
    except ValueError as exc:
        return render_template('file_viewer.html', error=str(exc)), 400

    if not file_path.exists() or not file_path.is_file():
        return render_template('file_viewer.html', error="File not found"), 404

    if file_path.suffix.lower() not in ('.xlsx', '.xls'):
        return render_template('file_viewer.html', file_name=file_path.name,
                               error="Preview is only available for Excel files (.xlsx, .xls)"), 400

    table_html = None
    error = None
    try:
        df = pd.read_excel(file_path)
        if df.empty:
            table_html = "<p class='text-center text-gray-500 py-6'>This file has no records.</p>"
        else:
            table_html = df.to_html(
                classes="min-w-full divide-y divide-gray-200 text-sm text-left",
                index=False,
                border=0,
                justify='center'
            )
    except Exception as exc:
        logger.error(f"Error previewing {file_path}: {exc}")
        error = f"Unable to read Excel file: {exc}"

    stats = file_path.stat()
    return render_template(
        'file_viewer.html',
        file_name=file_path.name,
        relative_path=to_relative_path(file_path),
        table_html=table_html,
        error=error,
        file_size=format_file_size(stats.st_size) if stats else "Unknown",
        modified_at=format_timestamp(stats.st_mtime) if stats else None
    )

@app.route('/api/open-directory', methods=['POST'])
def api_open_directory():
    """
    Open a directory in the host operating system's file explorer.
    Only allowed for directories under permitted roots.
    """
    data = request.get_json(silent=True) or {}
    relative_path = data.get('path', '').strip()
    if not relative_path:
        return jsonify({'success': False, 'error': 'Path is required'}), 400

    try:
        abs_path = get_safe_path(relative_path)
    except ValueError as exc:
        return jsonify({'success': False, 'error': str(exc)}), 400

    if not abs_path.exists() or not abs_path.is_dir():
        return jsonify({'success': False, 'error': 'Directory not found'}), 404

    try:
        if os.name == 'nt':
            os.startfile(str(abs_path))  # type: ignore[attr-defined]
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', str(abs_path)])
        else:
            subprocess.Popen(['xdg-open', str(abs_path)])
    except Exception as exc:
        logger.error(f"Error opening directory {abs_path}: {exc}")
        return jsonify({'success': False, 'error': f'Unable to open directory: {exc}'}), 500

    return jsonify({
        'success': True,
        'message': f'Opened {abs_path}'
    })

def load_students_data():
    """Load students data from students_data.xlsx file"""
    if not STUDENTS_DATA_FILE.exists():
        logger.warning(f"Students data file not found: {STUDENTS_DATA_FILE}")
        return {}
    
    try:
        # Read all sheets from the Excel file
        xl_file = pd.ExcelFile(STUDENTS_DATA_FILE)
        students_data = {}
        
        for sheet_name in xl_file.sheet_names:
            df = pd.read_excel(STUDENTS_DATA_FILE, sheet_name=sheet_name)
            
            # Create mappings for both roll number to USN and USN to roll number
            for _, row in df.iterrows():
                roll_no = str(row['Roll No']).strip()
                usn = str(row['USN']).strip()
                
                # Store roll number to USN mapping
                students_data[roll_no] = {
                    'usn': usn,
                    'class': sheet_name
                }
                
                # Store USN to roll number mapping
                students_data[usn] = {
                    'roll_no': roll_no,
                    'class': sheet_name
                }
        
        logger.info(f"Loaded students data with {len(students_data)} entries from {len(xl_file.sheet_names)} sheets")
        return students_data
    
    except Exception as e:
        logger.error(f"Error loading students data: {e}")
        return {}

def get_year_folders():
    """Get all year folders in the base directory"""
    year_folders = []
    for item in BASE_DIR.iterdir():
        if item.is_dir() and item.name.isdigit():
            year_folders.append(item.name)
    return sorted(year_folders, reverse=True)

def get_excel_files_in_year(year):
    """Get all Excel files in a specific year folder"""
    year_path = BASE_DIR / year
    excel_files = []
    
    if year_path.exists():
        for item in year_path.iterdir():
            if item.is_dir():
                # Look for Excel files in subdirectories (like BBA, BCA, etc.)
                for excel_file in item.glob("*.xlsx"):
                    if not excel_file.name.endswith("_result.xlsx"):
                        result_file = excel_file.parent / f"{excel_file.stem}_result.xlsx"
                        excel_files.append({
                            'path': str(excel_file),
                            'name': excel_file.name,
                            'course': item.name,
                            'semester': excel_file.stem,
                            'result_exists': result_file.exists(),
                            'result_path': str(result_file)
                        })
    return excel_files

def get_processing_status():
    """Get the overall processing status of all Excel files"""
    year_folders = get_year_folders()
    all_files = []
    files_needing_processing = []
    files_ready = []
    
    for year in year_folders:
        excel_files = get_excel_files_in_year(year)
        all_files.extend(excel_files)
        
        for file_info in excel_files:
            if file_info['result_exists']:
                files_ready.append(file_info)
            else:
                files_needing_processing.append(file_info)
    
    return {
        'total_files': len(all_files),
        'files_ready': len(files_ready),
        'files_needing_processing': len(files_needing_processing),
        'all_ready': len(files_needing_processing) == 0,
        'files_needing_processing_list': files_needing_processing,
        'files_ready_list': files_ready
    }

def parse_usn_or_roll(identifier):
    """Parse USN or roll number to extract student information"""
    identifier = identifier.strip().upper()
    
    # Load students data for roll number lookup
    students_data = load_students_data()
    
    # Check if it's a roll number (lookup in students_data)
    if identifier in students_data and 'usn' in students_data[identifier]:
        # It's a roll number, get the USN
        usn = students_data[identifier]['usn']
        class_name = students_data[identifier]['class']
        logger.info(f"Found roll number {identifier} -> USN: {usn}, Class: {class_name}")
    else:
        # Assume it's a USN
        usn = identifier
        class_name = None
        if usn in students_data and 'roll_no' in students_data[usn]:
            class_name = students_data[usn]['class']
            logger.info(f"Found USN {usn} -> Class: {class_name}")
    
    # Parse USN to extract year, course, and student number
    if len(usn) != 12:
        return None
    
    # Extract year and course code from USN
    # Format: U05BB23S0037
    match = re.match(r'^U05BB(\d{2})([SCM])(\d{4})$', usn)
    if not match:
        return None
    
    year = match.group(1)
    course_code = match.group(2)
    student_number = match.group(3)
    
    # Map course codes to course names
    course_mapping = {
        'S': 'BCA',
        'M': 'BBA', 
        'C': 'BCOM'
    }
    
    course = course_mapping.get(course_code)
    if not course:
        return None
    
    # Convert 2-digit year to 4-digit year
    # Assuming 20-29 means 2020-2029, 30-99 means 1930-1999
    if int(year) >= 20:
        full_year = f"20{year}"
    else:
        full_year = f"19{year}"
    
    # Get roll number if available
    roll_no = None
    if usn in students_data and 'roll_no' in students_data[usn]:
        roll_no = students_data[usn]['roll_no']
    
    return {
        'usn': usn,
        'roll_no': roll_no,
        'class': class_name,
        'year': full_year,
        'course': course,
        'course_code': course_code,
        'student_number': student_number
    }

def find_student_results(usn_info):
    """Find all result files for a specific student"""
    year = usn_info['year']
    course = usn_info['course']
    usn = usn_info['usn']
    
    # Look for result files in the specific year and course folder
    year_path = BASE_DIR / year
    course_path = year_path / course
    
    if not course_path.exists():
        return None
    
    # Find all result files for this course
    result_files = []
    for result_file in course_path.glob("*_result.xlsx"):
        try:
            # Read the result file
            df = pd.read_excel(result_file)
            
            # Check if USN exists in this file
            usn_row = df[df['USN'] == usn]
            if not usn_row.empty:
                semester = result_file.stem.replace('_result', '')
                logger.info(f"Found USN {usn} in {result_file.name}, semester: {semester}")
                result_files.append({
                    'semester': semester,
                    'file_path': str(result_file),
                    'data': usn_row.iloc[0].to_dict()
                })
        except Exception as e:
            logger.error(f"Error reading {result_file}: {e}")
            continue
    
    return result_files if result_files else None

def find_all_student_results_across_years(usn_info):
    """Find all result files for a specific student across all years, including Re-Exam folder"""
    usn = usn_info['usn']
    course = usn_info['course']
    
    # Get all year folders
    year_folders = get_year_folders()
    all_result_files = []
    
    # Also include Re-Exam folder for re-examination results
    folders_to_search = year_folders + ['Re-Exam']
    
    for year in folders_to_search:
        year_path = BASE_DIR / year
        
        # Skip if folder doesn't exist
        if not year_path.exists() or not year_path.is_dir():
            continue
            
        course_path = year_path / course
        
        if not course_path.exists():
            continue
            
        # Find all result files for this course in this year/folder
        for result_file in course_path.glob("*_result.xlsx"):
            try:
                # Read the result file
                df = pd.read_excel(result_file)
                
                # Check if USN exists in this file
                usn_row = df[df['USN'] == usn]
                if not usn_row.empty:
                    semester = result_file.stem.replace('_result', '')
                    logger.info(f"Found USN {usn} in {year}/{course}/{result_file.name}, semester: {semester}")
                    all_result_files.append({
                        'year': year,
                        'course': course,
                        'semester': semester,
                        'file_path': str(result_file),
                        'data': usn_row.iloc[0].to_dict()
                    })
            except Exception as e:
                logger.error(f"Error reading {result_file}: {e}")
                continue
    
    return all_result_files if all_result_files else None

def find_all_student_results_across_all_folders(usn):
    """
    Find all result files for a specific student across ALL folders and ALL courses.
    This comprehensive search is used for re-examination updates to find results in any folder.
    """
    all_result_files = []
    
    # Search in all allowed root folders
    folders_to_search = list(FILE_MANAGER_ALLOWED_ROOTS)
    
    for folder_name in folders_to_search:
        folder_path = BASE_DIR / folder_name
        
        # Skip if folder doesn't exist
        if not folder_path.exists() or not folder_path.is_dir():
            continue
        
        # Search in ALL courses within this folder
        for course_dir in folder_path.iterdir():
            if not course_dir.is_dir():
                continue
            
            course = course_dir.name
            
            # Find all result files in this course
            for result_file in course_dir.glob("*_result.xlsx"):
                try:
                    # Read the result file
                    df = pd.read_excel(result_file)
                    
                    # Check if USN exists in this file
                    usn_row = df[df['USN'] == usn]
                    if not usn_row.empty:
                        semester = result_file.stem.replace('_result', '')
                        logger.info(f"Found USN {usn} in {folder_name}/{course}/{result_file.name}, semester: {semester}")
                        all_result_files.append({
                            'year': folder_name,
                            'course': course,
                            'semester': semester,
                            'file_path': str(result_file),
                            'data': usn_row.iloc[0].to_dict()
                        })
                except Exception as e:
                    logger.error(f"Error reading {result_file}: {e}")
                    continue
    
    return all_result_files if all_result_files else None

def calculate_student_cgpa(usn_info):
    """
    Get CGPA for a specific student from the most recent semester result file
    in the same year folder (admission year).
    Returns the CGPA value from the most recent semester file, or None if blank/not found.
    """
    usn = usn_info['usn']
    course = usn_info['course']
    year = usn_info['year']
    
    # Look only in the admission year folder (same folder)
    year_path = BASE_DIR / year
    course_path = year_path / course
    
    if not course_path.exists():
        logger.warning(f"Course path does not exist: {course_path}")
        return None
    
    # Helper function to extract semester number from filename
    def get_semester_number(filename):
        """Extract semester number from filename (e.g., 'Sem 1_result.xlsx' -> 1)"""
        try:
            # Remove _result.xlsx extension
            name = filename.replace('_result.xlsx', '').replace('_result.xls', '')
            # Look for numbers in the name
            match = re.search(r'(\d+)', str(name))
            if match:
                return int(match.group(1))
        except Exception:
            pass
        return 0
    
    # Find all result files for this student in the same folder
    student_result_files = []
    
    for result_file in course_path.glob("*_result.xlsx"):
        try:
            # Read the result file
            df = pd.read_excel(result_file)
            
            # Check if USN exists in this file
            usn_row = df[df['USN'] == usn]
            if not usn_row.empty:
                semester_name = result_file.stem.replace('_result', '')
                semester_num = get_semester_number(result_file.name)
                cgpa_value = usn_row.iloc[0].get('CGPA')
                
                student_result_files.append({
                    'semester_num': semester_num,
                    'semester_name': semester_name,
                    'file_path': str(result_file),
                    'cgpa': cgpa_value
                })
                logger.debug(f"Found result for {usn} in {result_file.name}, semester: {semester_name}, CGPA: {cgpa_value}")
        except Exception as e:
            logger.error(f"Error reading {result_file} for CGPA: {e}")
            continue
    
    if not student_result_files:
        logger.warning(f"No result files found for {usn} in {course_path}")
        return None
    
    # Sort by semester number (descending) to get the most recent
    student_result_files.sort(key=lambda x: x['semester_num'], reverse=True)
    most_recent = student_result_files[0]
    
    logger.info(f"Most recent semester for {usn}: {most_recent['semester_name']} (semester {most_recent['semester_num']})")
    
    # Get CGPA from the most recent semester file
    cgpa = most_recent['cgpa']
    
    # Check if CGPA is valid (not blank, not null, not empty string)
    if pd.isna(cgpa) or cgpa is None or str(cgpa).strip() == '':
        logger.info(f"CGPA is blank for {usn} in most recent semester {most_recent['semester_name']}")
        return None
    
    try:
        # Try to convert to float to ensure it's a valid number
        cgpa_float = float(cgpa)
        logger.info(f"Found CGPA {cgpa_float} for {usn} from most recent semester {most_recent['semester_name']}")
        return cgpa_float
    except (ValueError, TypeError):
        logger.warning(f"CGPA value '{cgpa}' is not a valid number for {usn}")
        return None

def check_for_re_examination_updates(usn_info):
    """
    Check if there are re-examination results that should update older records.
    
    This function implements the re-examination update logic:
    1. Finds all result files for a student across ALL folders and ALL courses
    2. Looks for PASS results in newer folders/years for subjects that show FAIL/ABSENT in older folders/years
    3. Returns a list of updates that need to be applied
    
    Example scenario:
    - Student fails Math in 2023 Sem 1 (shows "FAIL")
    - Student passes Math in 2024 re-exam or Re-Exam folder (shows "PASS" or "65-PASS")
    - This function will identify that 2023 Sem 1 Math result should be updated to "PASS"
    """
    usn = usn_info['usn']
    
    # Use comprehensive search across ALL folders and ALL courses
    all_results = find_all_student_results_across_all_folders(usn)
    if not all_results:
        return None
    
    # Sort results by year/folder (newest first)
    # Treat "Re-Exam" as the newest (highest priority) for re-examination updates
    def sort_key(result):
        year = result['year']
        if year == 'Re-Exam':
            # Re-Exam should be treated as newest (highest priority)
            return (9999, 0)  # Use a very high number to ensure it's first
        try:
            year_int = int(year)
            return (year_int, 0)
        except (ValueError, TypeError):
            # If year is not a number, put it at the end
            return (0, 0)
    
    all_results.sort(key=sort_key, reverse=True)
    
    updates_needed = []
    
    # For each newer result, check if it has pass/absent results that should update older fails/absents
    for i, newer_result in enumerate(all_results):
        newer_year = newer_result['year']  # Keep as string, don't convert to int
        
        newer_data = newer_result['data']
        
        # Check each subject in the newer result
        for col, value in newer_data.items():
            if ' - ' in col and col not in ['Sl. No', 'Name', 'USN', 'Result', 'Total Marks', 'Max Marks', 'Percentage', 'CGPA', 'SGPA']:
                if pd.notna(value) and str(value).strip() != 'N/A' and str(value).strip() != '':
                    value_str = str(value).strip()
                    value_upper = value_str.upper()
                    
                    # Check if this is a PASS or ABSENT result
                    is_pass = False
                    is_absent = False
                    
                    # Check for ABSENT status first
                    if 'ABSENT' in value_upper or ('-' in value_str and 'ABSENT' in value_str.split('-')[-1].upper()):
                        is_absent = True
                    # Check for PASS status
                    elif 'PASS' in value_upper or ('-' in value_str and 'PASS' in value_str.split('-')[-1].upper()):
                        is_pass = True
                    
                    # If it's a PASS or ABSENT result, look for older results to update
                    if is_pass or is_absent:
                        # Look for older results that have FAIL/ABSENT for this subject
                        for j in range(i + 1, len(all_results)):
                            older_result = all_results[j]
                            older_year = older_result['year']  # Keep as string
                            
                            older_data = older_result['data']
                            
                            if col in older_data:
                                older_value = older_data[col]
                                if pd.notna(older_value) and str(older_value).strip() != 'N/A':
                                    older_value_str = str(older_value).strip()
                                    older_status_upper = older_value_str.upper()
                                    
                                    # Check if older result is FAIL or ABSENT
                                    is_older_fail = False
                                    is_older_absent = False
                                    
                                    if 'FAIL' in older_status_upper or ('-' in older_status_upper and 'FAIL' in older_status_upper.split('-')[-1]):
                                        is_older_fail = True
                                    elif 'ABSENT' in older_status_upper or ('-' in older_status_upper and 'ABSENT' in older_status_upper.split('-')[-1]):
                                        is_older_absent = True
                                    
                                    # Update if older result is FAIL or ABSENT
                                    # If newer result is ABSENT, update to show ABSENT
                                    # If newer result is PASS, update to show PASS (only if older was FAIL or ABSENT)
                                    if is_older_fail or is_older_absent:
                                        # Format the update value properly
                                        if is_absent:
                                            # If re-examination shows ABSENT, update to show ABSENT
                                            # Preserve marks format if available, otherwise just show ABSENT
                                            if '/' in value_str and '-' in value_str:
                                                # Format: "marks/max-ABSENT" or "marks-ABSENT"
                                                update_value = value_str
                                            elif '-' in value_str:
                                                # Format: "marks-ABSENT" or just "ABSENT"
                                                update_value = value_str
                                            else:
                                                # Just ABSENT, format as "0/0-ABSENT" or "ABSENT"
                                                update_value = "ABSENT"
                                        else:
                                            # PASS result, use the newer result value
                                            update_value = value_str
                                        
                                        updates_needed.append({
                                            'subject': col,
                                            'newer_result': update_value,
                                            'older_file': older_result['file_path'],
                                            'older_year': older_year,
                                            'older_course': older_result.get('course', 'Unknown'),
                                            'older_semester': older_result['semester'],
                                            'newer_file': newer_result['file_path'],
                                            'newer_year': newer_year,
                                            'newer_course': newer_result.get('course', 'Unknown'),
                                            'newer_semester': newer_result['semester'],
                                            'newer_row': newer_data,
                                            'update_type': 'ABSENT' if is_absent else 'PASS'
                                        })
                                        logger.info(f"Update needed: {col} in {older_year}/{older_result.get('course', 'Unknown')}/{older_result['semester']} should be updated from {older_value_str} to {update_value} (re-exam from {newer_year}/{newer_result.get('course', 'Unknown')}: {'ABSENT' if is_absent else 'PASS'})")
    
    return updates_needed if updates_needed else None

def apply_re_examination_updates(usn_info):
    """Apply re-examination updates to older result files"""
    updates_needed = check_for_re_examination_updates(usn_info)
    if not updates_needed:
        return {'success': True, 'message': 'No updates needed', 'updates_applied': 0}
    
    applied_updates = []
    errors = []
    
    # Group updates by file to minimize file operations
    updates_by_file = {}
    for update in updates_needed:
        file_path = update['older_file']
        if file_path not in updates_by_file:
            updates_by_file[file_path] = []
        updates_by_file[file_path].append(update)
    
    # Apply updates to each file
    for file_path, file_updates in updates_by_file.items():
        try:
            # Read the current file
            df = pd.read_excel(file_path)
            
            # Find the student's row
            usn = usn_info['usn']
            student_mask = df['USN'] == usn
            if not student_mask.any():
                logger.error(f"Student {usn} not found in {file_path}")
                continue
            
            # Apply updates
            updated = False
            summary_updates = {}
            for update in file_updates:
                subject = update['subject']
                new_value = update['newer_result']
                newer_row = update.get('newer_row', {})
                
                if subject in df.columns:
                    old_value = df.loc[student_mask, subject].iloc[0]
                    df.loc[student_mask, subject] = new_value
                    updated = True
                    applied_updates.append({
                        'file': file_path,
                        'subject': subject,
                        'old_value': str(old_value),
                        'new_value': new_value,
                        'year': update['older_year'],
                        'semester': update['older_semester']
                    })
                    logger.info(f"Updated {subject} in {file_path} to {new_value}")

                # Collect summary fields from the newer attempt (no recalculation)
                for summary_field in ['Total Marks', 'Max Marks', 'Percentage', 'Result', 'CGPA', 'SGPA']:
                    if summary_field in newer_row and pd.notna(newer_row[summary_field]):
                        summary_updates[summary_field] = newer_row[summary_field]
            
            # Save the updated file if changes were made
            if updated:
                # Update summary columns with values pulled from the newer record (if available)
                for field, value in summary_updates.items():
                    if field in df.columns:
                        df.loc[student_mask, field] = value
                        logger.info(f"Updated {field} in {file_path} to {value}")

                df.to_excel(file_path, index=False)
                logger.info(f"Successfully updated {file_path}")
            
        except Exception as e:
            error_msg = f"Error updating {file_path}: {str(e)}"
            logger.error(error_msg)
            errors.append(error_msg)
    
    return {
        'success': len(errors) == 0,
        'message': f'Applied {len(applied_updates)} updates',
        'updates_applied': len(applied_updates),
        'applied_updates': applied_updates,
        'errors': errors
    }

def format_student_data(student_results, usn_info):
    """Format student data for display"""
    if not student_results:
        return None
    
    try:
        # Get basic info from first result
        first_result = student_results[0]['data']
        student_name = first_result.get('Name', 'Not Found')
        
        formatted_results = []

        # Helper to extract semester index as integer (e.g., "Sem 3" -> 3)
        def get_semester_index(sem_name):
            try:
                match = re.search(r'(\d+)', str(sem_name))
                return int(match.group(1)) if match else 0
            except Exception:
                return 0

        # Helper to extract status token from a raw value string
        def extract_status_token(raw_value):
            if pd.isna(raw_value):
                return None
            value_str = str(raw_value).strip()
            if value_str == '' or value_str.upper() == 'N/A':
                return None
            # Formats can be "marks/max-status" or "marks-status" or just "marks"
            if '-' in value_str:
                status_part = value_str.split('-', 1)[-1].strip()
                return status_part.upper()
            return None

        # Sort results by semester (ascending) and then by year
        def sort_key(result_entry):
            semester_key = get_semester_index(result_entry.get('semester', ''))
            year_str = str(result_entry.get('year', '')).strip()
            try:
                year_key = int(year_str)
            except ValueError:
                year_key = 0
            return (semester_key, year_key)

        results_sorted = sorted(student_results, key=sort_key)

        # Group attempts by canonical semester name (e.g., "SEM 2")
        def canonical_semester_name(name):
            if not name:
                return str(name).strip()
            match = re.search(r'(sem\s*\d+)', str(name), re.IGNORECASE)
            if match:
                return match.group(1).upper().replace('  ', ' ')
            return str(name).strip()

        semester_groups = {}
        for res in results_sorted:
            canon = canonical_semester_name(res.get('semester', ''))
            semester_groups.setdefault(canon, []).append(res)

        # Helper functions defined before the loop for scope access
        def format_value(value, default='N/A'):
            """Format a value, handling NaN, None, and empty strings"""
            if pd.isna(value) or value is None or str(value).strip() == '':
                return default
            return str(value).strip()
        
        def to_float_safe(val):
            """Safely convert value to float"""
            try:
                if isinstance(val, str):
                    v = val.replace('%', '').replace(',', '').strip()
                else:
                    v = val
                return float(v)
            except Exception:
                return None
        
        def format_marks_as_int(marks_str):
            """Format marks string to show integers only"""
            if marks_str == 'Absent' or marks_str == 'N/A' or not marks_str:
                return marks_str
            try:
                # Handle "marks/max" format
                if '/' in marks_str:
                    parts = marks_str.split('/')
                    obtained = to_float_safe(parts[0].strip())
                    max_marks = to_float_safe(parts[1].strip()) if len(parts) > 1 else None
                    if obtained is not None:
                        obtained = int(round(obtained))
                    if max_marks is not None:
                        max_marks = int(round(max_marks))
                    if max_marks is not None:
                        return f"{obtained}/{max_marks}"
                    else:
                        return str(obtained) if obtained is not None else marks_str
                else:
                    # Just marks
                    marks_val = to_float_safe(marks_str)
                    if marks_val is not None:
                        return str(int(round(marks_val)))
                    return marks_str
            except Exception:
                return marks_str
        
        def parse_marks(marks_str):
            """Parse marks string to extract obtained and max marks as integers"""
            if marks_str == 'Absent' or marks_str == 'N/A':
                return None, None
            try:
                # Handle "marks/max" format
                if '/' in marks_str:
                    parts = marks_str.split('/')
                    obtained = to_float_safe(parts[0].strip())
                    max_marks = to_float_safe(parts[1].strip()) if len(parts) > 1 else None
                    # Round to integers
                    if obtained is not None:
                        obtained = round(obtained)
                    if max_marks is not None:
                        max_marks = round(max_marks)
                    return obtained, max_marks
                else:
                    # Just marks, assume max marks from context or return None
                    obtained = to_float_safe(marks_str)
                    if obtained is not None:
                        obtained = round(obtained)
                    return obtained, None
            except Exception:
                return None, None
        
        def percentage_to_sgpa(percentage):
            """Convert percentage to SGPA (10-point scale)
            SGPA is calculated as percentage / 10 for linear conversion
            """
            if percentage is None or percentage < 0:
                return 0.0
            # Linear conversion: SGPA = Percentage / 10
            # Round to 2 decimal places
            return round(percentage / 10.0, 2)
        
        def percentage_to_grade_point(percentage):
            """Convert percentage to grade point (10-point scale with thresholds)
            This is used for individual subject grade points
            """
            if percentage is None:
                return 0.0
            if percentage >= 90:
                return 10.0
            elif percentage >= 80:
                return 9.0
            elif percentage >= 70:
                return 8.0
            elif percentage >= 60:
                return 7.0
            elif percentage >= 50:
                return 6.0
            elif percentage >= 40:
                return 5.0
            else:
                return 0.0

        for canon_name, attempts in semester_groups.items():
            logger.debug(f"Processing canonical semester group: {canon_name}")
            logger.debug(f"Attempts in group: {[a.get('semester') for a in attempts]}")
            attempts_sorted = sorted(attempts, key=sort_key)
            latest_attempt = attempts_sorted[-1]
            latest_data = latest_attempt['data']
            display_semester = latest_attempt.get('semester') or canon_name
            if isinstance(display_semester, str) and '_' in display_semester:
                display_semester = canonical_semester_name(display_semester).title()
            
            subject_order = []
            subject_choices = {}
            for idx, attempt in enumerate(attempts_sorted):
                logger.debug(f"  Attempt idx={idx}, semester={attempt.get('semester')}, year={attempt.get('year')}")
                attempt_data = attempt.get('data', {})
                if not isinstance(attempt_data, dict):
                    continue
                for col, value in attempt_data.items():
                    if ' - ' in col and col not in ['Sl. No', 'Name', 'USN', 'Result', 'Total Marks', 'Max Marks', 'Percentage', 'CGPA', 'SGPA']:
                        if pd.isna(value) or str(value).strip() in ('', 'N/A'):
                            continue
                        if col not in subject_order:
                            subject_order.append(col)
                        status_token = extract_status_token(value) or ''
                        is_pass = 'PASS' in status_token
                        year_val = sort_key(attempt)[1]

                        value_str = str(value).strip()
                        entry_subject = None
                        if '-' in value_str:
                            if '/' in value_str and value_str.count('-') == 1:
                                marks_part, status = value_str.split('-', 1)
                                marks_part = marks_part.strip()
                                status = status.strip()
                                if status.upper() == 'ABSENT':
                                    entry_subject = {'subject': col, 'marks': 'Absent', 'status': 'Absent'}
                                else:
                                    formatted_marks = format_marks_as_int(marks_part)
                                    entry_subject = {'subject': col, 'marks': formatted_marks, 'status': status}
                            else:
                                marks, status = value_str.split('-', 1)
                                marks = marks.strip()
                                status = status.strip()
                                if status.upper() == 'ABSENT':
                                    entry_subject = {'subject': col, 'marks': 'Absent', 'status': 'Absent'}
                                else:
                                    formatted_marks = format_marks_as_int(marks)
                                    entry_subject = {'subject': col, 'marks': formatted_marks, 'status': status}
                        else:
                            formatted_marks = format_marks_as_int(value_str)
                            entry_subject = {'subject': col, 'marks': formatted_marks, 'status': 'N/A'}

                        if entry_subject is None:
                            continue

                        existing = subject_choices.get(col)
                        if existing is None:
                            subject_choices[col] = {
                                'data': entry_subject,
                                'is_pass': is_pass,
                                'year': year_val,
                                'order_idx': idx
                            }
                        else:
                            should_replace = False
                            if is_pass and not existing['is_pass']:
                                should_replace = True
                            elif is_pass == existing['is_pass']:
                                if year_val > existing['year']:
                                    should_replace = True
                                elif year_val == existing['year'] and idx > existing['order_idx']:
                                    should_replace = True
                            if should_replace:
                                logger.debug(f"    Replacing subject {col} with newer attempt data")
                                subject_choices[col] = {
                                    'data': entry_subject,
                                    'is_pass': is_pass,
                                    'year': year_val,
                                    'order_idx': idx
                                }

            subjects = [subject_choices[col]['data'] for col in subject_order if col in subject_choices]
            if not subjects:
                logger.debug(f"No subjects selected for semester {display_semester}; skipping.")
                continue

            summary_row = latest_data if isinstance(latest_data, dict) else {}
            statuses_upper = [s['status'].upper() for s in subjects if 'status' in s and isinstance(s['status'], str)]
            derived_result = 'FAIL' if any(status in ('FAIL', 'ABSENT') for status in statuses_upper) else ('PASS' if subjects else 'N/A')

            # Recompute cumulative totals after merging re-exam data
            total_obtained_sum = 0
            total_max_sum = 0
            has_obtained_values = False
            max_missing = False

            for subject in subjects:
                marks_str = subject.get('marks')
                obtained, max_marks = parse_marks(marks_str)
                if obtained is None:
                    continue
                has_obtained_values = True
                total_obtained_sum += obtained
                if max_marks is None:
                    max_missing = True
                else:
                    total_max_sum += max_marks

            total_marks_value = total_obtained_sum if has_obtained_values else None
            max_marks_value = None if max_missing or total_max_sum == 0 else total_max_sum
            percentage_value = None
            if total_marks_value is not None and max_marks_value is not None and max_marks_value > 0:
                percentage_value = round((total_marks_value / max_marks_value) * 100, 2)

            logger.debug(f"Summary row for {display_semester}: {summary_row}")
            summary = {
                'total_marks': str(total_marks_value) if total_marks_value is not None else format_value(summary_row.get('Total Marks')),
                'max_marks': str(max_marks_value) if max_marks_value is not None else format_value(summary_row.get('Max Marks')),
                'percentage': str(percentage_value) if percentage_value is not None else format_value(summary_row.get('Percentage')),
                'cgpa': format_value(summary_row.get('CGPA'), 'Not Available'),
                'sgpa': format_value(summary_row.get('SGPA'), 'Not Available'),
                'result': format_value(summary_row.get('Result'), derived_result)
            }

            formatted_results.append({
                'semester': display_semester,
                'subjects': subjects,
                'summary': summary
            })
        
        return {
            'usn': usn_info['usn'],
            'roll_no': usn_info.get('roll_no'),
            'class': usn_info.get('class'),
            'name': student_name,
            'course': usn_info['course'],
            'year': usn_info['year'],
            'results': formatted_results,
            'cgpa': None  # Will be calculated and set in search_student route
        }
    
    except Exception as e:
        logger.error(f"Error formatting student data: {str(e)}")
        return None

def recalculate_sgpa_cgpa_for_file(file_path):
    """
    Recalculate SGPA and CGPA for all students in a result file.
    Excludes ABSENT subjects from calculations.
    """
    try:
        # Read the result file
        df = pd.read_excel(file_path)
        
        if df.empty:
            return {'success': True, 'updated': 0, 'message': 'File is empty'}
        
        updated_count = 0
        errors = []
        
        # Helper functions for calculations
        def to_float_safe(val):
            """Safely convert value to float"""
            try:
                if isinstance(val, str):
                    v = val.replace('%', '').replace(',', '').strip()
                else:
                    v = val
                return float(v)
            except Exception:
                return None
        
        def parse_marks(marks_str):
            """Parse marks string to extract obtained and max marks as integers"""
            if not marks_str or str(marks_str).strip() == 'Absent' or str(marks_str).strip() == 'N/A':
                return None, None
            try:
                marks_str = str(marks_str).strip()
                # Handle "marks/max-status" or "marks/max" format
                if '/' in marks_str:
                    parts = marks_str.split('/')
                    obtained = to_float_safe(parts[0].strip())
                    max_marks = to_float_safe(parts[1].strip().split('-')[0].strip()) if len(parts) > 1 else None
                    if obtained is not None:
                        obtained = round(obtained)
                    if max_marks is not None:
                        max_marks = round(max_marks)
                    return obtained, max_marks
                else:
                    # Just marks or marks-status
                    if '-' in marks_str:
                        marks_part = marks_str.split('-')[0].strip()
                    else:
                        marks_part = marks_str
                    obtained = to_float_safe(marks_part)
                    if obtained is not None:
                        obtained = round(obtained)
                    return obtained, None
            except Exception:
                return None, None
        
        def percentage_to_sgpa(percentage):
            """Convert percentage to SGPA (10-point scale)
            SGPA is calculated as percentage / 10 for linear conversion
            """
            if percentage is None or percentage < 0:
                return 0.0
            # Linear conversion: SGPA = Percentage / 10
            # Round to 2 decimal places
            return round(percentage / 10.0, 2)
        
        def percentage_to_grade_point(percentage):
            """Convert percentage to grade point (10-point scale with thresholds)
            This is used for individual subject grade points
            """
            if percentage is None:
                return 0.0
            if percentage >= 90:
                return 10.0
            elif percentage >= 80:
                return 9.0
            elif percentage >= 70:
                return 8.0
            elif percentage >= 60:
                return 7.0
            elif percentage >= 50:
                return 6.0
            elif percentage >= 40:
                return 5.0
            else:
                return 0.0
        
        def extract_status(value_str):
            """Extract status (PASS/FAIL/ABSENT) from value string"""
            if not value_str or pd.isna(value_str):
                return None
            value_str = str(value_str).strip().upper()
            if 'ABSENT' in value_str:
                return 'ABSENT'
            elif 'FAIL' in value_str:
                return 'FAIL'
            elif 'PASS' in value_str:
                return 'PASS'
            return None
        
        # Process each student (row) in the file
        for idx, row in df.iterrows():
            try:
                # Get subject columns (columns with ' - ' in name)
                subject_columns = [col for col in df.columns if ' - ' in col and col not in ['Sl. No', 'Name', 'USN', 'Result', 'Total Marks', 'Max Marks', 'Percentage', 'CGPA', 'SGPA']]
                
                # Calculate SGPA for this semester
                subject_grade_points = []
                subject_total_marks = 0.0
                subject_max_marks = 0.0
                has_absent = False
                all_passed = True
                
                for col in subject_columns:
                    value = row[col]
                    if pd.notna(value) and str(value).strip() != 'N/A' and str(value).strip() != '':
                        value_str = str(value).strip()
                        status = extract_status(value_str)
                        
                        # Skip ABSENT subjects
                        if status == 'ABSENT':
                            has_absent = True
                            all_passed = False
                            continue
                        
                        # Check if passed
                        if status != 'PASS':
                            all_passed = False
                        
                        # Parse marks
                        obtained, max_marks = parse_marks(value_str)
                        
                        # Calculate grade point if we have both obtained and max marks
                        if obtained is not None and max_marks is not None and max_marks > 0:
                            percentage = (obtained / max_marks) * 100
                            grade_point = percentage_to_grade_point(percentage)
                            subject_grade_points.append(grade_point)
                            subject_total_marks += obtained
                            subject_max_marks += max_marks
                
                # Calculate SGPA from semester percentage (NOT from individual subject grade points)
                # SGPA = Percentage / 10 (linear conversion)
                calculated_sgpa = None
                
                # Priority 1: Calculate from subject totals (excluding ABSENT subjects)
                if subject_max_marks > 0 and subject_total_marks > 0:
                    percentage = (subject_total_marks / subject_max_marks) * 100
                    calculated_sgpa = percentage_to_sgpa(percentage)
                # Priority 2: Calculate from Excel file totals
                else:
                    total_marks = row.get('Total Marks')
                    max_marks = row.get('Max Marks')
                    if pd.notna(total_marks) and pd.notna(max_marks) and max_marks > 0:
                        try:
                            total = to_float_safe(total_marks)
                            max_val = to_float_safe(max_marks)
                            if total is not None and max_val is not None and max_val > 0:
                                percentage = (total / max_val) * 100
                                calculated_sgpa = percentage_to_sgpa(percentage)
                        except:
                            pass
                    
                    # Priority 3: Calculate from percentage column
                    if calculated_sgpa is None:
                        percentage_val = row.get('Percentage')
                        if pd.notna(percentage_val):
                            try:
                                perc = to_float_safe(percentage_val)
                                if perc is not None:
                                    calculated_sgpa = percentage_to_sgpa(perc)
                            except:
                                pass
                    
                    # Priority 4: Fallback to existing SGPA (but we'll recalculate it)
                    if calculated_sgpa is None:
                        existing_sgpa = row.get('SGPA')
                        if pd.notna(existing_sgpa):
                            try:
                                calculated_sgpa = round(float(existing_sgpa), 2)
                            except:
                                pass
                
                # Update SGPA if calculated
                if calculated_sgpa is not None and 'SGPA' in df.columns:
                    df.at[idx, 'SGPA'] = calculated_sgpa
                    updated_count += 1
                
                # Update Result status
                if 'Result' in df.columns:
                    if has_absent or not all_passed:
                        df.at[idx, 'Result'] = 'FAIL'
                    elif all_passed:
                        df.at[idx, 'Result'] = 'PASS'
                
            except Exception as e:
                error_msg = f"Error processing student at row {idx}: {str(e)}"
                logger.error(error_msg)
                errors.append(error_msg)
        
        # Save the updated file if changes were made
        if updated_count > 0:
            df.to_excel(file_path, index=False)
            logger.info(f"Successfully updated {file_path} with {updated_count} SGPA recalculations")
        
        return {
            'success': len(errors) == 0,
            'updated': updated_count,
            'errors': errors,
            'file': file_path
        }
        
    except Exception as e:
        error_msg = f"Error processing file {file_path}: {str(e)}"
        logger.error(error_msg)
        return {
            'success': False,
            'updated': 0,
            'errors': [error_msg],
            'file': file_path
        }

def recalculate_cgpa_for_all_students():
    """
    Calculate and update CGPA for all students across all semesters.
    This function finds all result files for each student and calculates CGPA
    as the average of all their SGPA values across all semesters.
    """
    try:
        year_folders = get_year_folders()
        total_cgpa_updated = 0
        errors = []
        
        # Step 1: Collect all SGPA values for each student (one per file)
        # Structure: {usn: {'files': [file_paths], 'sgpas': [sgpa_values]}}
        student_data = {}  # {usn: {'files': [file_paths], 'sgpas': [sgpa_values]}}
        
        for year in year_folders:
            year_path = BASE_DIR / year
            if not year_path.exists():
                continue
            
            for course_dir in year_path.iterdir():
                if not course_dir.is_dir():
                    continue
                
                for result_file in course_dir.glob("*_result.xlsx"):
                    try:
                        df = pd.read_excel(result_file)
                        file_path_str = str(result_file)
                        
                        for _, row in df.iterrows():
                            usn = row.get('USN')
                            sgpa = row.get('SGPA')
                            
                            if pd.notna(usn) and pd.notna(sgpa):
                                usn = str(usn).strip().upper()
                                try:
                                    sgpa_val = float(sgpa)
                                    if sgpa_val > 0:  # Only include valid SGPA values
                                        if usn not in student_data:
                                            student_data[usn] = {'files': [], 'sgpas': []}
                                        
                                        # Only add one SGPA per file per student (handle duplicates)
                                        # Check if we already have an SGPA from this file
                                        file_index = None
                                        for i, f in enumerate(student_data[usn]['files']):
                                            if f == file_path_str:
                                                file_index = i
                                                break
                                        
                                        if file_index is None:
                                            # New file, add it
                                            student_data[usn]['files'].append(file_path_str)
                                            student_data[usn]['sgpas'].append(sgpa_val)
                                        else:
                                            # File already exists, update SGPA if different (use latest)
                                            student_data[usn]['sgpas'][file_index] = sgpa_val
                                except:
                                    pass
                    except Exception as e:
                        error_msg = f"Error reading {result_file}: {str(e)}"
                        logger.error(error_msg)
                        errors.append(error_msg)
        
        # Step 2: Calculate CGPA for each student and update all their files
        files_to_update = {}  # {file_path: {usn: cgpa}} to batch updates
        
        for usn, data in student_data.items():
            if len(data['sgpas']) > 0:
                # Calculate CGPA as average of all SGPA values
                cgpa = round(sum(data['sgpas']) / len(data['sgpas']), 2)
                
                # Group updates by file for efficiency
                for file_path in data['files']:
                    if file_path not in files_to_update:
                        files_to_update[file_path] = {}
                    files_to_update[file_path][usn] = cgpa
        
        # Step 3: Update all files with CGPA values
        for file_path, updates in files_to_update.items():
            try:
                df = pd.read_excel(file_path)
                updated = False
                
                for usn, cgpa in updates.items():
                    mask = df['USN'] == usn
                    if mask.any() and 'CGPA' in df.columns:
                        df.loc[mask, 'CGPA'] = cgpa
                        updated = True
                        total_cgpa_updated += 1
                        logger.info(f"Updated CGPA for {usn} in {file_path} to {cgpa}")
                
                if updated:
                    df.to_excel(file_path, index=False)
                    logger.info(f"Saved updated CGPA values to {file_path}")
                    
            except Exception as e:
                error_msg = f"Error updating CGPA in {file_path}: {str(e)}"
                logger.error(error_msg)
                errors.append(error_msg)
        
        return {
            'success': len(errors) == 0,
            'total_students': len(student_data),
            'total_cgpa_updated': total_cgpa_updated,
            'errors': errors,
            'message': f'Updated CGPA for {len(student_data)} students, {total_cgpa_updated} updates made'
        }
        
    except Exception as e:
        logger.error(f"Error in recalculate_cgpa_for_all_students: {str(e)}")
        return {
            'success': False,
            'total_students': 0,
            'total_cgpa_updated': 0,
            'errors': [f"Unexpected error: {str(e)}"],
            'message': f'Error: {str(e)}'
        }

def recalculate_all_sgpa_cgpa():
    """
    Recalculate SGPA and CGPA for all result files across all years.
    This function processes all _result.xlsx files and updates SGPA/CGPA values.
    Step 1: Recalculate SGPA for all files
    Step 2: Calculate and update CGPA for all students across all semesters
    """
    try:
        year_folders = get_year_folders()
        total_updated = 0
        total_files = 0
        errors = []
        processed_files = []
        
        # Step 1: Recalculate SGPA for all files
        logger.info("Step 1: Recalculating SGPA for all files...")
        for year in year_folders:
            year_path = BASE_DIR / year
            if not year_path.exists():
                continue
            
            # Get all result files in this year
            for course_dir in year_path.iterdir():
                if not course_dir.is_dir():
                    continue
                
                # Find all result files
                for result_file in course_dir.glob("*_result.xlsx"):
                    total_files += 1
                    logger.info(f"Processing {result_file}")
                    
                    result = recalculate_sgpa_cgpa_for_file(str(result_file))
                    if result['success']:
                        total_updated += result['updated']
                        processed_files.append(result_file.name)
                    else:
                        errors.extend(result.get('errors', []))
        
        # Step 2: Calculate and update CGPA for all students
        logger.info("Step 2: Calculating CGPA for all students across all semesters...")
        cgpa_result = recalculate_cgpa_for_all_students()
        if not cgpa_result['success']:
            errors.extend(cgpa_result.get('errors', []))
        
        return {
            'success': len(errors) == 0,
            'total_files': total_files,
            'total_sgpa_updated': total_updated,
            'total_cgpa_updated': cgpa_result.get('total_cgpa_updated', 0),
            'total_students': cgpa_result.get('total_students', 0),
            'processed_files': processed_files,
            'errors': errors,
            'message': f'Successfully processed {total_files} files, updated {total_updated} SGPA values and {cgpa_result.get("total_cgpa_updated", 0)} CGPA values for {cgpa_result.get("total_students", 0)} students'
        }
        
    except Exception as e:
        logger.error(f"Error in recalculate_all_sgpa_cgpa: {str(e)}")
        return {
            'success': False,
            'total_files': 0,
            'total_sgpa_updated': 0,
            'total_cgpa_updated': 0,
            'total_students': 0,
            'processed_files': [],
            'errors': [f"Unexpected error: {str(e)}"],
            'message': f'Error: {str(e)}'
        }

def run_excel_processing():
    """Run the excelprocess.py script only on files that need processing"""
    try:
        # Get files that need processing
        status = get_processing_status()
        files_to_process = status['files_needing_processing_list']
        
        if not files_to_process:
            return {
                'success': True,
                'processed_files': [],
                'errors': [],
                'total_processed': 0,
                'message': 'All files are already processed and ready for analysis!'
            }
        
        processed_files = []
        errors = []
        
        # Group files by year for efficient processing
        year_groups = {}
        for file_info in files_to_process:
            year = Path(file_info['path']).parent.parent.name
            if year not in year_groups:
                year_groups[year] = []
            year_groups[year].append(file_info)
        
        # Process each year
        for year, files in year_groups.items():
            year_path = BASE_DIR / year
            if year_path.exists():
                # Run the excelprocess.py script on the year folder
                cmd = [sys.executable, "excelprocess.py", str(year_path)]
                logger.info(f"Running command: {' '.join(cmd)}")
                
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    cwd=BASE_DIR
                )
                
                if result.returncode == 0:
                    # Get the updated list of processed files for this year
                    excel_files = get_excel_files_in_year(year)
                    processed_files.extend([f for f in excel_files if f['result_exists']])
                    logger.info(f"Successfully processed year {year}")
                else:
                    error_msg = f"Error processing year {year}: {result.stderr}"
                    errors.append(error_msg)
                    logger.error(error_msg)
        
        return {
            'success': len(errors) == 0,
            'processed_files': processed_files,
            'errors': errors,
            'total_processed': len(processed_files),
            'message': f'Successfully processed {len(processed_files)} files. All files are now ready for analysis!'
        }
        
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return {
            'success': False,
            'processed_files': [],
            'errors': [f"Unexpected error: {str(e)}"],
            'total_processed': 0
        }

@app.route('/')
def index():
    """Main dashboard page"""
    processing_status = get_processing_status()
    return render_template('dashboard.html', processing_status=processing_status)

@app.route('/admin')
def admin_panel():
    """Admin panel for processing Excel files"""
    year_folders = get_year_folders()
    processing_status = get_processing_status()
    return render_template('admin.html', 
                         year_folders=year_folders, 
                         processing_status=processing_status)

@app.route('/process', methods=['POST'])
def process_excel_files():
    """Process all Excel files"""
    try:
        result = run_excel_processing()
        
        if result['success']:
            flash(f"Successfully processed {result['total_processed']} Excel files!", 'success')
        else:
            flash(f"Processing completed with {len(result['errors'])} errors. Check logs for details.", 'warning')
            for error in result['errors']:
                flash(error, 'error')
        
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"Error in process_excel_files: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e),
            'processed_files': [],
            'errors': [str(e)],
            'total_processed': 0
        }), 500

@app.route('/status')
def get_status():
    """Get current status of Excel files"""
    processing_status = get_processing_status()
    return jsonify(processing_status)

@app.route('/student-analysis')
def student_analysis():
    """Student analysis page"""
    processing_status = get_processing_status()
    return render_template('student_analysis.html', processing_status=processing_status)

@app.route('/search-student', methods=['POST'])
def search_student():
    """Search for student by USN or roll number - Fast direct fetch from processed files"""
    try:
        data = request.get_json()
        identifier = data.get('usn', '').strip()
        logger.info(f"Search request for identifier: {identifier}")
        
        if not identifier:
            return jsonify({'success': False, 'error': 'USN or Roll Number is required'}), 400
        
        # Parse USN or roll number
        usn_info = parse_usn_or_roll(identifier)
        if not usn_info:
            logger.error(f"Invalid identifier format: {identifier}")
            return jsonify({'success': False, 'error': 'Invalid USN or Roll Number format. Expected USN: U05BB23S0037 (12 characters) or Roll Number from students data'}), 400
        
        logger.info(f"Parsed USN info: {usn_info}")
        
        # Check if result files exist (are processed) - fast check without processing
        year = usn_info['year']
        course = usn_info['course']
        year_path = BASE_DIR / year
        course_path = year_path / course
        
        if not course_path.exists():
            return jsonify({'success': False, 'error': f'Result files not found for year {year} and course {course}. Please process the Excel files first.'}), 404
        
        # Check if any result files exist
        result_files = list(course_path.glob("*_result.xlsx"))
        if not result_files:
            return jsonify({'success': False, 'error': f'No processed result files found for year {year} and course {course}. Please process the Excel files first.'}), 404
        
        # Fast direct fetch - no processing, just read from processed files
        # Find student results (primary admission year)
        student_results = find_student_results(usn_info) or []
        def canonical_semester_name(name):
            if not name:
                return str(name).strip()
            match = re.search(r'(sem\s*\d+)', str(name), re.IGNORECASE)
            if match:
                return match.group(1).upper().replace('  ', ' ')
            return str(name).strip().upper()

        allowed_semesters = {canonical_semester_name(res.get('semester')) for res in student_results}

        # Always look for additional attempts in later years (re-exams, improvements)
        all_results = find_all_student_results_across_years(usn_info) or []
        admission_year_raw = usn_info.get('year')
        try:
            admission_year = int(admission_year_raw) if admission_year_raw else None
        except (TypeError, ValueError):
            admission_year = None

        merged_results = list(student_results)
        existing_paths = {res.get('file_path') for res in merged_results if isinstance(res, dict)}

        for result in all_results:
            file_path = result.get('file_path')
            try:
                result_year = int(result.get('year'))
            except (TypeError, ValueError):
                result_year = None

            include_result = False
            if admission_year is None:
                include_result = True
            elif result_year is not None and result_year > admission_year:
                include_result = True

            if include_result and file_path not in existing_paths:
                merged_results.append(result)
                existing_paths.add(file_path)

        if not merged_results and all_results:
            logger.info(f"No admission year results found; using cross-year results for {usn_info['usn']}")
            merged_results = all_results

        if allowed_semesters:
            filtered_results = []
            for res in merged_results:
                canon = canonical_semester_name(res.get('semester'))
                if canon in allowed_semesters:
                    filtered_results.append(res)
            merged_results = filtered_results

        if not merged_results:
            logger.error(f"No results found for USN {usn_info['usn']} in course {usn_info['course']} across all years")
            return jsonify({'success': False, 'error': f'No results found for USN {usn_info["usn"]} in course {usn_info["course"]}. The student may not exist in the processed result files.'}), 404
        
        logger.info(f"Found {len(merged_results)} result files for USN {usn_info['usn']} - fetching data directly (no processing)")
        
        # Format student data - fast direct fetch
        student_data = format_student_data(merged_results, usn_info)
        if not student_data:
            logger.error(f"Error formatting student data for USN {usn_info['usn']}")
            return jsonify({'success': False, 'error': 'Error formatting student data'}), 500
        
        # Get CGPA from most recent semester - fast direct fetch
        calculated_cgpa = calculate_student_cgpa(usn_info)
        if calculated_cgpa is not None:
            student_data['cgpa'] = calculated_cgpa
        else:
            student_data['cgpa'] = None  # Will display as "--" in frontend
        
        # Return data immediately - no processing updates
        response_data = {'success': True, 'student_data': student_data}
        
        logger.info(f"Successfully fetched student data for USN {usn_info['usn']} with CGPA: {student_data.get('cgpa')} (fast direct fetch)")
        return jsonify(response_data)
        
    except Exception as e:
        logger.error(f"Error in search_student: {str(e)}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/check-re-exam-updates', methods=['POST'])
def check_re_exam_updates():
    """Check for available re-examination updates for a student"""
    try:
        data = request.get_json()
        identifier = data.get('usn', '').strip()
        
        if not identifier:
            return jsonify({'success': False, 'error': 'USN or Roll Number is required'}), 400
        
        # Parse USN or roll number
        usn_info = parse_usn_or_roll(identifier)
        if not usn_info:
            return jsonify({'success': False, 'error': 'Invalid USN or Roll Number format'}), 400
        
        # Check for updates without applying them
        updates_needed = check_for_re_examination_updates(usn_info)
        
        if updates_needed:
            return jsonify({
                'success': True,
                'updates_available': len(updates_needed),
                'updates': updates_needed,
                'message': f'Found {len(updates_needed)} subjects that can be updated from FAIL to PASS'
            })
        else:
            return jsonify({
                'success': True,
                'updates_available': 0,
                'updates': [],
                'message': 'No re-examination updates available'
            })
            
    except Exception as e:
        logger.error(f"Error in check_re_exam_updates: {str(e)}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/apply-re-exam-updates', methods=['POST'])
def apply_re_exam_updates():
    """Manually apply re-examination updates for a student"""
    try:
        data = request.get_json()
        identifier = data.get('usn', '').strip()
        
        if not identifier:
            return jsonify({'success': False, 'error': 'USN or Roll Number is required'}), 400
        
        # Parse USN or roll number
        usn_info = parse_usn_or_roll(identifier)
        if not usn_info:
            return jsonify({'success': False, 'error': 'Invalid USN or Roll Number format'}), 400
        
        # Apply updates
        update_result = apply_re_examination_updates(usn_info)
        
        return jsonify(update_result)
        
    except Exception as e:
        logger.error(f"Error in apply_re_exam_updates: {str(e)}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/apply-all-re-exam-updates', methods=['POST'])
def apply_all_re_exam_updates():
    """Apply re-examination updates for all students across all years"""
    try:
        logger.info("Starting bulk re-examination updates for all students")
        
        # Get all year folders
        year_folders = get_year_folders()
        total_updates = 0
        processed_students = 0
        errors = []
        
        # Process each year
        for year in year_folders:
            year_path = BASE_DIR / year
            if not year_path.exists():
                continue
                
            # Get all courses in this year
            for course_dir in year_path.iterdir():
                if not course_dir.is_dir():
                    continue
                    
                course = course_dir.name
                
                # Get all result files in this course
                for result_file in course_dir.glob("*_result.xlsx"):
                    try:
                        # Read the result file to get all USNs
                        df = pd.read_excel(result_file)
                        
                        # Process each student in this file
                        for _, row in df.iterrows():
                            usn = row['USN']
                            if pd.notna(usn) and str(usn).strip():
                                processed_students += 1
                                
                                # Parse USN info
                                usn_info = parse_usn_or_roll(str(usn).strip())
                                if usn_info:
                                    # Apply updates for this student
                                    update_result = apply_re_examination_updates(usn_info)
                                    if update_result['success']:
                                        total_updates += update_result['updates_applied']
                                    else:
                                        errors.extend(update_result.get('errors', []))
                                        
                    except Exception as e:
                        error_msg = f"Error processing {result_file}: {str(e)}"
                        logger.error(error_msg)
                        errors.append(error_msg)
        
        logger.info(f"Bulk update completed. Processed {processed_students} students, applied {total_updates} updates")
        
        return jsonify({
            'success': len(errors) == 0,
            'total_updates_applied': total_updates,
            'processed_students': processed_students,
            'errors': errors,
            'message': f'Successfully applied {total_updates} re-examination updates for {processed_students} students'
        })
        
    except Exception as e:
        logger.error(f"Error in apply_all_re_exam_updates: {str(e)}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/recalculate-sgpa-cgpa', methods=['POST'])
def recalculate_sgpa_cgpa_endpoint():
    """Recalculate SGPA and CGPA for all result files"""
    try:
        logger.info("Starting SGPA/CGPA recalculation for all files")
        result = recalculate_all_sgpa_cgpa()
        
        if result['success']:
            flash(f"Successfully recalculated SGPA for {result['total_updated']} students across {result['total_files']} files!", 'success')
        else:
            flash(f"Recalculation completed with {len(result['errors'])} errors. Check logs for details.", 'warning')
            for error in result['errors']:
                flash(error, 'error')
        
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"Error in recalculate_sgpa_cgpa_endpoint: {str(e)}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/update-all-re-exam-and-recalculate', methods=['POST'])
def update_all_re_exam_and_recalculate():
    """Apply re-examination updates for all students and then recalculate SGPA/CGPA"""
    try:
        logger.info("Starting bulk re-examination updates and SGPA/CGPA recalculation")
        
        # Step 1: Apply re-examination updates for all students
        year_folders = get_year_folders()
        total_updates = 0
        processed_students = 0
        errors = []
        
        # Process each year
        for year in year_folders:
            year_path = BASE_DIR / year
            if not year_path.exists():
                continue
                
            # Get all courses in this year
            for course_dir in year_path.iterdir():
                if not course_dir.is_dir():
                    continue
                    
                course = course_dir.name
                
                # Get all result files in this course
                for result_file in course_dir.glob("*_result.xlsx"):
                    try:
                        # Read the result file to get all USNs
                        df = pd.read_excel(result_file)
                        
                        # Process each student in this file
                        for _, row in df.iterrows():
                            usn = row['USN']
                            if pd.notna(usn) and str(usn).strip():
                                processed_students += 1
                                
                                # Parse USN info
                                usn_info = parse_usn_or_roll(str(usn).strip())
                                if usn_info:
                                    # Apply updates for this student
                                    update_result = apply_re_examination_updates(usn_info)
                                    if update_result['success']:
                                        total_updates += update_result['updates_applied']
                                    else:
                                        errors.extend(update_result.get('errors', []))
                                        
                    except Exception as e:
                        error_msg = f"Error processing {result_file}: {str(e)}"
                        logger.error(error_msg)
                        errors.append(error_msg)
        
        logger.info(f"Applied {total_updates} re-examination updates for {processed_students} students")
        
        # Step 2: Recalculate SGPA/CGPA for all files
        logger.info("Starting SGPA/CGPA recalculation after re-examination updates")
        recalculation_result = recalculate_all_sgpa_cgpa()
        
        # Combine results
        return jsonify({
            'success': len(errors) == 0 and recalculation_result['success'],
            're_exam_updates': {
                'total_updates_applied': total_updates,
                'processed_students': processed_students,
                'errors': errors
            },
            'recalculation': {
                'total_files': recalculation_result['total_files'],
                'total_updated': recalculation_result['total_updated'],
                'processed_files': recalculation_result['processed_files'],
                'errors': recalculation_result['errors']
            },
            'message': f'Applied {total_updates} re-examination updates and recalculated SGPA for {recalculation_result["total_updated"]} students across {recalculation_result["total_files"]} files'
        })
        
    except Exception as e:
        logger.error(f"Error in update_all_re_exam_and_recalculate: {str(e)}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)