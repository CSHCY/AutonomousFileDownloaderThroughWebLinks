import pandas as pd
import requests
import os
import re
from urllib.parse import urlparse
from pathlib import Path
import time

def sanitize_filename(filename):
    """
    Sanitize filename to remove invalid characters for file systems
    """
    # Remove or replace invalid characters
    invalid_chars = r'[<>:"/\\|?*]'
    filename = re.sub(invalid_chars, '_', filename)
    # Remove extra spaces and dots
    filename = re.sub(r'\s+', ' ', filename).strip()
    filename = filename.strip('.')
    return filename

def get_file_extension_from_url(url):
    """
    Try to get file extension from URL
    """
    parsed_url = urlparse(url)
    path = parsed_url.path
    if '.' in path:
        return os.path.splitext(path)[1]
    return ''

def download_file(url, filename, max_retries=3):
    """
    Download file from URL with retry mechanism
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    for attempt in range(max_retries):
        try:
            print(f"Downloading: {filename} (attempt {attempt + 1})")
            
            response = requests.get(url, headers=headers, timeout=30, stream=True)
            response.raise_for_status()
            
            # Get file extension from URL or Content-Type
            file_extension = get_file_extension_from_url(url)
            if not file_extension:
                content_type = response.headers.get('Content-Type', '')
                if 'pdf' in content_type.lower():
                    file_extension = '.pdf'
                elif 'zip' in content_type.lower():
                    file_extension = '.zip'
                elif 'doc' in content_type.lower():
                    file_extension = '.doc'
                elif 'excel' in content_type.lower() or 'spreadsheet' in content_type.lower():
                    file_extension = '.xlsx'
            
            # Add extension if not present
            if file_extension and not filename.endswith(file_extension):
                filename += file_extension
            
            # Create downloads directory if it doesn't exist
            downloads_dir = Path("downloads")
            downloads_dir.mkdir(exist_ok=True)
            
            filepath = downloads_dir / filename
            
            # Download the file
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            print(f"Successfully downloaded: {filepath}")
            return True
            
        except requests.exceptions.RequestException as e:
            print(f"Error downloading {filename} (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                print("Retrying...")
                time.sleep(2)
        except Exception as e:
            print(f"Unexpected error downloading {filename}: {e}")
            break
    
    print(f"Failed to download: {filename}")
    return False

def main():
    # Read the Excel file
    try:
        excel_file = "StudentListAndCorrespondingWebLinks.xlsx"
        df = pd.read_excel(excel_file)
        print(f"Successfully loaded Excel file with {len(df)} rows and {len(df.columns)} columns")
        print(f"Columns: {list(df.columns)}")
        
    except FileNotFoundError:
        print(f"Error: Could not find {excel_file}")
        return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Find columns containing "下载"
    download_columns = [col for col in df.columns if "下载" in str(col)]
    print(f"Found {len(download_columns)} download columns: {download_columns}")
    
    if not download_columns:
        print("No columns containing '下载' found!")
        return
    
    # Find the student name column (usually the first column or one containing "姓名", "学生", "name")
    student_name_column = None
    possible_name_columns = ["姓名", "学生姓名", "学生", "name", "Name", "学员姓名","1、你的姓名："]
    
    for col_name in possible_name_columns:
        if col_name in df.columns:
            student_name_column = col_name
            break
    
    # If no specific name column found, use the first column
    if student_name_column is None:
        student_name_column = df.columns[0]
        print(f"Using first column as student name column: {student_name_column}")
    else:
        print(f"Using student name column: {student_name_column}")
    
    # Statistics
    total_downloads = 0
    successful_downloads = 0
    
    # Process each student
    for index, row in df.iterrows():
        student_name = str(row[student_name_column]).strip()
        
        if pd.isna(student_name) or student_name == "" or student_name.lower() == "nan":
            print(f"Skipping row {index + 1}: No student name")
            continue
        
        print(f"\nProcessing student: {student_name}")
        
        # Process each download column for this student
        for col_name in download_columns:
            download_url = row[col_name]
            
            # Skip if URL is empty or NaN
            if pd.isna(download_url) or str(download_url).strip() == "" or str(download_url).lower() == "nan":
                print(f"  Skipping {col_name}: No URL provided")
                continue
            
            download_url = str(download_url).strip()
            
            # Basic URL validation
            if not (download_url.startswith('http://') or download_url.startswith('https://')):
                print(f"  Skipping {col_name}: Invalid URL format: {download_url}")
                continue
            
            # Create filename: student_name + column_name
            column_clean = col_name.replace("下载", "").strip()
            if not column_clean:
                column_clean = "文件"
            
            filename = f"{sanitize_filename(student_name)}_{sanitize_filename(column_clean)}"
            
            print(f"  Downloading from {col_name}: {download_url}")
            
            total_downloads += 1
            if download_file(download_url, filename):
                successful_downloads += 1
            
            # Small delay between downloads to be respectful
            time.sleep(1)
    
    # Print summary
    print(f"\n{'='*50}")
    print(f"Download Summary:")
    print(f"Total download attempts: {total_downloads}")
    print(f"Successful downloads: {successful_downloads}")
    print(f"Failed downloads: {total_downloads - successful_downloads}")
    print(f"{'='*50}")

if __name__ == "__main__":
    main() 