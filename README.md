# Student File Download Script

This repository provides a Python script to automate the downloading of student files from an Excel spreadsheet. The script is designed for teachers, administrators, or anyone who needs to batch-download files submitted by students, such as assignments, reports, or other resources.
本项目服务于复旦大学“步青计划”

## Background

In many educational settings, students submit their work via online forms or spreadsheets, resulting in an Excel file with student names and download links. Manually downloading each file is tedious and error-prone. This script streamlines the process by:

- Automatically detecting download columns (containing "下载") in the Excel file.
- Downloading each file and naming it according to the student and assignment.
- Organizing all files in a dedicated `downloads/` folder.
- Handling errors, retries, and filename sanitization for maximum compatibility.

## Features

- **Automatic Excel Parsing:** Detects student name columns (e.g., "姓名", "学生姓名", "name") and download link columns (headers containing "下载").
- **Batch Downloading:** Downloads all files referenced in the spreadsheet, supporting various file types (PDF, ZIP, DOC, XLSX, etc.).
- **Custom File Naming:** Files are saved as `StudentName_ColumnName.extension` for easy identification.
- **Robust Error Handling:** Skips invalid or missing URLs, retries failed downloads up to 3 times, and sanitizes filenames for all operating systems.
- **Progress Reporting:** Prints real-time progress and a summary of download results.
- **Organized Output:** All files are saved in a `downloads/` directory for easy access.

## Getting Started

### 1. Clone the Repository

```bash
git clone <repository-url>
cd FileDownload
```

### 2. Install Dependencies

Make sure you have Python 3.7+ installed, then install the required packages:

```bash
pip install -r requirements.txt
```

Or install packages individually:

```bash
pip install pandas requests openpyxl pathlib2
```

### 3. Prepare Your Excel File

Ensure your Excel file is named `StudentListAndCorrespondingWebLinks.xlsx` and place it in the same directory as the script.

### 4. Run the Script

```bash
python download_student_files.py
```

## Excel File Requirements

### Expected File Structure

Your Excel file should contain:

1. **Student Name Column**: One of the following column headers:
   - `姓名` (Name)
   - `学生姓名` (Student Name)
   - `学生` (Student)
   - `name` or `Name`
   - `学员姓名` (Trainee Name)
   - `1、你的姓名：` (1. Your Name:)
   - Or the first column if none of the above are found

2. **Download Columns**: Any column header containing "下载" (download)

### Example Excel Structure

| 姓名 | 作业下载 | 报告下载 | 项目文件下载 |
|------|----------|----------|--------------|
| 张三 | https://example.com/zhang_homework.pdf | https://example.com/zhang_report.docx | https://example.com/zhang_project.zip |
| 李四 | https://example.com/li_homework.pdf | https://example.com/li_report.docx | https://example.com/li_project.zip |
| 王五 | https://example.com/wang_homework.pdf | https://example.com/wang_report.docx | https://example.com/wang_project.zip |

## Output

### File Naming Convention

Downloaded files are automatically renamed using the format:
```
StudentName_ColumnName.extension
```

For example:
- Student: "张三", Column: "作业下载", Original: "homework.pdf"
- Result: `张三_作业.pdf`

### Directory Structure

```
project-folder/
├── download_student_files.py
├── StudentListAndCorrespondingWebLinks.xlsx
├── requirements.txt
├── README.md
└── downloads/
    ├── 张三_作业.pdf
    ├── 张三_报告.docx
    ├── 张三_项目文件.zip
    ├── 李四_作业.pdf
    ├── 李四_报告.docx
    └── ...
```

## Usage Examples

### Basic Usage

```bash
python download_student_files.py
```

### Sample Output

```
Successfully loaded Excel file with 50 rows and 4 columns
Columns: ['姓名', '作业下载', '报告下载', '项目文件下载']
Found 3 download columns: ['作业下载', '报告下载', '项目文件下载']
Using student name column: 姓名

Processing student: 张三
  Downloading from 作业下载: https://example.com/homework1.pdf
Downloading: 张三_作业 (attempt 1)
Successfully downloaded: downloads/张三_作业.pdf
  Downloading from 报告下载: https://example.com/report1.docx
Downloading: 张三_报告 (attempt 1)
Successfully downloaded: downloads/张三_报告.docx

Processing student: 李四
  Downloading from 作业下载: https://example.com/homework2.pdf
Downloading: 李四_作业 (attempt 1)
Successfully downloaded: downloads/李四_作业.pdf
  Skipping 报告下载: No URL provided

==================================================
Download Summary:
Total download attempts: 45
Successful downloads: 43
Failed downloads: 2
==================================================
```

## Error Handling

The script includes comprehensive error handling for common issues:

### Automatic Handling

- **Missing URLs**: Empty cells are automatically skipped
- **Invalid URLs**: Non-HTTP/HTTPS URLs are skipped with a warning
- **Network Errors**: Failed downloads are retried up to 3 times
- **Filename Issues**: Special characters are sanitized for cross-platform compatibility
- **Missing Extensions**: File extensions are detected from URLs or HTTP headers

### Common Issues and Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| "No columns containing '下载' found!" | Column headers don't contain "下载" | Ensure download columns have "下载" in their names |
| "Could not find StudentListAndCorrespondingWebLinks.xlsx" | Excel file not found | Place the Excel file in the script directory |
| "Invalid URL format" | URL doesn't start with http/https | Check that URLs are properly formatted |
| Download timeouts | Slow network or large files | Script will automatically retry |

## Configuration

### Customizing Student Name Detection

If your Excel uses different column names for student names, you can modify the `possible_name_columns` list in the script:

```python
possible_name_columns = ["姓名", "学生姓名", "学生", "name", "Name", "学员姓名", "1、你的姓名：", "Your Custom Column"]
```

### Adjusting Download Settings

You can modify these settings in the `download_file()` function:

- **Timeout**: Change `timeout=30` (seconds)
- **Retries**: Change `max_retries=3`
- **Delay**: Change `time.sleep(1)` (seconds between downloads)

## System Requirements

- **Python**: 3.7 or higher
- **Operating System**: Windows, macOS, or Linux
- **Memory**: Minimal (suitable for hundreds of files)
- **Storage**: Ensure adequate space for downloaded files

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| pandas | ≥1.3.0 | Excel file reading and data manipulation |
| requests | ≥2.25.0 | HTTP requests for file downloading |
| openpyxl | ≥3.0.0 | Excel file format support |
| pathlib2 | ≥2.3.0 | Cross-platform file path handling |

## Troubleshooting

### Permission Issues

If you encounter permission errors:

```bash
# On Windows (run as administrator)
python download_student_files.py

# On macOS/Linux
sudo python download_student_files.py
```

### Large File Downloads

For very large files or slow connections:
- Increase the timeout value in the script
- Consider running during off-peak hours
- Check available disk space

### Character Encoding Issues

If you see garbled characters in filenames:
- Ensure your Excel file is saved with UTF-8 encoding
- Check that your terminal supports Unicode characters

## Contributing

We welcome contributions! Please feel free to:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/improvement`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If you encounter any issues or have questions:

1. Check the [Troubleshooting](#troubleshooting) section
2. Review the [Issues](../../issues) page for similar problems
3. Create a new issue with detailed information about your problem

## Changelog

### v1.0.0
- Initial release
- Basic Excel parsing and file downloading
- Automatic file naming and organization
- Error handling and retry mechanism
