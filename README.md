# Student File Download Script

This repository provides a Python script to automate the downloading of student files from an Excel spreadsheet. The script is designed for teachers, administrators, or anyone who needs to batch-download files submitted by students, such as assignments, reports, or other resources.

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