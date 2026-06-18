# Olympiad Smart Bulk Downloader

A simple Streamlit application to bulk download Olympiad question papers and solution PDFs from an Excel file.

This app allows users to upload an Excel file containing multiple subject sheets and URLs, select the required subjects, exams, and file types, then download all selected files in one organized ZIP file.

## Features

- Upload Excel file with question paper and solution URLs
- Automatically detect URLs from multiple sheets
- Supports subject-wise filtering
- Supports exam-wise filtering
- Supports question paper and solution filtering
- Downloads all selected files in one ZIP
- Organizes downloaded files by subject and exam
- Generates a download report in Excel format
- Optional merged PDF creation by subject
- Optional PDF text extraction

## Supported Input File Format

The app works best with Excel files containing subject sheets such as:

```text
CHEMISTRY
PHYSICS
MATHEMATICS