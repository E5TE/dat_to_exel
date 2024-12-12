# DAT to Excel Converter

This Python script converts .dat files to Excel (.xlsx) format. It's designed to be flexible and handle various data formats commonly found in .dat files.

## Features

- Supports multiple file encodings (UTF-8, Latin1, CP1252, ISO-8859-1)
- Handles different data formats:
  - Delimited text (tabs, commas, pipes, semicolons)
  - Fixed-width format
  - Binary data (float values)
- Automatically processes all .dat files in the current directory
- Creates Excel files with the same name as the source .dat file

## Requirements

- Python 3.x
- pandas library
- openpyxl library (for Excel file creation)

## Installation

1. Ensure Python 3.x is installed on your system
2. Install required dependencies:
```
pip install pandas openpyxl
```

## Usage

1. Place your .dat files in the same directory as the script
2. Run the script:
```
python convert_dat_to_excel.py
```

The script will:
- Automatically detect and process all .dat files in the directory
- Create corresponding .xlsx files with the same name
- Print progress messages during conversion

## How It Works

The script employs a multi-step approach to handle various data formats:

1. Attempts to read the file using different text encodings
2. Tests various delimiters (tab, comma, pipe, semicolon)
3. Tries fixed-width format parsing
4. Falls back to binary data interpretation if text-based approaches fail

If any method succeeds, the data is converted to an Excel file. If all methods fail, an error message is displayed.

## Error Handling

- The script provides informative error messages if conversion fails
- Each file is processed independently, so an error with one file won't affect others
- Invalid or corrupted .dat files are reported with appropriate error messages

## Output

- Excel files are created in the same directory as the source .dat files
- The output filename matches the input filename with .xlsx extension
- Data is preserved without row indices in the Excel output
