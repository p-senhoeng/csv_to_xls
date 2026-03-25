# CSV to XLS Converter

A simple desktop app that converts **CSV files** into real **Excel 97-2003 (.xls)** files.

## Features

- Select and convert multiple CSV files at once
- Outputs real `.xls` (Excel 97-2003) format
- Keeps the original filename
- Saves converted files to a `Converted/` subfolder next to your originals
- Originals are never modified or deleted
- Simple UI with progress tracking

## Requirements

- Python 3.8+

## Installation

```bash
git clone https://github.com/p-senhoeng/csv_to_xls.git
cd csv_to_xls
pip install pandas xlwt
```

## Usage

```bash
python converter.py
```

1. Click **Browse** and select one or more `.csv` files
2. Click **Convert All**
3. Converted files will appear in a `Converted/` folder next to your original files

## Output

```
Downloads/
├── report.csv          ← original (untouched)
└── Converted/
    └── report.xls      ← real Excel file
```

## Building an Executable (Windows)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "CSV to XLS Converter" converter.py
```

The executable will be at `dist/CSV to XLS Converter.exe`.

## Building an Executable (Mac)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "CSV to XLS Converter" converter.py
```

The app will be at `dist/CSV to XLS Converter`.

## Dependencies

| Package | Purpose |
|---|---|
| `pandas` | Reads CSV files |
| `xlwt` | Writes real Excel 97-2003 (.xls) files |
