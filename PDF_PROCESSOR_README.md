# PDF Order Processor System

This system processes PDF order sheets and updates a master Excel file with extracted information.

## Features

- **Fixed Excel File**: Uses a single master Excel file that persists between runs
- **Multiple PDF Processing**: Can process individual PDFs or entire directories
- **Duplicate Prevention**: Checks for existing invoice numbers to avoid duplicates
- **Cell Formatting**: Automatically centers all data in Excel cells
- **Archiving**: Option to move processed PDFs to a separate folder
- **Command-line Interface**: Flexible options for different usage scenarios

## Quick Start

1. **Run the batch file** to process all PDFs in the current directory:
   ```
   process_orders.bat
   ```

2. **Or run the Python script directly** with more options:
   ```
   python production_pdf_processor.py [options]
   ```

## Command-line Options

- `--excel FILE`: Specify the master Excel file (default: dispatch_master.xlsx)
- `--pdf FILE`: Process a single PDF file
- `--dir DIRECTORY`: Process all PDFs in a specific directory
- `--pattern PATTERN`: Specify file pattern (default: *.pdf)

## Examples

```bash
# Process a single PDF file
python production_pdf_processor.py --pdf order123.pdf

# Process all PDFs in a specific directory
python production_pdf_processor.py --dir incoming_orders

# Use a different Excel file
python production_pdf_processor.py --excel orders_2025.xlsx

# Process only PDFs with specific naming pattern
python production_pdf_processor.py --pattern "order_*.pdf"
```

## Excel File Structure

The master Excel file has the following columns:
- Date
- Invoice Number
- PO Number
- Company Name
- Pick/Delivery
- Pick up number-Time
- Pallets
- Done

## Requirements

- Python 3.6+
- Required Python packages:
  - PyMuPDF (fitz)
  - pytesseract
  - Pillow
  - pandas
  - openpyxl

## Installation

1. Install Python from [python.org](https://www.python.org/downloads/)
2. Install Tesseract OCR from [github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)
3. Install required Python packages:
   ```
   pip install PyMuPDF pytesseract Pillow pandas openpyxl
   ```

## How It Works

1. **PDF Processing**:
   - Converts PDF pages to high-resolution images
   - Uses OCR to extract text from images
   - Applies pattern matching to find specific data fields

2. **Excel Management**:
   - Creates master Excel file if it doesn't exist
   - Checks for duplicate invoice numbers
   - Appends new rows with extracted data
   - Applies center alignment to all cells

3. **Workflow**:
   - Place PDF order sheets in the designated folder
   - Run the processor
   - New orders are added to the Excel file
   - Optionally, processed PDFs are moved to an archive folder
