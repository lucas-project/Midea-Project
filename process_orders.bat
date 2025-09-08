@echo off
echo PDF Order Processor
echo =================
echo.

REM Check if Python is installed
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found. Please install Python.
    pause
    exit /b
)

REM Set the Excel file name
set EXCEL_FILE="Dispatch Schedule.xlsx"

REM Process PDF file
if "%~1"=="" (
    echo No PDF file specified. Processing all PDFs in current directory...
    for %%f in (*.pdf) do (
        echo Processing: %%f
        python improved_pdf_processor.py --pdf "%%f" --excel %EXCEL_FILE%
    )
) else (
    echo Processing PDF: %1
    python improved_pdf_processor.py --pdf "%~1" --excel %EXCEL_FILE%
)

echo.
echo Processing complete!
pause