# PDF Keyword Highlight Tool

Python automation tool for PDF keyword search, highlight keyword, and file processing.

## Overview

This tool automatically searches for keywords in PDF files, highlights them with a specified color, moves processed files to a destination folder, and logs all results to a single shared Excel file.

Built as a business automation solution to reduce manual PDF review work and improve accuracy.

## Features

- Keyword search and highlight in PDF files
- Excel-driven configuration — no coding required for users
- Modern GUI with dark Charcoal + Mint theme
- **Single folder mode** — process one customer at a time
- **All customers mode** — process all customers from Excel in one run
- **Overwrite option** — reprocess files when settings change
- **Open Log button** — open Excel log directly from the tool
- Real-time color-coded processing log
- Summary stats (Completed / Skipped / Errors)
- Progress bar
- Shared log file — all customers' results in one Excel file
- Excel auto-reload on every Run — no restart needed after config changes
- Never overwrites existing files unless Overwrite is enabled

## Technologies Used

- Python 3.10+
- PyMuPDF (fitz) — PDF processing and annotation
- openpyxl — Excel read/write
- tkinter — GUI interface

## File Structure

```
C:\PDFTool\                         Tool files
    highlight_gui_v2.py             Main GUI application
    processor.py                    PDF processing engine
    excel_master.py                 Excel configuration reader
    utils.py                        Shared utility functions
    run_highlight.bat               Windows launcher
    highlight_master_v3.xlsx        Configuration template
    requirements.txt                Required libraries

C:\PDFs\                            PDF working folder
    pdf_highlight_log.xlsx          Shared log (all customers)
    CustomerA\
        Input\                      Place PDF files here
        Done\                       Highlighted PDFs moved here
    CustomerB\
        Input\
        Done\
    CustomerC\
        Input\
        Done\
```

## Configuration (highlight_master_v3.xlsx)

| Column | Required | Description |
|--------|----------|-------------|
| Customer | Yes | Customer name |
| CustomerFolder | Yes | Path to customer folder |
| DestinationFolder | Yes | Output folder for highlighted PDFs |
| Keywords | Yes | Keywords to search (comma separated) |
| CaseSensitive | No | TRUE / FALSE (default: FALSE) |
| WholeWord | No | TRUE / FALSE (default: FALSE) |
| Color | No | Blue / Yellow / Green / Pink / Orange / Purple / Red |
| Begin_Page | No | First page to search (default: 1) |
| End_Page | No | Last page to search (blank = all pages) |

## Processing Modes

### Single folder mode
Select one customer's Input folder manually.
The tool automatically finds the matching rule in Excel.

### All customers mode
Process all customers listed in Excel in a single run.
Each customer's `CustomerFolder\Input\` is used automatically.

## Status Log

| Status | Meaning |
|--------|---------|
| OK_HITS_n | Processed — n keyword matches found |
| OK_HITS_0 | Processed — no keywords found |
| SKIPPED_EXISTS | File already exists in destination |
| ERROR | Processing error |
| CANCELLED | Stopped by user |

## Requirements

```
pymupdf==1.24.9
openpyxl==3.1.5
```

Install:

```
python -m pip install -r requirements.txt
```

## Author

Kaori Kashiwagi  
Business Systems & Automation Analyst
