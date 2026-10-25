# PDF Invoice Automation

A Python starter project that extracts labelled fields from **text-based PDF invoices** and writes them to a formatted Excel report.

> **Important:** `examples/INV-1001.pdf` contains fictional demonstration data only. This project is a starter automation, not accounting or tax software. Review extracted values against each source invoice before relying on them.

## Features

- Extracts invoice number, customer, line-item description, and total amount.
- Processes PDF files in batches and continues after individual file errors.
- Flags missing fields and duplicate invoice numbers for manual review.
- Creates a formatted Excel report with filters, frozen headers, currency formatting, and review highlighting.
- Saves extracted text for troubleshooting.
- Writes logs to both the console and a rotating log file.
- Uses paths relative to the project directory, so it can be launched from PyCharm or a terminal.

## Project structure

```text
pdf-invoice-automation/
├── main.py
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
├── input/
│   └── .gitkeep
├── output/
│   ├── .gitkeep
│   └── text/
│       └── .gitkeep
├── examples/
│   └── INV-1001.pdf
└── tests/
    └── test_extractors.py
```

## Requirements

- Python 3.10 or later
- `pypdf`
- `openpyxl`

## Setup on Windows (PowerShell)

Open a terminal in this project folder and run:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If PowerShell prevents activation, you can skip activation and use `.\.venv\Scripts\python.exe` in place of `python` for the commands below.

## Run the sample invoice

Copy the synthetic example into the input folder:

```powershell
Copy-Item .\examples\INV-1001.pdf .\input\INV-1001.pdf
python .\main.py
```

Expected output files:

- `output/extracted_invoices.xlsx` — report for the current batch
- `output/text/INV-1001.txt` — extracted text for review
- `output/automation.log` — execution log

Close the Excel workbook before rerunning, because Excel may lock the report file.

## Run tests

```powershell
python -m unittest discover -s tests -v
```

## Extraction limitations

The extraction rules expect common labels such as `Invoice Number`, `Customer`, `Description Amount`, and `Total Amount`. Different vendors use different layouts; validate and adapt the rules to each template. Scanned/image-only PDFs need an OCR stage, which is not included. Duplicate detection only checks invoice numbers within the current run. The program flags missing data, but it does not verify tax compliance, payment status, or invoice arithmetic.

## Privacy and GitHub hygiene

- Never commit real client invoices, extracted records, logs, credentials, or other confidential information.
- `input/*.pdf` and generated `output/` files are excluded by `.gitignore`.
- The fictional test PDF is intentionally kept in `examples/` for reproducible testing.
- Confirm `git status` before every push.

## License

See [LICENSE](LICENSE).
