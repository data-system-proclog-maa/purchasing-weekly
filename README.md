# purchasing-weekly
This project automates the weekly data processing workflow for Purchasing Weekly Report Division MAA Group, converting raw PO and RFM data into structured weekly reports.

## Folder Structure
```
PURCHASING-WEEKLY/
├── src/                       # SOURCE CODE ROOT
│   ├── app/                   # APPLICATION PACKAGE
│   │   ├── __init__.py          # Marks 'app' as a Python package
│   │   ├── loader.py            # Data ingestion and file I/O
│   │   ├── localization.py      # Department mapping and business rules
│   │   ├── main.py              # Application entry point & CLI
│   │   ├── processor.py         # Core data transformation logic
│   │   └── styler.py            # Excel styling and formatting
│   └── legacy/                # Archive for old scripts/notebooks
│       └── procurement_processor_updated copy.ipynb # Original notebook
├── requirements.txt           # Python dependencies
├── walkthrough.md             # Usage guide
└── README.md                  # Project documentation
```

## Architecture Overview
The application is split into five main modules:
| Module | Responsibility | 
| ------------- | ------------- |
| `loader.py` | Data ingestion. Handles loading Excel files and Google Sheets data with robust error handling. |
| `localization.py` | Business Logic. Centralizes department mappings (`OBI`, `LAR`, `HO`) and splitting logic. |
| `processor.py` | Core Transformation. Filters data, applies normalization, calculates dates, and prepares datasets for reporting. |
| `styler.py` | Presentation Layer. Applies specific Excel formatting, coloring, and formula injection (`PR-PO` calculation). |
| `main.py` | Entry Point. Orchestrates the workflow, supporting both **Interactive Mode** and **Command Line Arguments**. |

## Data Processing Highlight
The pipeline performs the following key operations:
+ **Dynamic Normalization**: Pulls live normalization data from Google Sheets to enrich local Excel data.
+ **Department Assignment**: Automatically assigns departments based on Procurement Name using predefined rules in `localization.py`.
+ **Date Logic**: Handles complex date parsing and filtering for weekly reporting periods.
+ **Excel Styling**: Automates the formatting of output files, including color-coding tabs and highlighting key columns.

## Prerequisites
Python (3.10+) with the following libraries installed:
```bash
pip install -r requirements.txt
```

## Usage
### Interactive Mode
Simply run the script to be prompted for inputs:
```bash
python src/app/main.py
```

### Command Line Mode
For automation:
```bash
python src/app/main.py \
  --po-file "/path/to/PO.xlsx" \
  --rfm-file "/path/to/RFM.xlsx" \
  --start-date "14-11-2025" \
  --end-date "21-11-2025" \
  --output-dir "/path/to/output"
```
