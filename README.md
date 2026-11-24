# Excel Sales Report Generator

A Python tool that combines monthly retail sales Excel files, filters transactions by an optional date range, cleans malformed rows, and exports a styled multi-sheet workbook (`consolidated_sales_report.xlsx`) with product, category, and overall summaries.

## Features

- **Multi-File Consolidation:** Reads `.xlsx` and `.xls` files from a target folder (e.g., `sales/january.xlsx` through `april.xlsx`) and combines records with source tracking.
- **Data Cleaning & Validation:** Safely handles missing fields, invalid dates, negative/non-numeric prices or quantities, and inconsistent text casing.
- **Date Range Filtering:** Optionally filters transactions between a user-specified start and end date.
- **Multi-Sheet Styled Report:** Generates `All Transactions`, `Product Summary` (ranked by quantity sold), `Category Summary`, and `Overall Summary` sheets with formatted headers and auto-fitted column widths.

## Tech Stack

- **Python 3** (`pathlib`)
- **pandas**
- **openpyxl**

## Getting Started

### Prerequisites

- Python 3.9 or higher with `pandas` and `openpyxl` installed.

### Installation & Usage

1. Clone or download the repository and install dependencies:

   ```bash
   git clone <your-repo-url>
   cd <your-repo-folder>
   pip install pandas openpyxl

2. Run

    python excel_report_generator.py