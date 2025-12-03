# Inefficient VMs' Report Beautification Script

## Overview
This Python script transforms Prism Central's CSV report of inefficient VMs into a clean, structured XLSX workbook.
The output makes it easier to analyze VM efficiency and resource usage across clusters.

## Features
- Generates an organized XLSX workbook with each cluster on a separate sheet
- Each sheet includes three clearly separated sections with VM counts:
    - Overprovisioned VMs
    - Overprovisioned and/or Constrained VMs
    - Inactive VMs
- Improves readability and accessibility for administrators

## Requirements
- Python 3.8+
- Python packages:
    - `pandas`
    - `openpyxl`
- Last tested on Inefficient VMs' CSV Report downloaded from Prism Central version 7.3.0.5

Install dependencies with:

```bash
pip install pandas openpyxl
```

## Usage

1. Clone the repo:
```bash
git clone https://github.com/harshvardhan26/ntnx-scripts.git
```
2. Navigate to the directory:
```bash
cd ntnx-scripts/inefficient-vms-report-beautification
```
3. Examine the files in the **/input** and **/output** directories to see examples of the input CSV and the generated XLSX output
4. Downoad/Export the inefficient VMs' CSV report from Prism Central and place it under the **/input** directory
5. Rename the downloaded CSV report to **inefficient_vms_report.csv**
6. Run the script:
```bash
python inefficient_vms_report_beautification.py
```
7. The output XLSX workbook **inefficient_vms_report_beautified.xlsx** is saved under the **/output** directory