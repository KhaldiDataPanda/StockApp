# Stock Reconciliation App - Backend

This folder contains the Python backend processor for the Stock Reconciliation desktop application.

## Requirements

- Python 3.8+
- pandas
- openpyxl (for Excel file support)
- numpy

## Install Dependencies

```bash
pip install pandas openpyxl numpy
```

## Usage

The processor is called automatically by the Electron frontend via python-shell.

### Manual Testing

```bash
python processor.py '{"action": "process", "unit": "Fath1", "stockFile": {"path": "path/to/stock.xlsx"}, "matchedFiles": {"bloc": {"path": "path/to/mov.xlsx"}}, "month": "12"}'
```
