# Excel Data Cleaner

Clean messy Excel/CSV files with one command. Removes duplicates, fixes formatting, standardizes columns, handles missing data.

## Features

- Remove duplicate rows
- Standardize column names (lowercase, underscores)
- Smart missing value handling (auto-detect strategy)
- Standardize data types
- Remove empty rows/columns
- Clean text (whitespace, encoding issues)
- Generate cleaning reports
- Supports `.xlsx`, `.xls`, and `.csv` files

## Requirements

```
pandas>=2.0.0
numpy>=1.24.0
openpyxl>=3.1.0
```

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Clean all steps (default)
```bash
python excel_data_cleaner.py messy_data.csv
python excel_data_cleaner.py data.xlsx -o clean_data.xlsx
```

### Specific steps
```bash
python excel_data_cleaner.py sales.csv --steps duplicates missing text
```

### Custom missing value strategy
```bash
python excel_data_cleaner.py data.csv --missing-strategy fill_mean
python excel_data_cleaner.py data.csv --missing-strategy drop
```

### Show cleaning report
```bash
python excel_data_cleaner.py data.csv --report
```

## As a Library

```python
from excel_data_cleaner import ExcelDataCleaner

cleaner = ExcelDataCleaner('messy_file.xlsx')
cleaner.clean(all_steps=True)
cleaner.save('clean_output.xlsx')

# Or step by step
cleaner.load()
cleaner.remove_duplicates()
cleaner.handle_missing_values(strategy='auto')
cleaner.save('output.csv')

# Get report
print(cleaner.generate_report())
```

## Common Use Cases

1. **Survey data cleanup** - Remove incomplete responses, standardize categories
2. **Sales data normalization** - Fix column names, handle missing totals
3. **Contact list deduplication** - Remove duplicate entries, clean formatting
4. **Log file processing** - Standardize timestamps, remove empty rows

## License

MIT License - Use freely for personal and commercial projects.
