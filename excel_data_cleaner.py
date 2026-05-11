"""
Excel Data Cleaner - Clean messy Excel/CSV files with one command.

Removes duplicates, fixes formatting, standardizes columns, handles missing data.
Supports batch processing and custom rules.
"""

import argparse
import sys
import os
import pandas as pd
import numpy as np
from pathlib import Path


class ExcelDataCleaner:
    """Clean and standardize messy Excel/CSV files."""
    
    def __init__(self, file_path: str, output_path: str = None):
        self.file_path = Path(file_path)
        self.output_path = output_path or self._get_default_output()
        self.df = None
        self.log = []
        
    def _get_default_output(self) -> str:
        """Generate default output filename."""
        return str(self.file_path.parent / f"cleaned_{self.file_path.name}")
    
    def load(self) -> pd.DataFrame:
        """Load Excel or CSV file."""
        ext = self.file_path.suffix.lower()
        if ext in ['.xls', '.xlsx']:
            self.df = pd.read_excel(self.file_path)
        elif ext == '.csv':
            self.df = pd.read_csv(self.file_path)
        else:
            raise ValueError(f"Unsupported file type: {ext}")
        
        initial_rows, initial_cols = self.df.shape
        self.log.append(f"Loaded: {initial_rows} rows, {initial_cols} columns")
        return self.df
    
    def remove_duplicates(self) -> pd.DataFrame:
        """Remove duplicate rows."""
        before = len(self.df)
        self.df = self.df.drop_duplicates()
        after = len(self.df)
        removed = before - after
        if removed > 0:
            self.log.append(f"Removed {removed} duplicate rows")
        return self.df
    
    def clean_column_names(self) -> pd.DataFrame:
        """Standardize column names: lowercase, strip whitespace, replace spaces with underscores."""
        original = list(self.df.columns)
        self.df.columns = (
            self.df.columns
            .str.strip()
            .str.lower()
            .str.replace(r'\s+', '_', regex=True)
            .str.replace(r'[^\w]', '', regex=True)
        )
        if list(self.df.columns) != original:
            self.log.append("Cleaned column names")
        return self.df
    
    def handle_missing_values(self, strategy: str = 'auto') -> pd.DataFrame:
        """Handle missing values based on column type.
        
        Args:
            strategy: 'auto' (default), 'drop', 'fill_mean', 'fill_median', 'fill_zero'
        """
        null_counts = self.df.isnull().sum()
        total_nulls = null_counts.sum()
        
        if total_nulls == 0:
            self.log.append("No missing values found")
            return self.df
        
        self.log.append(f"Found {total_nulls} missing values across {null_counts[null_counts > 0].shape[0]} columns")
        
        for col in self.df.columns:
            if self.df[col].isnull().sum() == 0:
                continue
            
            col_dtype = self.df[col].dtype
            
            if strategy == 'drop':
                self.df = self.df.dropna(subset=[col])
            elif strategy in ['fill_mean', 'fill_median', 'fill_zero'] or strategy == 'auto':
                if pd.api.types.is_numeric_dtype(col_dtype):
                    fill_val = {
                        'auto': self.df[col].median(),
                        'fill_mean': self.df[col].mean(),
                        'fill_median': self.df[col].median(),
                        'fill_zero': 0
                    }[strategy]
                    self.df[col] = self.df[col].fillna(fill_val)
                else:
                    # For text columns, fill with mode or empty string
                    mode_val = self.df[col].mode()
                    fill_val = mode_val[0] if len(mode_val) > 0 else ''
                    self.df[col] = self.df[col].fillna(fill_val)
        
        self.log.append("Handled missing values")
        return self.df
    
    def standardize_data_types(self) -> pd.DataFrame:
        """Standardize common data types."""
        # Strip whitespace from string columns
        for col in self.df.select_dtypes(include=['object']).columns:
            self.df[col] = self.df[col].astype(str).str.strip()
            # Replace 'nan' strings with empty
            self.df[col] = self.df[col].replace('nan', '')
        
        self.log.append("Standardized data types")
        return self.df
    
    def remove_empty_rows_columns(self) -> pd.DataFrame:
        """Remove completely empty rows and columns."""
        before_rows, before_cols = self.df.shape
        
        # Remove empty columns
        self.df = self.df.dropna(axis=1, how='all')
        
        # Remove empty rows
        self.df = self.df.dropna(axis=0, how='all')
        
        after_rows, after_cols = self.df.shape
        if before_rows > after_rows:
            self.log.append(f"Removed {before_rows - after_rows} empty rows")
        if before_cols > after_cols:
            self.log.append(f"Removed {before_cols - after_cols} empty columns")
        
        return self.df
    
    def clean_text_columns(self) -> pd.DataFrame:
        """Clean text: remove extra whitespace, standardize case, fix encoding issues."""
        for col in self.df.select_dtypes(include=['object']).columns:
            # Remove extra whitespace
            self.df[col] = self.df[col].str.replace(r'\s+', ' ', regex=True)
            # Remove leading/trailing whitespace
            self.df[col] = self.df[col].str.strip()
            # Remove non-printable characters
            self.df[col] = self.df[col].str.replace(r'[^\x20-\x7E]', '', regex=True)
        
        self.log.append("Cleaned text columns")
        return self.df
    
    def generate_report(self) -> str:
        """Generate a cleaning summary report."""
        report = "=== Data Cleaning Report ===\n"
        report += f"File: {self.file_path.name}\n"
        report += f"Final dimensions: {self.df.shape[0]} rows, {self.df.shape[1]} columns\n\n"
        report += "Actions taken:\n"
        for action in self.log:
            report += f"  - {action}\n"
        
        report += "\nColumn summary:\n"
        for col in self.df.columns:
            report += f"  {col}: {self.df[col].dtype}, {self.df[col].nunique()} unique values\n"
        
        return report
    
    def clean(self, all_steps: bool = True) -> pd.DataFrame:
        """Run all cleaning steps."""
        self.load()
        
        if all_steps:
            self.remove_duplicates()
            self.clean_column_names()
            self.handle_missing_values()
            self.standardize_data_types()
            self.remove_empty_rows_columns()
            self.clean_text_columns()
        
        return self.df
    
    def save(self, output_path: str = None) -> str:
        """Save cleaned data to file."""
        out_path = Path(output_path or self.output_path)
        ext = out_path.suffix.lower()
        
        if ext in ['.xls', '.xlsx']:
            self.df.to_excel(out_path, index=False)
        elif ext == '.csv':
            self.df.to_csv(out_path, index=False)
        else:
            # Default to CSV
            if out_path.suffix == '':
                out_path = out_path.with_suffix('.csv')
            self.df.to_csv(out_path, index=False)
        
        self.log.append(f"Saved to {out_path}")
        return str(out_path)


def main():
    parser = argparse.ArgumentParser(description='Clean messy Excel/CSV files')
    parser.add_argument('input', help='Input file path')
    parser.add_argument('-o', '--output', help='Output file path')
    parser.add_argument('--missing-strategy', choices=['auto', 'drop', 'fill_mean', 'fill_median', 'fill_zero'],
                       default='auto', help='Strategy for missing values')
    parser.add_argument('--report', action='store_true', help='Show cleaning report')
    parser.add_argument('--steps', nargs='+', help='Specific steps to run')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"Error: File not found: {args.input}")
        sys.exit(1)
    
    cleaner = ExcelDataCleaner(args.input, args.output)
    
    cleaner.load()
    
    if args.steps:
        step_map = {
            'duplicates': cleaner.remove_duplicates,
            'columns': cleaner.clean_column_names,
            'missing': lambda: cleaner.handle_missing_values(args.missing_strategy),
            'types': cleaner.standardize_data_types,
            'empty': cleaner.remove_empty_rows_columns,
            'text': cleaner.clean_text_columns,
        }
        for step in args.steps:
            if step in step_map:
                step_map[step]()
            else:
                print(f"Unknown step: {step}")
    else:
        # Run all steps
        cleaner.clean(all_steps=True)
    
    output_path = cleaner.save(args.output)
    print(f"Cleaned file saved to: {output_path}")
    print(f"Final shape: {cleaner.df.shape[0]} rows, {cleaner.df.shape[1]} columns")
    
    if args.report:
        print()
        print(cleaner.generate_report())


if __name__ == '__main__':
    main()
