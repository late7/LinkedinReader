#!/usr/bin/env python3
"""
Excel Preview Script
Reads an Excel file and prints the first 3 rows to understand data structure.
"""

import pandas as pd
import sys
import argparse
import os


def preview_excel(filename: str, num_rows: int = 3):
    """
    Read and preview the first few rows of an Excel file
    
    Args:
        filename: Path to the Excel file
        num_rows: Number of rows to preview (default: 3)
    """
    try:
        # Check if file exists
        if not os.path.exists(filename):
            print(f"❌ Error: File '{filename}' not found")
            return False
            
        print(f"📊 Reading Excel file: {filename}")
        print("=" * 80)
        
        # Read the Excel file
        df = pd.read_excel(filename)
        
        # Print basic info about the file
        print(f"📈 Total rows: {len(df)}")
        print(f"📈 Total columns: {len(df.columns)}")
        print(f"📈 Column names: {list(df.columns)}")
        print("=" * 80)
        
        # Print the first few rows
        print(f"🔍 First {min(num_rows, len(df))} rows:")
        print("-" * 80)
        
        # Display with better formatting
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 50)
        
        if len(df) == 0:
            print("⚠️  Excel file is empty")
        else:
            # Show row numbers starting from 1 (Excel-style)
            preview_df = df.head(num_rows).copy()
            preview_df.index = pd.RangeIndex(start=1, stop=len(preview_df) + 1)
            print(preview_df)
        
        print("-" * 80)
        
        # Show data types
        print("\n📋 Column data types:")
        for col, dtype in df.dtypes.items():
            print(f"  {col}: {dtype}")
            
        # Check for empty cells in first few rows
        print(f"\n🔍 Empty cells in first {min(num_rows, len(df))} rows:")
        if len(df) > 0:
            for idx in range(min(num_rows, len(df))):
                row_num = idx + 1
                empty_cols = df.iloc[idx].isna()
                if empty_cols.any():
                    empty_col_names = [col for col, is_empty in empty_cols.items() if is_empty]
                    print(f"  Row {row_num}: {empty_col_names}")
                else:
                    print(f"  Row {row_num}: No empty cells")
        
        return True
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return False


def parse_args():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description="Preview Excel file structure and first few rows"
    )
    parser.add_argument(
        "filename",
        help="Excel filename to preview (e.g., Investors2025.xlsx)"
    )
    parser.add_argument(
        "--rows", "-r",
        type=int,
        default=3,
        help="Number of rows to preview (default: 3)"
    )
    
    return parser.parse_args()


def main():
    """Main function"""
    args = parse_args()
    
    print("🚀 Excel Preview Tool")
    print("=" * 80)
    
    success = preview_excel(args.filename, args.rows)
    
    if success:
        print("\n✅ Preview completed successfully!")
    else:
        print("\n❌ Preview failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()