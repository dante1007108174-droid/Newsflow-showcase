#!/usr/bin/env python3
"""Read and analyze badcase合集.xlsx"""

import pandas as pd
import sys

def main():
    file_path = '/d/daily-ai-news/docs/badcase合集.xlsx'
    
    try:
        # Read all sheets
        xl = pd.ExcelFile(file_path)
        print("=== Sheets in file ===")
        for i, sheet in enumerate(xl.sheet_names):
            print(f"{i+1}. {sheet}")
        
        print("\n=== Reading first sheet ===")
        df = pd.read_excel(file_path, sheet_name=0)
        
        print(f"\nShape: {df.shape}")
        print(f"Columns ({len(df.columns)}): {list(df.columns)}")
        print(f"\nFirst 30 rows:")
        print(df.head(30).to_string())
        
        # Check for any other sheets
        if len(xl.sheet_names) > 1:
            for sheet_name in xl.sheet_names[1:]:
                print(f"\n\n=== Sheet: {sheet_name} ===")
                df2 = pd.read_excel(file_path, sheet_name=sheet_name)
                print(f"Shape: {df2.shape}")
                print(f"Columns: {list(df2.columns)}")
                print(df2.head(10).to_string())
        
        return 0
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
