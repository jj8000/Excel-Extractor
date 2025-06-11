import pandas as pd
import os
import argparse
import fnmatch

def sheet_filter(sheet_name: str, patterns: list) -> bool:
    return any(fnmatch.fnmatch(sheet_name, pattern) for pattern in patterns)

def extract(filepath: str, output_dir: str, sheet_names: list):
    try:
        if sheet_names == ['*']:
            file_sheets = pd.read_excel(filepath, sheet_name=None)
        else:
            file_sheets = pd.read_excel(filepath, sheet_name=sheet_names)

        for sheet_name, data in file_sheets.items():
            filename = os.path.basename(filepath)
            filename_stripped = os.path.splitext(filename)[0]
            csv_name = f"{filename_stripped}_{sheet_name}.csv"
            output_path = os.path.join(output_dir, csv_name)
            data.to_csv(output_path, index=False)
            print(f"[✓] Saved: {output_path}")
    except Exception as e:
        print(f"[!] Failed to process {filepath}: {e}")

def args_interpreter():
    parser = argparse.ArgumentParser(description="Batch extract data from Excel sheets into .csv files.")
    parser.add_argument('input_folder', nargs='?', default='.',
                        help='Input folder containing Excel files (default: current dir).')
    parser.add_argument('output_folder', nargs='?', default='.',
                        help='Output folder for CSV files (default: current dir).')
    parser.add_argument('--sheets', nargs='*', default=['*'],
                        help='Specify sheet names to extract from. Supports Unix shell-style wildcards (default: all).')
    return parser.parse_args()

if __name__ == "__main__":
    cl_args = args_interpreter()
    os.makedirs(cl_args.output_folder, exist_ok=True)

    excel_files = [f for f in os.listdir(cl_args.input_folder) if f.lower().endswith((".xls", ".xlsx"))]

    if not excel_files:
        print("[!] No Excel files found.")
    else:
        for file in excel_files:
            filepath = os.path.join(cl_args.input_folder, file)
            print(f"[→] Processing: {filepath}")
            filtered_sheets = \
                [name for name in pd.read_excel(filepath, sheet_name=None).keys() if sheet_filter(name, cl_args.sheets)]
            if not filtered_sheets:
                print(f"[!] No matching sheets found in {file}.")
                continue
            extract(filepath, cl_args.output_folder, filtered_sheets)





