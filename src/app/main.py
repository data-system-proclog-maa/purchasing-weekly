import argparse
import os
import sys
import pandas as pd
from processor import process_procurement_data
from styler import style_and_reorder_excel_by_process
from typing import Optional

def get_input(prompt: str, required: bool = True) -> Optional[str]:
    while True:
        value = input(prompt).strip()
        if value:
            return value
        if not required:
            return None
        print("This field is required.")

def main():
    parser = argparse.ArgumentParser(description="Process procurement data for weekly reporting.")
    
    parser.add_argument("--po-file", help="Path to the PO Excel file.")
    parser.add_argument("--rfm-file", help="Path to the RFM Excel file.")
    parser.add_argument("--start-date", help="Start date in DD-MM-YYYY format.")
    parser.add_argument("--end-date", help="End date in DD-MM-YYYY format.")
    parser.add_argument("--normalization-file", help="Path to the Normalisasi.xlsx file (optional).")
    parser.add_argument("--output-dir", help="Directory to save output files (optional).")

    args = parser.parse_args()

    # Interactive fallback
    if not any(vars(args).values()):
        print("No arguments provided. Switching to interactive mode.")
        args.po_file = get_input("Enter path to PO file: ")
        args.rfm_file = get_input("Enter path to RFM file: ")
        args.start_date = get_input("Enter start date (DD-MM-YYYY): ")
        args.end_date = get_input("Enter end date (DD-MM-YYYY): ")
        print("Normalization file is pulled fromm google sheet")
        args.output_dir = get_input("Enter output directory (optional, press Enter to skip): ", required=False)
    else:
        # Validate required arguments if not in interactive mode
        missing = []
        if not args.po_file: missing.append("--po-file")
        if not args.rfm_file: missing.append("--rfm-file")
        if not args.start_date: missing.append("--start-date")
        if not args.end_date: missing.append("--end-date")
        
        if missing:
            print(f"Error: Missing required arguments: {', '.join(missing)}")
            parser.print_help()
            sys.exit(1)

    normalization_file = pd.read_csv(f'https://docs.google.com/spreadsheets/d/1EZ7kPPvnRqvR5UN0Vi0NNLpLTNXEArzRklsVTIGb1vc/gviz/tq?tqx=out:csv&gid=0')
    
    # Validate input files
    if not os.path.exists(args.po_file):
        print(f"Error: PO file not found at {args.po_file}")
        sys.exit(1)
    if not os.path.exists(args.rfm_file):
        print(f"Error: RFM file not found at {args.rfm_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    if args.output_dir and not os.path.exists(args.output_dir):
        os.makedirs(args.output_dir)

    try:
        # Process data
        print("Starting data processing...")
        output_files = process_procurement_data(
            po_file=args.po_file,
            rfm_file=args.rfm_file,
            datestart=args.start_date,
            dateend=args.end_date,
            normalization_file=normalization_file,
            output_dir=args.output_dir
        )
        
        # Style files
        print("Starting file styling...")
        style_and_reorder_excel_by_process(output_files['po_output_path'])
        style_and_reorder_excel_by_process(output_files['rfm_output_path'])
        
        print("Processing completed successfully.")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
