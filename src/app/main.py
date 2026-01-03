import argparse
import os
import sys
import pandas as pd
from app.processor import process_procurement_data #added app for ingestion
from app.styler import style_and_reorder_excel_by_process #added app for ingestion
from typing import Optional

def get_input(prompt: str, required: bool = True) -> Optional[str]:
    while True:
        value = input(prompt).strip()
        if value:
            return value
        if not required:
            return None
        print("This field is required.")

def run(
    po_file: str,
    rfm_file: str,
    start_date: str,
    end_date: str,
    normalization_file: Optional[pd.DataFrame] = None,
    output_dir: Optional[str] = None
):
    """
    Programmatic entry point for processing procurement data.
    """
    # Validate input files
    if not os.path.exists(po_file):
        raise FileNotFoundError(f"PO file not found at {po_file}")
    if not os.path.exists(rfm_file):
        raise FileNotFoundError(f"RFM file not found at {rfm_file}")

    # Create output directory if it doesn't exist
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Use default normalization file if not provided
    if normalization_file is None:
        try:
            normalization_file = pd.read_csv(f'https://docs.google.com/spreadsheets/d/1EZ7kPPvnRqvR5UN0Vi0NNLpLTNXEArzRklsVTIGb1vc/gviz/tq?tqx=out:csv&gid=0')
        except Exception as e:
            print(f"Warning: Could not fetch normalization file: {e}")
            normalization_file = None

    try:
        # Process data
        print("Starting data processing...")
        output_files = process_procurement_data(
            po_file=po_file,
            rfm_file=rfm_file,
            datestart=start_date,
            dateend=end_date,
            normalization_file=normalization_file,
            output_dir=output_dir
        )
        
        # Style files
        print("Starting file styling...")
        style_and_reorder_excel_by_process(output_files['po_output_path'])
        style_and_reorder_excel_by_process(output_files['rfm_output_path'])
        
        print("Processing completed successfully.")
        return output_files
        
    except Exception as e:
        print(f"An error occurred: {e}")
        raise

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

    try:
        run(
            po_file=args.po_file,
            rfm_file=args.rfm_file,
            start_date=args.start_date,
            end_date=args.end_date,
            output_dir=args.output_dir
        )
    except Exception as e:
        sys.exit(1)

if __name__ == "__main__":
    main()
