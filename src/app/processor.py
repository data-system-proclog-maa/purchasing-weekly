import pandas as pd
import numpy as np
import os
from datetime import datetime
from typing import Dict, Optional, List
from app.localization import assign_department, split_by_department
from app.loader import load_excel_data

def process_procurement_data(
    po_file: str,
    rfm_file: str,
    datestart: str,
    dateend: str,
    normalization_file: Optional[pd.DataFrame] = None,
    output_dir: Optional[str] = None
) -> Dict[str, pd.DataFrame]:
    """
    Process procurement data for weekly reporting and save to Excel.
    
    Args:
        po_file: Path to the PO Excel file.
        rfm_file: Path to the RFM Excel file.
        datestart: Start date string in 'DD-MM-YYYY' format.
        dateend: End date string in 'DD-MM-YYYY' format.
        normalization_file: Optional DataFrame containing normalization data pulled from google sheet.
        output_dir: Optional directory to save output files. If None, uses po_file directory.
        
    Returns:
        Dictionary containing the processed dataframes.
    """
    
    if output_dir is None:
        output_dir = os.path.dirname(po_file)
        
    # Load original files
    df_po_original = load_excel_data(po_file)
    df_rfm_original = load_excel_data(rfm_file)

    # Make working copies
    df_po = df_po_original.copy()
    df_rfm = df_rfm_original.copy()

    # Define exclusions
    exclude_category = ['Jasa Logistik', 'Jasa/Service', 'Kontrak', 'Solar']
    exclude_requisition_type = ['Consignment']
    exclude_department = ['test']

    # Load Normalisasi file for updates
    picnorm_indexed = None
    if normalization_file is not None:
        try:
            # normalization_file is already a DataFrame
            picnorm = normalization_file
            # Ensure 'Requisition Number' exists
            if 'Requisition Number' in picnorm.columns:
                picnorm_indexed = picnorm.drop_duplicates(subset='Requisition Number').set_index('Requisition Number')
            else:
                 print("Warning: 'Requisition Number' column not found in normalization data.")
        except Exception as e:
            print(f"Error processing Normalization data: {e}. Skipping data enrichment.")
    else:
        print("Normalization data not provided. Skipping data enrichment.")

    # Apply data enrichment from Normalisasi file to both PO and RFM DataFrames
    if picnorm_indexed is not None:
        # Process df_po
        df_po['Updated Requisition Approved Date'] = df_po['Requisition Number'].map(picnorm_indexed.get('Updated Requisition Approved Date'))
        df_po['Updated Requisition Required Date'] = df_po['Requisition Number'].map(picnorm_indexed.get('Updated Requisition Required Date'))
        df_po['Background Update'] = df_po['Requisition Number'].map(picnorm_indexed.get('Background Update'))
        df_po['used_approved_date'] = df_po['Updated Requisition Approved Date'].fillna(df_po['Requisition Approved Date'])
        
        # Process df_rfm
        df_rfm['Updated Requisition Approved Date'] = df_rfm['Requisition Number'].map(picnorm_indexed.get('Updated Requisition Approved Date'))
        df_rfm['Updated Requisition Required Date'] = df_rfm['Requisition Number'].map(picnorm_indexed.get('Updated Requisition Required Date'))
        df_rfm['Background Update'] = df_rfm['Requisition Number'].map(picnorm_indexed.get('Background Update'))
        df_rfm['used_approved_date'] = df_rfm['Updated Requisition Approved Date'].fillna(df_rfm['Requisition Approved Date'])
    else:
        # If no Normalisasi file, use original dates
        df_po['used_approved_date'] = df_po['Requisition Approved Date']
        df_po['used_required_date'] = df_po['Requisition Required Date']
        df_rfm['used_approved_date'] = df_rfm['Requisition Approved Date']
        df_rfm['used_required_date'] = df_rfm['Requisition Required Date']
        df_po['Updated Requisition Approved Date'] = np.nan
        df_po['Updated Requisition Required Date'] = np.nan
        df_po['Background Update'] = np.nan
        df_rfm['Updated Requisition Approved Date'] = np.nan
        df_rfm['Updated Requisition Required Date'] = np.nan
        df_rfm['Background Update'] = np.nan

    # Base filters
    base_filter_po = (
        ~df_po['Item Category'].isin(exclude_category) &
        ~df_po['Requisition Type'].isin(exclude_requisition_type) &
        ~df_po['Department'].str.lower().isin([dept.lower() for dept in exclude_department])
    )

    base_filter_rfm = (
        ~df_rfm['Item Category'].isin(exclude_category) &
        ~df_rfm['Requisition Type'].isin(exclude_requisition_type) &
        ~df_rfm['Project'].str.lower().isin([dept.lower() for dept in exclude_department])
    )

    # Convert date columns
    df_po['PO Approval Date'] = pd.to_datetime(df_po['PO Approval Date'], errors='coerce')
    df_po['used_approved_date'] = pd.to_datetime(df_po['used_approved_date'], errors='coerce')
    df_rfm['used_approved_date'] = pd.to_datetime(df_rfm['used_approved_date'], errors='coerce')
    df_po['Updated Requisition Approved Date'] = pd.to_datetime(df_po['Updated Requisition Approved Date'], errors='coerce', dayfirst= True)
    df_po['Updated Requisition Required Date'] = pd.to_datetime(df_po['Updated Requisition Required Date'], errors='coerce', dayfirst= True)
    df_rfm['Updated Requisition Approved Date'] = pd.to_datetime(df_rfm['Updated Requisition Approved Date'], errors='coerce', dayfirst= True)
    df_rfm['Updated Requisition Required Date'] = pd.to_datetime(df_rfm['Updated Requisition Required Date'], errors='coerce', dayfirst= True)

    # Date filters
    try:
        datestart_dt = pd.to_datetime(datestart, format='%d-%m-%Y')
        dateend_dt = pd.to_datetime(dateend, format='%d-%m-%Y')
    except ValueError as e:
        raise ValueError(f"Error parsing dates. Please use DD-MM-YYYY format. Error: {e}")

    # Filtered data, using 'used_approved_date'
    New_RFMfromPO = df_po[base_filter_po &
                          (df_po['used_approved_date'] >= datestart_dt) &
                          (df_po['used_approved_date'] <= dateend_dt)]

    Inprocess_PO = df_po[base_filter_po &
                         ((df_po['PO Approval Date'].isna()) |
                          (df_po['PO Approval Date'] > dateend_dt))]

    PO_Approved = df_po[base_filter_po &
                        (df_po['PO Approval Date'] >= datestart_dt) &
                        (df_po['PO Approval Date'] <= dateend_dt)]

    New_RFMfromRFM = df_rfm[base_filter_rfm &
                            (df_rfm['Requisition Status'] == 'Approve') &
                            (df_rfm['used_approved_date'] >= datestart_dt) &
                            (df_rfm['used_approved_date'] <= dateend_dt)]

    Inprocess_RFM = df_rfm[base_filter_rfm &
                           (df_rfm['Requisition Status'] == 'Approve') &
                           (df_rfm['used_approved_date'] <= dateend_dt)]

    # Prepare PO results
    po_results = {
        'PO_Approved': assign_department(PO_Approved),
        'New_RFMfromPO': assign_department(New_RFMfromPO),
        'Inprocess_PO': assign_department(Inprocess_PO)
    }

    # Prepare RFM results
    rfm_results = {
        'New_RFMfromRFM': assign_department(New_RFMfromRFM),
        'Inprocess_RFM': assign_department(Inprocess_RFM)
    }

    # === Save PO ===
    po_output_path = os.path.join(output_dir, 'data_PO_Weekly.xlsx')
    print(f"Saving PO results to: {po_output_path}")
    with pd.ExcelWriter(po_output_path, engine='xlsxwriter') as writer:
        df_po_original.to_excel(writer, sheet_name='Sheet', index=False)  # Save raw input
        for base_name, df in po_results.items():
            # Base export
            df.drop(columns='Department_Assigned').to_excel(writer, sheet_name=base_name[:31], index=False)
            # Dept export
            dept_split = split_by_department(df)
            for dept, df_dept in dept_split.items():
                df_dept.to_excel(writer, sheet_name=f"{base_name}_{dept}"[:31], index=False)

    # === Save RFM ===
    rfm_output_path = os.path.join(output_dir, 'data_RFM_Weekly.xlsx')
    print(f"Saving RFM results to: {rfm_output_path}")
    with pd.ExcelWriter(rfm_output_path, engine='xlsxwriter') as writer:
        df_rfm_original.to_excel(writer, sheet_name='Sheet', index=False)  # Save raw input
        for base_name, df in rfm_results.items():
            # Base export
            df.drop(columns='Department_Assigned').to_excel(writer, sheet_name=base_name[:31], index=False)
            # Dept export
            dept_split = split_by_department(df)
            for dept, df_dept in dept_split.items():
                df_dept.to_excel(writer, sheet_name=f"{base_name}_{dept}"[:31], index=False)

    return {
        'po_output_path': po_output_path,
        'rfm_output_path': rfm_output_path
    }
