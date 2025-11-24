import pandas as pd
import os
from typing import Optional

def load_excel_data(file_path: str, sheet_name: Optional[str] = 0) -> pd.DataFrame:
    """
    load excel data from cps.
    
    Args:
        file_path: path to file.
        sheet_name: sheet name or index. Defaults to 0 (first sheet).
        
    Returns:
        loaded dataframe.
        
    Raises:
        FileNotFoundError: if file not found.
        ValueError: if file cannot be read.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    try:
        print(f"Loading file: {file_path}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        raise ValueError(f"Error reading file {file_path}: {e}")
