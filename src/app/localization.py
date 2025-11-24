import pandas as pd
from typing import Dict, List, Optional

# define pic list, name in list with / is from rfm name in 
pic_OBI: List[str] = ['Rona / Joko', 'Joko', 'Victo', 'Rakan', 'Rona Justhafist', 'Rona / Victo / Rakan / Joko']
pic_LAR: List[str] = ['Fairus / Irwan', 'Fairus Mubakri', 'Irwan', 'Ady', 'Fairus / Ady']
pic_HO: List[str] = ['Linda / Puji / Syifa R / Stheven', 'Syifa Ramadhani', 'Syifa Alifia', 'Rizal Agus Fianto',
      'Auriel', 'Puji Astuti', 'Linda Permata Sari']

def assign_department(df: pd.DataFrame) -> pd.DataFrame:
    """
    assign dept based on procurement name.
    
    Args:
        df: input po dataframe.
        
    Returns:
        df with 'Department_Assigned' column.
    """
    def get_dept(name):
        if pd.isna(name): return None
        for dept, names in [('OBI', pic_OBI), ('LAR', pic_LAR), ('HO', pic_HO)]:
            if any(n in name for n in names):
                return dept
        return None
    
    df = df.copy()
    df['Department_Assigned'] = df['Procurement Name'].apply(get_dept)
    return df

def split_by_department(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """
    split df by department.
    
    Args:
        df: input po dataframe.
        
    Returns:
        dictionary where keys are department names and values are df.
    """
    result: Dict[str, pd.DataFrame] = {}
    for dept in ['OBI', 'LAR', 'HO']:
        if 'Department_Assigned' not in df.columns:
             return {}
             
        df_dept = df[df['Department_Assigned'] == dept].drop(columns='Department_Assigned')
        if not df_dept.empty:
            result[dept] = df_dept
    return result
