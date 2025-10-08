# data_handler.py
import pandas as pd
from typing import List, Dict, Tuple

def get_excel_data(file_path: str) -> Tuple[List[str], List[Dict[str, str]]]:
    try:
        df = pd.read_excel(file_path)
        df.fillna('', inplace=True)
        
        for col in df.columns:
            df[col] = df[col].astype(str)

        columns = df.columns.tolist()
        records = df.to_dict('records')
        return columns, records
    except Exception as e:
        raise ValueError(f"No se pudo leer el archivo Excel: {e}")