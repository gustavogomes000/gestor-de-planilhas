import pandas as pd

def validate_spreadsheet(file, expected_columns):
    """
    Valida se a planilha tem a estrutura esperada
    """
    try:
        file_ext = file.name.split('.')[-1].lower()
        
        if file_ext in ['csv', 'txt']:
            df = pd.read_csv(file, encoding='utf-8', nrows=0)  # Apenas cabeçalho
        else:
            df = pd.read_excel(file, engine='openpyxl', nrows=0)
        
        missing_columns = [col for col in expected_columns if col not in df.columns]
        extra_columns = [col for col in df.columns if col not in expected_columns]
        
        return {
            'is_valid': len(missing_columns) == 0,
            'missing_columns': missing_columns,
            'extra_columns': extra_columns,
            'total_columns': len(df.columns)
        }
        
    except Exception as e:
        return {'is_valid': False, 'error': str(e)}