import pandas as pd
import os
from math import ceil

def split_spreadsheet(file_path, lines_per_file=None, num_files=None):
    """
    Função corrigida para dividir planilhas
    """
    try:
        file_ext = os.path.splitext(file_path)[1].lower()
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        
        # Determinar se é CSV ou Excel
        is_csv = file_ext in ['.csv', '.txt']
        
        # Ler o arquivo completo
        if is_csv:
            df = pd.read_csv(file_path, encoding='utf-8')
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
        
        total_lines = len(df)
        
        # Calcular linhas por arquivo
        if num_files:
            lines_per_file = ceil(total_lines / num_files)
        elif not lines_per_file:
            raise ValueError("Defina o número de linhas ou arquivos")
        
        # Dividir o DataFrame
        total_parts = ceil(total_lines / lines_per_file)
        generated_files = []
        
        for i in range(total_parts):
            start_idx = i * lines_per_file
            end_idx = min((i + 1) * lines_per_file, total_lines)
            part_df = df.iloc[start_idx:end_idx]
            
            # Salvar a parte
            if is_csv:
                output_file = f"{base_name}_parte_{i+1:02d}.csv"
                part_df.to_csv(output_file, index=False, encoding='utf-8')
            else:
                output_file = f"{base_name}_parte_{i+1:02d}.xlsx"
                part_df.to_excel(output_file, index=False, engine='openpyxl')
            
            generated_files.append(output_file)
        
        return generated_files, total_parts
        
    except Exception as e:
        raise Exception(f"Erro ao dividir planilha: {str(e)}")