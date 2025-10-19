import pandas as pd

def merge_with_template(files, template_type):
    """
    Junta planilhas seguindo template específico
    """
    try:
        # Mapeamento de templates
        templates = {
            "Clientes": "ClientesModeloExcel_Financeiro.xlsx",
            "Equipamentos": "EquipamentosModeloExcel.xlsx", 
            "Produtos": "ProdutosModeloExcel.xlsx",
            "Questionarios": "QuestionariosModeloExcel.xlsx"
        }
        
        template_file = templates.get(template_type)
        if not template_file:
            raise ValueError(f"Template {template_type} não encontrado")
        
        # Ler template para estrutura
        template_df = pd.read_excel(f"modelos/{template_file}", engine='openpyxl', nrows=0)
        template_columns = template_df.columns.tolist()
        
        dataframes = []
        
        for file in files:
            file_ext = file.name.split('.')[-1].lower()
            
            if file_ext in ['csv', 'txt']:
                df = pd.read_csv(file, encoding='utf-8')
            else:
                df = pd.read_excel(file, engine='openpyxl')
            
            # Reordenar colunas conforme template e preencher faltantes
            for col in template_columns:
                if col not in df.columns:
                    df[col] = ""
            
            df = df[template_columns]
            dataframes.append(df)
        
        # Combinar
        merged_df = pd.concat(dataframes, axis=0, ignore_index=True)
        
        return merged_df, template_columns
        
    except Exception as e:
        raise Exception(f"Erro no juntador: {str(e)}")