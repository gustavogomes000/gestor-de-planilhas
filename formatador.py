import pandas as pd
import os

def formatar_com_modelo(lista_arquivos, caminho_modelo, progresso_callback=None):
    """
    Formata planilhas usando a planilha modelo como base.
    """
    modelo = pd.read_excel(caminho_modelo)
    colunas_modelo = list(modelo.columns)

    for idx, arquivo in enumerate(lista_arquivos):
        ext = os.path.splitext(arquivo)[1].lower()
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(arquivo)
        else:
            df = pd.read_csv(arquivo)

        df = df.reindex(columns=colunas_modelo, fill_value="")

        df.to_excel(arquivo, index=False)

        if progresso_callback:
            progresso_callback(int((idx+1)/len(lista_arquivos)*100))
