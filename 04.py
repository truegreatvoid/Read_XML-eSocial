import pandas as pd
import os
from datetime import datetime

def add_columns(file_path):
    # Carregar o Excel com todas as abas
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names  # Lista de todas as abas

    # Caminho para salvar o arquivo modificado
    output_path = os.path.join(os.path.dirname(file_path), f"modified_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}_" + os.path.basename(file_path))

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Processar cada aba individualmente
        for sheet in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            # Converter todas as colunas para texto
            df = df.astype(str).replace('nan', '')

            # Criar colunas adicionais e armazenar em um dicionário temporário
            additional_data = {}
            for column in df.columns:
                
                additional_data[column + " (ATHENAS)"] = [""] * len(df)  # Coluna ATHENAS vazia inicializada com strings vazias
                additional_data[column + " DIFERENÇA (SIM OU NÃO)"] = [""] * len(df)  # Coluna DIFERENÇA vazia

            # Usar pd.concat para adicionar todas as colunas de uma vez
            new_df = pd.concat([df, pd.DataFrame(additional_data, index=df.index)], axis=1)

            # Salvar o novo DataFrame modificado na aba correspondente
            new_df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"Modified file saved as: {output_path}")

def main():
    # Caminho para o arquivo Excel existente
    directory_path = 'C:\\Users\\tiago.xavier\\Desktop\\PS2-Sergio\\xml'
    file_name = 'Todos_Eventos_eSocial_17-04-2024_16-07-46.xlsx'  # Substitua pelo nome do seu arquivo
    file_path = os.path.join(directory_path, file_name)

    # Chamar a função para adicionar colunas
    add_columns(file_path)

if __name__ == '__main__':
    main()
