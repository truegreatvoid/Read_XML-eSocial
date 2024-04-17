import pandas as pd
import os
from datetime import datetime
import xlsxwriter

def add_columns(file_path):
    # Carregar o Excel com todas as abas
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names  # Lista de todas as abas

    # Caminho para salvar o arquivo modificado
    output_path = os.path.join(f"modified_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}_" + os.path.basename(file_path))

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Processar cada aba individualmente
        for sheet in sheet_names:
            # Especificar que todas as colunas devem ser lidas como string
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)

            # Criar um novo DataFrame para armazenar os resultados
            new_df = pd.DataFrame()

            # Iterar sobre as colunas do DataFrame original
            for column in df.columns:
                # Adicionar a coluna original
                new_df[column] = df[column]
                # Adicionar coluna ATHENAS
                new_df[column + " (ATHENAS)"] = ""
                # Adicionar coluna DIFERENÇA
                new_df[column + " DIFERENÇA (SIM OU NÃO)"] = ""

            # Salvar o novo DataFrame modificado na aba correspondente
            new_df.to_excel(writer, sheet_name=sheet, index=False)

            # Formatar todas as colunas para texto no Excel
            workbook  = writer.book
            worksheet = writer.sheets[sheet]
            text_format = workbook.add_format({'num_format': '@'})  # Formato de texto
            for col_num in range(len(new_df.columns)):
                worksheet.set_column(col_num, col_num, None, text_format)

    print(f"Modified file saved as: {output_path}")
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
