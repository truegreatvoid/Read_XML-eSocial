import pandas as pd
import os

def add_columns(file_path):
    # Carregar o Excel com todas as abas
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names  # Lista de todas as abas

    # Caminho para salvar o arquivo modificado
    output_path = os.path.join(os.path.dirname(file_path), "modified_" + os.path.basename(file_path))

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Processar cada aba individualmente
        for sheet in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            # Criar um novo DataFrame para armazenar os resultados
            new_df = pd.DataFrame()

            # Iterar sobre as colunas do DataFrame original
            for column in df.columns:
                # Assegurar que os dados sejam tratados como texto
                df[column] = df[column].astype(str)

                new_df[column] = df[column]
                # Adicionar colunas vazias formatadas como texto
                new_df[column + " (ATHENAS)"] = ''
                new_df[column + " DIFERENÇA (SIM OU NÃO)"] = ''

            # Definir as colunas como texto no Excel
            for col in new_df.columns:
                new_df[col] = new_df[col].astype(str)
                col_idx = new_df.columns.get_loc(col)
                writer.sheets[sheet].set_column(col_idx, col_idx, len(new_df[col].max()) + 1)

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
