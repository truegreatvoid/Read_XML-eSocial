import pandas as pd
import os

def add_columns(file_path):
    # Carregar o arquivo Excel
    df = pd.read_excel(file_path)

    # Criar um novo DataFrame para armazenar os resultados
    new_df = pd.DataFrame()

    # Iterar sobre as colunas do DataFrame original
    for column in df.columns:
        new_df[column] = df[column]
        new_df[column + " (ATHENAS)"] = ""  # Coluna ATHENAS vazia
        new_df[column + " DIFERENÇA (SIM OU NÃO)"] = ""

    # Definir o caminho de saída
    output_path = os.path.join(os.path.dirname(file_path), "modified_" + os.path.basename(file_path))
    
    # Salvar o novo DataFrame modificado de volta em um arquivo Excel
    new_df.to_excel(output_path, index=False)
    print(f"Modified file saved as: {output_path}")

def main():
    # Caminho para o arquivo Excel existente
    directory_path = 'C:\\Users\\tiago.xavier\\Desktop\\PS2-Sergio\\xml'
    file_name = 'Todos_Eventos_eSocial_17-04-2024_16-07-46.xlsx'  # Substitua 'example.xlsx' pelo nome do seu arquivo
    file_path = os.path.join(directory_path, file_name)

    # Chamar a função para adicionar colunas
    add_columns(file_path)

if __name__ == '__main__':
    main()
