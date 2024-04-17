import os
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def parse_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Extrair e usar namespaces
    namespaces = {k.split('}')[-1]: v for k, v in root.attrib.items() if 'xmlns' in k}
    ns_url = list(namespaces.values())[0] if namespaces else ''

    dados = {}

    def recurse_elements(element, path=[]):
        for child in element:
            child_tag = child.tag.split('}')[-1]
            new_path = path + [child_tag]
            if child.text and child.text.strip():
                column_name = '/'.join(new_path[-2:]) if len(new_path) > 1 else new_path[-1]
                if column_name in dados:
                    dados[column_name].append(child.text.strip())
                else:
                    dados[column_name] = [child.text.strip()]
            recurse_elements(child, new_path)

    evt_admissao_element = root.find('.//{*}eSocial', namespaces={'': ns_url})
    if evt_admissao_element is not None:
        recurse_elements(evt_admissao_element)

    return dados

def process_files(directory_path, output_directory):
    try:
        event_files = {}
        for f in os.listdir(directory_path):
            if f.endswith('.xml'):
                event_type = f.split('.')[-2].upper()
                if event_type in event_files:
                    event_files[event_type].append(os.path.join(directory_path, f))
                else:
                    event_files[event_type] = [os.path.join(directory_path, f)]

        with pd.ExcelWriter(os.path.join(output_directory, f'Todos_Eventos_eSocial_{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx')) as writer:
            for event_type, files in event_files.items():
                all_data = []
                for xml_file in files:
                    parsed_data = parse_xml(xml_file)
                    if parsed_data:
                        df = pd.DataFrame({k: pd.Series(v) for k, v in parsed_data.items()})
                        all_data.append(df)
                if all_data:
                    full_df = pd.concat(all_data, ignore_index=True)
                    full_df.to_excel(writer, sheet_name=event_type, index=False)
            messagebox.showinfo("Relatório Gerado!", "O relatório foi processado e salvo com sucesso.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocorreu um erro: {e}")

def main():
    root = tk.Tk()
    root.title("eSocial XML - Relatório por Eventos")
    
    ttk.Label(root, text="Selecione o diretório que contém os arquivos:").grid(row=0, column=0, padx=10, pady=10)
    input_dir = tk.StringVar()
    input_entry = ttk.Entry(root, textvariable=input_dir, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=10)
    ttk.Button(root, text="Selecionar", command=lambda: input_dir.set(filedialog.askdirectory())).grid(row=0, column=2, padx=10, pady=10)
    
    ttk.Label(root, text="Selecione o diretório de saída do relatório:").grid(row=1, column=0, padx=10, pady=10)
    output_dir = tk.StringVar()
    output_entry = ttk.Entry(root, textvariable=output_dir, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=10)
    ttk.Button(root, text="Selecionar", command=lambda: output_dir.set(filedialog.askdirectory())).grid(row=1, column=2, padx=10, pady=10)
    
    ttk.Button(root, text="Gerar relatório", command=lambda: process_files(input_dir.get(), output_dir.get())).grid(row=2, column=1, padx=10, pady=20)
    
    root.mainloop()

if __name__ == '__main__':
    main()
