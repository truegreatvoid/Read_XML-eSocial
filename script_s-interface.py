import os
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

def parse_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Extrair e usar namespaces
    namespaces = {k.split('}')[-1]: v for k, v in root.attrib.items() if 'xmlns' in k}
    ns_url = list(namespaces.values())[0] if namespaces else ''

    dados = {}

    def recurse_elements(element, path=[]):
        for child in element:
            child_tag = child.tag.split('}')[-1]  # Remove namespace
            new_path = path + [child_tag]
            if child.text and child.text.strip():
                # Pegar apenas o penúltimo e último segmento para o nome da coluna
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

def main(directory_path, output_directory):
    event_files = {}
    # Identify all XML files and group them by event type
    for f in os.listdir(directory_path):
        if f.endswith('.xml'):
            event_type = f.split('.')[-2].upper()  # Get event type from file name
            if event_type in event_files:
                event_files[event_type].append(os.path.join(directory_path, f))
            else:
                event_files[event_type] = [os.path.join(directory_path, f)]

    # Process files for each event type and store in different sheets
    with pd.ExcelWriter(os.path.join(output_directory, f'All_Events_{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx')) as writer:
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
                print(f"Data for {event_type} saved.")
            else:
                print(f"No data extracted for {event_type}.")

if __name__ == '__main__':
    directory_path = 'C:\\Users\\tiago.xavier\\Desktop\\PS2-Sergio\\xml\\1610\\all'
    output_directory = 'C:\\Users\\tiago.xavier\\Desktop\\PS2-Sergio\\xml'
    main(directory_path, output_directory)
