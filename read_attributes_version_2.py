  
# LIBRARIES
import os
import xml.etree.ElementTree as ET
import pandas as pd

# CODE
def process_attribute(attribute_element):
    attr_dict = {}
    for child in attribute_element:
        attr_name = child.tag
        attr_value = child.text.strip() if child.text and child.text.strip() else '<none>'
        attr_dict[attr_name] = attr_value
    return attr_dict

def process_xml_file(file_path):
    parser = ET.XMLParser()
    parser.entity["ematrixProductDtd"] = " "
    tree = ET.parse(file_path, parser=parser)
    root = tree.getroot()

    object_type_element = root.find('.//businessObject/objectType')
    object_type = object_type_element.text.strip() if object_type_element is not None and object_type_element.text and object_type_element.text.strip() else '<none>'

    attribute_list = root.find('.//businessObject/attributeList')
    
    attributes = []
    for attribute in attribute_list.findall('attribute'):
        processed_attribute = process_attribute(attribute)
        attributes.append(processed_attribute)
        
    df = pd.DataFrame(attributes)
    
    file_name = os.path.basename(file_path)
    df.insert(0, 'object_type', object_type)
    df.insert(1, 'file_name', file_name)
    
    return df

def process_folder(folder_path):
    dfs = {}
    file_names = []
    for root_dir, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xml'):
                file_path = os.path.join(root_dir, file)
                df = process_xml_file(file_path)
                object_type = df['object_type'][0]
                if object_type not in dfs:
                    dfs[object_type] = []
                dfs[object_type].append(df)
                file_names.append(file)
    return dfs, file_names

folder_path = input('Enter the main folder to process: ')

dfs, file_names = process_folder(folder_path)

output_file = 'Attributes.xlsx'
with pd.ExcelWriter(output_file) as writer:
    for object_type, df_list in dfs.items():
        combined_df = pd.concat(df_list, ignore_index=True)
        combined_df.to_excel(writer, sheet_name=object_type, index=False)

print(f'Totally added: {len(file_names)} files to report {output_file}:')
for file_name in file_names:
    print(f'- {file_name}')

