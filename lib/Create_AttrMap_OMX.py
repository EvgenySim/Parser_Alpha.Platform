import pandas as pd
import os


def attribut_create(name: str, param:str, all_object:str):
    
    example_object = f'\t<item id="{name}" value="{param}" />'
    all_object += example_object+"\n"

    return all_object

def main_create(name:str, sheet:str, name_element:str, parametr:str)->None:
    """
    Создание файла для импорта атрибутов
        name - имя файла с описанием столбцов;
        sheet - имя листа;
        name_element - имя файла хранения значений атрибутов с разрещением xml;
        parametr - параметр атрибутов.
    """

    data = pd.read_excel(f'input\\{name}', sheet_name=sheet)

    data_map = f'<omx xmlns="system" migration="29" xmlns:dp="automation.deployment">\n\t<dp:attributes-map name="{sheet}" type="Server.Attributes.{parametr}" file="{name_element}.xml" uuid="" />\n</omx>'

    with open(f'output\\Attrib.{sheet}.omx-export', "w", encoding='utf-8') as f:
        f.write(data_map)

    data_map = '<AttributesMap>\n'

    for i in range(len(data)):
            data_map = attribut_create(data["FullName"][i], data[f"{parametr}"][i], data_map)

    data_map += "</AttributesMap>"

    with open(f'output\\{name_element}.xml', "w", encoding='utf-8') as f:
        f.write(data_map)



#main_create("Input.xlsx","CreateAttrib","Attr","Unit")
