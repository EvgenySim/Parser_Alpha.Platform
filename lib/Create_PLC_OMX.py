import pandas as pd
import os


def plc_create(name:str, base_type:str, desc:str, all_object:str)->str:
    """
    Создание строки-описания элемента в представлении ПЛК
        name - имя элемента (берется из ТС);
        base_type - базовый тип элемента (путь до типа элемента в проекте DevStudio);
        desc - описание элемента (берется из ТС);
        all_object - строка содержащая описание ранее созданных элементов.
    """
    
    example_object = f'\t<ct:object name="{name}" uuid="" base-type="{base_type}" aspect="Aspects.PLC">\n\t\t<attribute type="unit.System.Attributes.Description" value="{desc}" />\n\t</ct:object>'
    all_object += example_object+"\n"

    return all_object

def main_create(name:str, sheet:str)->None:
    """
    Создание файла для импорта элементов в представлении ПЛК
        name - имя файла с описанием столбцов;
        sheet - имя листа.
    """
    
    data = pd.read_excel("input/"+name, sheet_name=sheet)

    data_plc = '<omx xmlns="system" migration="29" xmlns:ct="automation.control">\n'

    for i in range(len(data)):
        if not "-" in data["CodePLC"][i] and not "Канал" in data["CodePLC"][i]:
            data_plc = plc_create(data["CodePLC"][i], data["BaseType"][i], data["Description"][i], data_plc)

    data_plc += "</omx>"

    with open(f'output/PLC.{sheet}.omx-export', "w", encoding='utf-8') as f:
        f.write(data_plc)

