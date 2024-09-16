import pandas as pd
import os


def type_PLC_create(name:str, direction:str, type_:str, desc:str, unit:str, all_object:str)->str:
    """
    Создание строки-описания элемента ПЛК в Типе
        name - имя элемента;
        direction - направление элемента;
        type_ - тип элемента;
        desc - описание элемента;
        unit - еденицы измерения элемента;
        all_object - строка содержащая описание ранее созданных элементов.
    """
    
    example_object = f'\t\t\t<ct:parameter name="{name}" access-level="public" access-scope="global" direction="{direction}" type="{type_}" uuid="">\n\t\t\t\t<attribute type="unit.System.Attributes.Description" value="{desc}" />\n\t\t\t\t<attribute type="unit.Server.Attributes.Unit" value="{unit}" />\n\t\t\t</ct:parameter>'
    all_object += example_object+"\n"

    return all_object

def type_IOS_create(type_name:str, name:str, direction:str, type_:str, all_object:str)->str:
    """
    Создание строки-описания элемента Сервера в Типе
        type_name - имя типа;
        name - имя элемента;
        direction - направление элемента;
        type_ - тип элемента;
        all_object - строка содержащая описание ранее созданных элементов.
    """
    example_object = ""
    if direction == "in":
        example_object += f'\t\t\t<ct:bind source="In{name}" target="_{type_name}PLC.{name}" action="set_all" />\n'
        example_object += f'\t\t\t<ct:parameter name="In{name}" access-level="public" access-scope="global" direction="{direction}" type="{type_}" uuid="" />'
    elif direction == "out":
        example_object += f'\t\t\t<ct:bind source="_{type_name}PLC.{name}" target="Out{name}" action="set_all" />\n'
        example_object += f'\t\t\t<ct:parameter name="Out{name}" access-level="public" access-scope="global" direction="{direction}" type="{type_}" uuid="" />'
    elif direction == "in-out":
        example_object += f'\t\t\t<ct:bind source="In{name}" target="_{type_name}PLC.{name}" action="set_all" />\n\t\t\t<ct:bind source="_{type_name}PLC.{name}" target="Out{name}" action="set_all" />\n'
        example_object += f'\t\t\t<ct:parameter name="In{name}" access-level="public" access-scope="global" direction="{direction}" type="{type_}" uuid="" />\n'
        example_object += f'\t\t\t<ct:parameter name="Out{name}" access-level="public" access-scope="global" direction="{direction}" type="{type_}" uuid="" />'
    all_object += example_object+"\n"
    
    return all_object

def main_create(name:str, sheet:str, type_name:str, namespace:str)->None:
    """
    Создание файла для импорта Типа
        name - имя файла с описанием столбцов;
        sheet - имя листа;
        type_name - имя типа;
        namespace - папка, в которой храниться тип.
    """
    
    data = pd.read_excel("input/"+name, sheet_name=sheet)
    data_type = f'<omx xmlns="system" migration="29" xmlns:ct="automation.control" xmlns:r="automation.reference">\n\t<namespace name="{namespace}" uuid="">\n\t\t<ct:type name="{type_name}PLC" aspect="Aspects.PLC" uuid="">\n'
    for i in range(len(data)):
        data_type = type_PLC_create(data["Name"][i], data["Direction"][i], data["Type"][i], data["Description"][i], data["Unit"][i], data_type)
    data_type += f'\t\t</ct:type>\n\t\t<ct:type name="{type_name}IOS" aspect="Aspects.IOS" original="{type_name}PLC" uuid="">\n'
    for i in range(len(data)):
        data_type = type_IOS_create(type_name, data["Name"][i], data["Direction"][i], data["Type"][i], data_type)
        
    data_type += f'\t\t\t<r:ref name="_{type_name}PLC" type="{type_name}PLC" const-access="false" aspected="true" uuid="" />\n'
    data_type += "\t\t</ct:type>\n\t</namespace>\n</omx>"
    
    with open(f'output/Type.{sheet}.omx-export', "w", encoding='utf-8') as f:
        f.write(data_type)
