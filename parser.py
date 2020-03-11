import pandas as pd
import pandas.compat
import sys
import os.path
from pathlib2 import Path
from pandas import DataFrame
from pandas.io import json
from pandas.io.json._json import to_json
import json


#функция проверяет является ли передаваемый параметр json строкой и если да, преобразовывает ее в json object
def get_normalized_value(item):
    try:
        json_str=json.loads(str(item))
        return json_str
    except ValueError:
        return item

#функция создает путь к готовому файлу в зависимости от указанного параметра output_file
def output_file_path(input_file, output_file):
    if output_file ==  'None': 
        output_file = os.path.splitext(input_file)[0]+".json" 
    elif output_file.find('/') == -1:
        output_file = os.path.dirname(input_file)+'/'+output_file
    return output_file


class Json_parser:

    def __init__(self, input_file, output_file):
        self.input_file=input_file
        self.output_file=output_file


    def make_json(self):
        #создаем dict в котором key это вложенные листы excel, value это список из DataFrame 
        data_parced = pd.read_excel(self.input_file, None) 
        dict_jsons = {}

        for key, value in data_parced.items():
            json_obj_list = json.loads(value.to_json( #конвертируем DataFrame в json строку
                None, orient='records', date_format='iso'))
            for item in json_obj_list:
                for key, value in item.items():
                    item[key]=get_normalized_value(value)#меняем данные в ячеке на json object, если данные это json строка
            dict_jsons[key] = json_obj_list

        with open(self.output_file, 'w') as result:
                #кладем получившийся dict в файл по указанному пути
            json.dump(dict_jsons, result, ensure_ascii=False,
                    indent=4, separators=(',', ': '))


def main(input_file, output_file='None'):
    output_file=output_file_path(input_file,output_file) 

    if os.path.exists(output_file):
        print("File already exists!")
    else:
        parsing_data=Json_parser(input_file,output_file)
        parsing_data.make_json()

try:
    main(sys.argv[1], sys.argv[2])
except IndexError:
    main(sys.argv[1])

