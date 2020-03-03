import pandas as pd
import sys
import os.path
from pathlib2 import Path
from pandas import DataFrame
from pandas.io import json
from pandas.io.json._json import to_json
from pprint import pprint
import json


def output_file_path(input_file, output_file):
    if output_file == None:
        output_file = os.path.splitext(input_file)[0]+".json"
    elif output_file.find('/') == -1:
        output_file = os.path.dirname(input_file)+'/'+output_file
    return output_file
    


def make_json(input_file, output_file):
    data_path = input_file
    data_parced = pd.read_excel(data_path, None)
    dict_jsons = {}
    for key, value in data_parced.items():
        value = json.loads(value.to_json(
            None, orient='records', date_format='iso'))
        dict_jsons[key] = value

    with open(output_file, 'w') as result:
        json.dump(dict_jsons, result, ensure_ascii=False,
                  indent=4, separators=(',', ': '))


def main(input_file, output_file=None):

    output_file = output_file_path(input_file, output_file)
    

    if os.path.exists(output_file):
        print("File already exists!")
    else:
        make_json(input_file, output_file)


if __name__ == "__main__":
    try:
        main(sys.argv[1], sys.argv[2])
    except IndexError:
        main(sys.argv[1])

