"""
This script converts Microsoft Excel Open XML Spreadsheet (.xlsx) file to tab separated value file (.tsv)
that can be used in Jira to import projects and issues.
"""

import pandas as pd
import json
import sys

try:
    df = pd.read_excel(sys.argv[1])
except OSError:
    print(f'Incorrect data reading from {0}', sys.argv[1])
    sys.exit(1)
# df = pd.read_excel('test/exportSD.xlsx')

with open("config.json", "r", encoding='utf-8') as config:
    config_map = json.load(config)
with open("dt_columns.json", "r", encoding='utf-8') as date:
    dt_columns = json.load(date)

for key in dt_columns:
    df[key] = df[key].dt.round('1s')

for key in config_map.keys():
    df[key] = df[key].str.replace('\n', ' ')
    df[key] = df[key].str.replace('\t', '')
    df = df.replace({key: config_map[key]})

df.to_csv(sys.argv[1].replace('.xlsx', '.csv'), sep='\t')
