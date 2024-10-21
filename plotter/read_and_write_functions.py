import pandas as pd
from os.path import join, split, splitext
from os import listdir
from typing import Dict
import re


def loadEMT(infFile: str) -> pd.DataFrame:
    '''
    Load EMT results from a collection of csv files defined by the given inf file. Returns a dataframe with index 'time'.
    '''
    folder, filename = split(infFile)
    filename, fileext = splitext(filename)

    assert fileext.lower() == '.inf'

    adjFiles = listdir(folder)
    csvMap: Dict[int, str] = dict()
    pat = re.compile(r'^' + filename.lower() + r'(?:_([0-9]+))?.csv$')

    for file in adjFiles:
        rem = re.match(pat, file.lower())
        if rem:
            if rem.group(1) is None:
                id = -1
            else:
                id = int(rem.group(1))
            csvMap[id] = join(folder, file)
    csvMaps = list(csvMap.keys())
    csvMaps.sort()

    df = pd.DataFrame()
    firstFile = True
    loadedColumns = 0
    for map in csvMaps:
        dfMap = pd.read_csv(csvMap[map], skiprows=1, header=None)  # type: ignore
        if not firstFile:
            dfMap = dfMap.iloc[:, 1:]
        else:
            firstFile = False
        dfMap.columns = list(range(loadedColumns, loadedColumns + len(dfMap.columns)))
        loadedColumns = loadedColumns + len(dfMap.columns)
        df = pd.concat([df, dfMap], axis=1)  # type: ignore

    columns = emtColumns(infFile)
    columns[0] = 'time'
    df = df[columns.keys()]
    df.rename(columns, inplace=True, axis=1)
    print(f"Loaded {infFile}, length = {df['time'].iloc[-1]}s")  # type: ignore
    return df


def emtColumns(infFilePath: str) -> Dict[int, str]:
    '''
    Reads EMT result columns from the given inf file and returns a dictionary with the column number as key and the column name as value.
    '''
    columns: Dict[int, str] = dict()
    with open(infFilePath, 'r') as file:
        for line in file:
            rem = re.match(
                r'^PGB\(([0-9]+)\) +Output +Desc="(\w+)" +Group="(\w+)" +Max=([0-9\-\.]+) +Min=([0-9\-\.]+) +Units="(\w*)" *$',
                line)
            if rem:
                columns[int(rem.group(1))] = rem.group(2)
    return columns