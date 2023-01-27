from os import listdir, makedirs
from os.path import join, split, splitext, exists
import re
import csv
import pandas as pd
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from configparser import ConfigParser
from types import SimpleNamespace
from typing import List, Dict, Union, Tuple, Set

def idFile(filePath: str) -> Tuple[Union[int, None], Union[str, None], Union[int, None]]:
    path, fileName = split(filePath)
    match = re.match(r'^(\w+?)_([0-9]+).(inf|csv)$', fileName.lower())
    if match:
        caseId = int(match.group(2))
        project = join(path, match.group(1))
        with open(filePath, 'r') as file:
            firstLine = file.readline()
            if match.group(3) == 'inf' and firstLine.startswith('PGB(1)'):
                fileType = 1
                return (fileType, project, caseId)
            elif match.group(3) == 'csv':
                secondLine = file.readline()
                if secondLine.startswith(r'"b:tnow in s"'):
                    fileType = 0
                    return (fileType, project, caseId)
    return (None, None, None)

def mapResultFiles(dirs: list) -> Tuple[Dict[int, List[Tuple[int, str, str]]], Set[str]]:
    files = [join(dir, p) for dir in dirs for p in listdir(dir)]

    cases = {}
    relevantProjects = set()

    for file in files:
        typ, project, id = idFile(file)
        if typ is None:
            continue

        if typ == 0:
            project = f"{project}_RMS"
        elif typ == 1:
            project = f"{project}_EMT"

        cases.setdefault(id, []).append((typ, project, file))
        relevantProjects.add(project)
    return cases, relevantProjects

def readFigureSetup(setupFile : str) -> List[Dict[str, Union[int, str]]]:
    setup = list()
    with open(setupFile, newline='') as setupFile:
        setupReader = csv.DictReader(setupFile, delimiter = ';')
        for row in setupReader:
            setup.append(row)
    return setup

def initFigure(figureName : str, ncolumns : int, nrows : int):
    figure = make_subplots(rows=nrows,cols=ncolumns)
    figure.update_layout(title_text=figureName)
    return figure
    
def readConfig() -> SimpleNamespace:
    cp = ConfigParser()
    cp.read('config.ini')
    parsedConf = cp['config']
    config = SimpleNamespace()
    config.resultsDir = parsedConf['resultsDir']
    config.figureSetupfilePath = parsedConf['figureSetupfilePath']
    config.columns = parsedConf.getint('columns')                 
    config.genHTML = parsedConf.getboolean('genHTML')
    config.genJPEG = parsedConf.getboolean('genJPEG')
   
    config.simDataDirs = list()
    simPaths = cp.items('Simulation data paths')
    for _, path in simPaths:
        config.simDataDirs.append(path)

    return config

def emtColumns(file : str) -> Dict[int, str]:
    columns = dict()
    with open(file, 'r') as file:
        for line in file: 
            rem = re.match(r'^PGB\(([0-9]+)\) +Output +Desc="(\w+)" +Group="(\w+)" +Max=([0-9\-\.]+) +Min=([0-9\-\.]+) +Units="(\w*)" *$', line)
            if rem:
                columns[int(rem.group(1))] = rem.group(2) 
    return columns

def loadEMT(infFile : str) -> pd.DataFrame:
    folder, filename = split(infFile)
    filename, fileext = splitext(filename)
    if fileext != '.inf':
        return None

    adjFiles = listdir(folder)
    csvMap = dict()
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
        dfMap = pd.read_csv(csvMap[map], skiprows = 1, header = None)
        if not firstFile:
            dfMap = dfMap.iloc[: , 1:]
        else:
            firstFile = False
        dfMap.columns = list(range(loadedColumns, loadedColumns + len(dfMap.columns)))
        loadedColumns = loadedColumns + len(dfMap.columns)
        df = pd.concat([df, dfMap], axis=1)

    columns = emtColumns(infFile)
    columns[0] = 'time'
    df = df[columns.keys()]
    df.rename(columns, inplace=True, axis=1)    
    return df

def addResultToFig(typ: int, result: pd.DataFrame, figureSetup: List[Dict[str, Union[int, str]]], figure, project: str, file: str, colors: Dict[str, List[str]], nColumns: int) -> None: 
    for fSetup in figureSetup:
        fid = int(fSetup['figure'])
        rowPos = (fid - 1) // nColumns + 1
        colPos = (fid - 1) % nColumns + 1
        traces = 0

        for sig in range(1,4):
            signalKey = 'rms' if typ == 0 else 'emt'
            sigName = fSetup.get(f"{signalKey}_signal_{sig}", "")
            if sigName in result.columns:
                figure.append_trace(
                    go.Scatter(
                    x=result['time' if typ == 1 else 'b:tnow in s'],
                    y=result[sigName],
                    line_color=colors[project][traces], 
                    name=f"{file}:{sigName}",
                    legendgroup=project,
                    showlegend=True
                ),
                row=rowPos, col=colPos
                )
                traces += 1
            elif sigName != '':
                print(f"Signal '{sigName}' not recognized in resultfile '{file}'")
        
        figure.update_xaxes(
            title_text='Time[s]',  
            row=rowPos, col=colPos
        )
        figure.update_yaxes(
            title_text=f"{fSetup['title']}[{fSetup['units']}]",  
            row=rowPos, col=colPos
        )

def colorMap(projects: List[str]) -> Dict[str, List[str]]:
    colors = ['#e6194B', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#42d4f4', '#f032e6', '#bfef45', '#fabed4', '#469990', '#dcbeff', '#9A6324', '#fffac8', '#800000', '#aaffc3', '#808000', '#ffd8b1', '#000075', '#a9a9a9', '#000000']
    cMap = dict()
    if len(projects) > 2:
        i = 0
        for p in projects:
            cMap[p] = [colors[i % len(colors)]] * 3
            i += 1
        return cMap
    else:
        i = 0
        for p in projects:
            cMap[p] = colors[i:i+3]
            i += 3
    return cMap

def main() -> None:
    config = readConfig()
    figureSetup = readFigureSetup(config.figureSetupfilePath)
    cases, allProjects = mapResultFiles(config.simDataDirs)
    cMap = colorMap(allProjects)

    nfig = len(figureSetup)
    nrows = (nfig + config.columns - nfig%config.columns)//config.columns

    for caseId in cases.keys():
        pPath = config.resultsDir
        if not exists(pPath):
            makedirs(pPath)

        figurePath = join(pPath, str(caseId))
        figure = initFigure(figurePath, config.columns, nrows)

        for simData in cases[caseId]:
            typ = simData[0]
            project = simData[1]
            path = simData[2]
            print(f"Plotting {path} in case {caseId}.")

            if typ == 0:
                result = pd.read_csv(path,sep=';',decimal=',',header=1)
            elif typ == 1:
                result = loadEMT(path)
            addResultToFig(typ, result, figureSetup, figure, project, path, cMap, config.columns)

        if config.genHTML:
            figure.write_html('{}.html'.format(figurePath))
            
        if config.genJPEG:
            figure.write_image('{}.jpeg'.format(figurePath), width=500*nrows, height=500*config.columns)

if __name__ == "__main__":
    main()