'''
Minimal script to plot simulation results from PSCAD and PowerFactory.
'''
from os import listdir, makedirs
from os.path import join, split, splitext, exists
import re
import csv
import pandas as pd
from plotly.subplots import make_subplots #type: ignore
import plotly.graph_objects as go #type: ignore
from configparser import ConfigParser
from typing import List, Dict, Union, Tuple, Set

from threading import Thread
import time

class ReadConfig:
    def __init__(self) -> None:
        cp = ConfigParser()
        cp.read('config.ini')
        parsedConf = cp['config']
        self.resultsDir = parsedConf['resultsDir']
        self.figureSetupfilePath = parsedConf['figureSetupfilePath']
        self.columns = parsedConf.getint('columns')                 
        self.genHTML = parsedConf.getboolean('genHTML')
        self.genJPEG = parsedConf.getboolean('genJPEG')
        self.emtAndRms = parsedConf.getboolean('emtAndRms') 
        self.emtMinTime = parsedConf.getfloat('emtMinTime')
        self.threads = parsedConf.getint('threads')
        self.pfFlatTIme = parsedConf.getfloat('pfFlatTime')
        self.pscadInitTime = parsedConf.getfloat('pscadInitTime')

        self.simDataDirs : List[str] = list()
        simPaths = cp.items('Simulation data paths')
        for _, path in simPaths:
            self.simDataDirs.append(path)

def idFile(filePath: str) -> Tuple[Union[int, None], Union[str, None], Union[int, None]]:
    '''
    Identifies the type (EMT or RMS), project and case id of a given file. If the file is not recognized, a none tuple is returned.
    '''
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

def mapResultFiles(dirs: List[str]) -> Tuple[Dict[int, List[Tuple[int, str, str]]], Set[str]]:
    '''
    Goes through all files in the given directories and maps them to a dictionary of cases.
    '''
    files = [join(dir, p) for dir in dirs for p in listdir(dir)]

    cases : Dict[int, List[Tuple[int, str, str]]] = {}
    relevantProjects : Set[str] = set()

    for file in files:
        typ, project, id = idFile(file)
        if typ is None:
            continue
        
        assert project is not None
        assert id is not None

        if typ == 0:
            project = f"{project}_RMS"
        elif typ == 1:
            project = f"{project}_EMT"

        cases.setdefault(id, []).append((typ, project, file))

        relevantProjects.add(project)        
    return cases, relevantProjects

def readFigureSetup(filePath : str) -> List[Dict[str, str]]:
    '''
    Reads the figure setup from the given file and returns a list of dictionaries containing the information.
    The index of the returned list corresponds to the figure number.
    '''
    setup : List[str]= list()
    with open(filePath, newline='') as setupFile:
        setupReader = csv.DictReader(setupFile, delimiter = ';')
        for row in setupReader:
            setup.append(row) #type: ignore
    return setup #type: ignore

def emtColumns(infFilePath : str) -> Dict[int, str]:
    '''
    Reads EMT result columns from the given inf file and returns a dictionary with the column number as key and the column name as value.
    '''
    columns : Dict[int, str] = dict()
    with open(infFilePath, 'r') as file:
        for line in file: 
            rem = re.match(r'^PGB\(([0-9]+)\) +Output +Desc="(\w+)" +Group="(\w+)" +Max=([0-9\-\.]+) +Min=([0-9\-\.]+) +Units="(\w*)" *$', line)
            if rem:
                columns[int(rem.group(1))] = rem.group(2) 
    return columns

def loadEMT(infFile : str) -> pd.DataFrame:
    '''
    Load EMT results from a collection of csv files defined by the given inf file. Returns a dataframe with index 'time'.
    '''
    folder, filename = split(infFile)
    filename, fileext = splitext(filename)
 
    assert fileext.lower() == '.inf'

    adjFiles = listdir(folder)
    csvMap : Dict[int, str]= dict()
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
        dfMap = pd.read_csv(csvMap[map], skiprows = 1, header = None) #type: ignore
        if not firstFile:
            dfMap = dfMap.iloc[: , 1:]
        else:
            firstFile = False
        dfMap.columns = list(range(loadedColumns, loadedColumns + len(dfMap.columns)))
        loadedColumns = loadedColumns + len(dfMap.columns)
        df = pd.concat([df, dfMap], axis=1) #type: ignore

    columns = emtColumns(infFile)
    columns[0] = 'time'
    df = df[columns.keys()]
    df.rename(columns, inplace=True, axis=1)
    print(f"Loaded {infFile}, length = {df['time'].iloc[-1]}s") #type: ignore   
    return df

def addResultToFig(typ: int, result: pd.DataFrame, figureSetup: List[Dict[str, str]], figure : go.Figure, project: str, file: str, colors: Dict[str, List[str]], nColumns: int, pfFlatTIme : float, pscadInitTime : float) -> None: 
    for fSetup in figureSetup:
        fid = int(fSetup['figure'])
        rowPos = (fid - 1) // nColumns + 1
        colPos = (fid - 1) % nColumns + 1
        traces = 0

        for sig in range(1,4):
            signalKey = 'rms' if typ == 0 else 'emt'
            rawSigName = fSetup.get(f"{signalKey}_signal_{sig}", "")
            
            if typ == 0:
                splitSigName = rawSigName.split('\\')
                
                if len(splitSigName) == 2:
                    sigColumn = (splitSigName[0], splitSigName[1])
                else:
                    sigColumn = ''
            else:
                sigColumn = rawSigName

            timeColName = 'time' if typ == 1 else ('Results','b:tnow in s')
            timeoffset = pfFlatTIme if typ == 0 else pscadInitTime    

            if sigColumn in result.columns:
                figure.append_trace( #type: ignore
                    go.Scatter(
                    x=result[timeColName] - timeoffset,
                    y=result[sigColumn],
                    line_color=colors[project][traces], 
                    name=f"{file}:{rawSigName}",
                    legendgroup=project,
                    showlegend=True
                ),
                row=rowPos, col=colPos
                )
                traces += 1
            elif sigColumn != '':
                print(f"Signal '{rawSigName}' not recognized in resultfile '{file}'")
                figure.append_trace( #type: ignore
                    go.Scatter(
                    x=None,
                    y=None,
                    line_color=colors[project][traces], 
                    name=f"{file}:{rawSigName} (Unknown)",
                    legendgroup=project,
                    showlegend=True
                ),
                row=rowPos, col=colPos
                )
                traces += 1               
        
        figure.update_xaxes( #type: ignore
            title_text='Time[s]',  
            row=rowPos, col=colPos
        )
        figure.update_yaxes( #type: ignore
            title_text=f"{fSetup['title']}[{fSetup['units']}]",  
            row=rowPos, col=colPos
        )

def colorMap(projects: List[str]) -> Dict[str, List[str]]:
    '''
    Select colors for the given projects. Return a dictionary with the project name as key and a list of colors as value.
    '''
    colors = ['#e6194B', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#42d4f4', '#f032e6', '#bfef45', '#fabed4', '#469990', '#dcbeff', '#9A6324', '#fffac8', '#800000', '#aaffc3', '#808000', '#ffd8b1', '#000075', '#a9a9a9', '#000000']
    cMap : Dict[str, List[str]] = dict()
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

def drawFigure(figurePath : str, config : ReadConfig, nrows : int, cases : Dict[int, List[Tuple[int, str, str]]], caseId : int, figureSetup :  List[Dict[str, str]], cMap : Dict[str, List[str]]):
    figure = make_subplots(rows = nrows, cols = config.columns)
    figure.update_layout(title_text = figurePath) #type: ignore

    addedRmsResults = 0
    addedEmtResults = 0

    for typ, project, path in cases[caseId]:
        print(f"Plotting {path} in case {caseId}.")

        if typ == 0:
            result : pd.DataFrame = pd.read_csv(path,sep=';',decimal=',',header=[0,1]) #type: ignore
            addedRmsResults += 1
        elif typ == 1:
            result = loadEMT(path,)
            if result['time'].iloc[-1] < config.emtMinTime: #type: ignore
                print(f"Resultfile '{path}' is too short. Skipping.")
                continue
            addedEmtResults += 1
        
        addResultToFig(typ, result, figureSetup, figure, project, path, cMap, config.columns, config.pfFlatTIme, config.pscadInitTime) #type: ignore


    if config.emtAndRms and addedRmsResults > 0 and addedEmtResults > 0 or not config.emtAndRms and ( addedRmsResults > 0 or addedEmtResults > 0):
        if config.genHTML:
            figure.write_html('{}.html'.format(figurePath)) #type: ignore
            
        if config.genJPEG:
            figure.write_image('{}.jpeg'.format(figurePath), width=500*nrows, height=500*config.columns) #type: ignore

def main() -> None:
    config = ReadConfig()
    figureSetup = readFigureSetup(config.figureSetupfilePath)
    cases, allProjects = mapResultFiles(config.simDataDirs)
    cMap = colorMap(list(allProjects))

    if not exists(config.resultsDir):
        makedirs(config.resultsDir)

    nfig = len(figureSetup)
    nrows = (nfig + config.columns - nfig%config.columns)//config.columns

    threads : List[Thread] = list()

    for caseId in cases.keys():
        types = 0
        for simData in cases[caseId]:
            types += simData[0]

        if len(cases[caseId]) > types > 0 or not config.emtAndRms:
            figurePath = join(config.resultsDir, str(caseId))
            if config.threads > 1:
                threads.append(Thread(target = drawFigure, args = (figurePath, config, nrows, cases, caseId, figureSetup, cMap,)))
            else:
                drawFigure(figurePath, config, nrows, cases, caseId, figureSetup, cMap)
    
    NoT = len(threads)
    if NoT > 0:  
        sched = threads.copy()
        inProg : List[Thread] = []

        while len(sched) > 0:
            for t in inProg:
                if not t.is_alive():
                    inProg.remove(t)

            while len(inProg) < config.threads and len(sched) > 0:
                nextThread = sched.pop()
                nextThread.start()
                inProg.append(nextThread)

            time.sleep(0.5)

if __name__ == "__main__":
    main()