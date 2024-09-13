'''
Minimal script to plot simulation results from PSCAD and PowerFactory.
'''
from __future__ import annotations
from os import listdir, makedirs
from os.path import join, split, splitext, exists
import re
import csv
import pandas as pd
from plotly.subplots import make_subplots #type: ignore
import plotly.graph_objects as go #type: ignore
from configparser import ConfigParser
from typing import List, Dict, Union, Tuple, Set
import sampling_functions
from down_sampling_method import DownSamplingMethod
from threading import Thread, Lock
import time
import sys
from enum import Enum
from math import ceil
from collections import defaultdict

try:
    LOG_FILE = open('plotter.log', 'w')
except:
    print('Failed to open log file. Logging to file disabled.')
    LOG_FILE = None #type: ignore

gLock = Lock()

def print(*args): #type: ignore
    '''
    Overwrites the print function to also write to a log file.
    '''
    gLock.acquire()
    outputString = ''.join(map(str, args)) + '\n' #type: ignore
    sys.stdout.write(outputString)
    if LOG_FILE:
        try:
            LOG_FILE.write(outputString)
            LOG_FILE.flush()
        except:
            pass
    gLock.release()

class ResultType(Enum):
    RMS = 0
    EMT = 1

    @classmethod
    def from_string(cls, string : str):
        try:
            return cls[string.upper()]
        except KeyError:
            raise ValueError(f'{string} is not a valid {cls.__name__}')

class Figure:
    def __init__(self, 
                 id : int, 
                 title : str, 
                 units : str, 
                 emt_signal_1 : str, 
                 emt_signal_2 : str, 
                 emt_signal_3 : str, 
                 rms_signal_1 : str, 
                 rms_signal_2 : str, 
                 rms_signal_3 : str, 
                 gradient_threshold : float, 
                 down_sampling_method : DownSamplingMethod,
                 include_in_case : List[int],
                 exclude_in_case : List[int]) -> None:
        
        self.id = id
        self.title = title
        self.units = units
        self.emt_signal_1 = emt_signal_1
        self.emt_signal_2 = emt_signal_2
        self.emt_signal_3 = emt_signal_3
        self.rms_signal_1 = rms_signal_1
        self.rms_signal_2 = rms_signal_2
        self.rms_signal_3 = rms_signal_3
        self.gradient_threshold = float(gradient_threshold)
        self.down_sampling_method = down_sampling_method
        self.include_in_case : List[int] = include_in_case
        self.exclude_in_case : List[int] = exclude_in_case

class Result:
    def __init__(self, typ : ResultType, rank : int, name : str, bulkname : str, fullpath : str, group : str) -> None:
        self.typ = typ
        self.rank = rank
        self.name = name
        self.bulkname = bulkname
        self.fullpath = fullpath
        self.group = group

class ReadConfig:
    def __init__(self) -> None:
        cp = ConfigParser()
        cp.read('config.ini')
        parsedConf = cp['config']
        self.resultsDir = parsedConf['resultsDir']
        self.columns = parsedConf.getint('columns')
        assert self.columns > 0                 
        self.genHTML = parsedConf.getboolean('genHTML')
        self.genImage = parsedConf.getboolean('genImage')
        self.imageFormat = parsedConf['imageFormat']
        self.threads = parsedConf.getint('threads')
        assert self.threads > 0
        self.pfFlatTIme = parsedConf.getfloat('pfFlatTime')
        assert self.pfFlatTIme >= 0.1
        self.pscadInitTime = parsedConf.getfloat('pscadInitTime')
        assert self.pscadInitTime >= 1.0
        self.optionalCasesheet = parsedConf['optionalCasesheet']
        self.simDataDirs : List[Tuple[str, str]] = list()
        simPaths = cp.items('Simulation data paths')
        for name, path in simPaths:
            self.simDataDirs.append((name, path))

class Case:
    def __init__(self, case: 'pd.Series[Union[str, int, float, bool]]') -> None:
        self.rank: int = int(case['Rank'])
        self.RMS: bool = bool(case['RMS'])
        self.EMT: bool = bool(case['EMT'])
        self.Name: str = str(case['Name'])
        self.U0: float = float(case['U0'])
        self.P0: float = float(case['P0'])
        self.Pmode: str = str(case['Pmode'])
        self.Qmode: str = str(case['Qmode'])
        self.Qref0: float = float(case['Qref0'])
        self.SCR0: float = float(case['SCR0'])
        self.XR0: float = float(case['XR0'])
        self.Simulationtime: float = float(case['Simulationtime'])
        self.Events : List[Tuple[str, float, Union[float, str], Union[float, str]]] = []

        index : pd.Index[str] = case.index # type: ignore
        i = 0
        while(True):
            typeLabel = f'type.{i}' if i > 0 else 'type'
            timeLabel = f'time.{i}' if i > 0 else 'time'
            x1Label = f'X1.{i}' if i > 0 else 'X1'
            x2Label = f'X2.{i}' if i > 0 else 'X2'

            if typeLabel in index and timeLabel in index and x1Label in index and x2Label in index:
                try:
                    x1value = float(str(case[x1Label]).replace(' ',''))
                except ValueError:
                    x1value = str(case[x1Label])

                try:
                    x2value = float(str(case[x2Label]).replace(' ',''))
                except ValueError:
                    x2value = str(case[x2Label])

                self.Events.append((str(case[typeLabel]), float(case[timeLabel]), x1value, x2value))
                i += 1
            else:
                break

def readFigureSetup(filePath : str) -> Dict[int, List[Figure]]:
    '''
    Read figure setup file.
    '''
    setup : List[Dict[str, str|List[int]]] = list()
    with open(filePath, newline='') as setupFile:
        setupReader = csv.DictReader(setupFile, delimiter = ';')
        for row in setupReader:
            row['exclude_in_case'] = list(set([int(item.strip()) for item in row.get('exclude_in_case', '').split(',') if item.strip() != '']))
            row['include_in_case']  = list(set([int(item.strip())  for item in row.get('include_in_case', '').split(',') if item.strip() != '']))
            setup.append(row)
    
    figureList : List[Figure] = list()
    for figureStr in setup:
        figureList.append(
               Figure(int(figureStr['figure']),                                        #type: ignore
               figureStr['title'],                                                #type: ignore
               figureStr['units'],                                                #type: ignore
               figureStr['emt_signal_1'],                                         #type: ignore
               figureStr['emt_signal_2'],                                         #type: ignore
               figureStr['emt_signal_3'],                                         #type: ignore
               figureStr['rms_signal_1'],                                         #type: ignore
               figureStr['rms_signal_2'],                                         #type: ignore
               figureStr['rms_signal_3'],                                         #type: ignore
               figureStr['gradient_threshold'],                                   #type: ignore
               DownSamplingMethod.from_string(figureStr['down_sampling_method']), #type: ignore
               figureStr['include_in_case'],                                      #type: ignore
               figureStr['exclude_in_case']))                                     #type: ignore

    defaultSetup = [fig for fig in figureList if fig.include_in_case == []]
    figDict : Dict[int, List[Figure]] = defaultdict(lambda: defaultSetup)

    for fig in figureList:
        if fig.include_in_case != []:
            for inc in fig.include_in_case:
                if not inc in figDict.keys():
                    figDict[inc] = defaultSetup.copy()
                figDict[inc].append(fig)
        else:
            for exc in fig.exclude_in_case:
                if not exc in figDict.keys():
                    figDict[exc] = defaultSetup.copy()
                figDict[exc].remove(fig)
    return figDict

def idFile(filePath: str) -> Tuple[Union[ResultType, None], Union[int, None], Union[str, None], Union[str, None], Union[str, None]]:
    '''
    Identifies the type (EMT or RMS), root and case id of a given file. If the file is not recognized, a none tuple is returned.
    '''
    path, fileName = split(filePath)
    match = re.match(r'^(\w+?)_([0-9]+).(inf|csv)$', fileName.lower())
    if match:
        rank = int(match.group(2))
        name = match.group(1)
        bulkName = join(path, match.group(1))
        fullpath = filePath
        with open(filePath, 'r') as file:
            firstLine = file.readline()
            if match.group(3) == 'inf' and firstLine.startswith('PGB(1)'):
                fileType = ResultType.EMT
                return (fileType, rank, name, bulkName, fullpath)
            elif match.group(3) == 'csv':
                secondLine = file.readline()
                if secondLine.startswith(r'"b:tnow in s"'):
                    fileType = ResultType.RMS
                    return (fileType, rank, name, bulkName, fullpath)
    return (None, None, None, None, None)

def mapResultFiles(config : ReadConfig) -> Dict[int, List[Result]]:
    '''
    Goes through all files in the given directories and maps them to a dictionary of cases.
    '''
    files : List[Tuple[str, str]] = list()
    for dir_ in config.simDataDirs:
        for file_ in listdir(dir_[1]):
            files.append((dir_[0], join(dir_[1], file_)))

    results : Dict[int, List[Result]] = dict()

    for file in files:
        group = file[0]
        fullpath = file[1]
        typ, rank, name, bulkName, fullpath = idFile(fullpath)
        if typ is None:
            continue
        assert rank is not None
        assert name is not None
        assert bulkName is not None
        assert fullpath is not None
        newResult = Result(typ, rank, name, bulkName, fullpath, group)

        if rank in results.keys():
            results[rank].append(newResult)
        else:
            results[rank] = [newResult]

    return results

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

def addResults(    plots : List[go.Figure],
                   typ: ResultType,
                   data: pd.DataFrame,
                   figures: List[Figure],
                   name: str,
                   file: str, #Only for error messages
                   colors: Dict[str, List[str]],
                   nColumns: int,
                   pfFlatTIme : float,
                   pscadInitTime : float) -> None:
    '''
    Add result to plot.
    '''

    assert len(plots) == len(figures)
    assert nColumns > 0

    if nColumns > 1:    
        plotlyFigure = plots[0]

    rowPos = 1
    colPos = 1
    fi = -1
    for figure in figures:
        fi += 1
        
        if nColumns == 1:
            plotlyFigure = plots[fi]
        else:
            rowPos = (fi // nColumns) + 1
            colPos = (fi % nColumns) + 1
            
        downsampling_method = figure.down_sampling_method
        traces = 0
        for sig in range(1,4):
            signalKey = typ.name.lower()
            rawSigName = getattr(figure, f'{signalKey}_signal_{sig}')
            
            if typ == ResultType.RMS:
                while rawSigName.startswith('#'):
                    rawSigName = rawSigName[1:]
                splitSigName = rawSigName.split('\\')

                if len(splitSigName) == 2:
                    sigColumn = ('##' + splitSigName[0], splitSigName[1])
                else:
                    sigColumn = rawSigName
            else:
                sigColumn = rawSigName

            timeColName = 'time' if typ == ResultType.EMT else data.columns[0]
            timeoffset = pfFlatTIme if typ == ResultType.RMS else pscadInitTime

            if sigColumn in data.columns:
                x_value = data[timeColName] - timeoffset #type: ignore
                y_value = data[sigColumn] #type: ignore
                if downsampling_method == DownSamplingMethod.GRADIENT:
                    x_value, y_value = sampling_functions.downsample_based_on_gradient(x_value, y_value, figure.gradient_threshold) #type: ignore
                elif downsampling_method == DownSamplingMethod.AMOUNT:
                    x_value, y_value = sampling_functions.down_sample(x_value, y_value) #type: ignore

                plotlyFigure.add_trace( #type: ignore
                    go.Scatter(
                        x=x_value,
                        y=y_value,
                        line_color=colors[name][traces],
                        name=f"{name}:{rawSigName}",
                        legendgroup=name,
                        showlegend=True
                    ),
                    row=rowPos, col=colPos
                )  
                traces += 1
            elif sigColumn != '':
                print(f"Signal '{rawSigName}' not recognized in resultfile '{file}'")
                plotlyFigure.add_trace( #type: ignore
                    go.Scatter(
                    x=None,
                    y=None,
                    line_color=colors[name][traces], 
                    name=f"{name}:{rawSigName} (Unknown)",
                    legendgroup=name,
                    showlegend=True
                ),
                row=rowPos, col=colPos
                )
                traces += 1               
        
        plotlyFigure.update_xaxes( #type: ignore
            title_text='Time[s]',  
            row=rowPos, col=colPos
        )
        if nColumns == 1:
            yaxisTitle = f"[{figure.units}]"
        else:
            yaxisTitle = f"{figure.title}[{figure.units}]"

        plotlyFigure.update_yaxes( #type: ignore
            title_text=yaxisTitle ,  
            row=rowPos, col=colPos
        )

def colorMap(names: List[str]) -> Dict[str, List[str]]:
    '''
    Select colors for the given projects. Return a dictionary with the project name as key and a list of colors as value.
    '''
    colors = ['#e6194B', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#42d4f4', '#f032e6', '#bfef45', '#fabed4', '#469990', '#dcbeff', '#9A6324', '#fffac8', '#800000', '#aaffc3', '#808000', '#ffd8b1', '#000075', '#a9a9a9', '#000000']
    cMap : Dict[str, List[str]] = dict()
    if len(names) > 2:
        i = 0
        for p in names:
            cMap[p] = [colors[i % len(colors)]] * 3
            i += 1
        return cMap
    else:
        i = 0
        for p in names:
            cMap[p] = colors[i:i+3]
            i += 3
    return cMap

def drawPlot( rank : int,
                resultDict : Dict[int, List[Result]],
                figureDict : Dict[int, List[Figure]],
                caseDict : Dict[int, str],
                colorMap : Dict[str, List[str]],
                config : ReadConfig):
    
    '''
    Draws plot.    
    '''

    print(f'Drawing plot for rank {rank}.')

    resultList = resultDict.get(rank, [])
    figureList = figureDict[rank]

    if resultList == [] or figureList == []:
        return

    figurePath = join(config.resultsDir, str(rank))

    plot : List[go.Figure] = list()

    if config.columns == 1:
         for fig in figureList:
            plot.append(make_subplots())
            plot[-1].update_layout(title_text = fig.title, height = 500, width = 1920, #type: ignore
                                   legend=dict(
                                        orientation="h",
                                        yanchor="top",
                                        y=1.3,
                                        xanchor="left",
                                        x = 0.2,
                                    )) 
    elif config.columns > 1:
        plot.append(make_subplots(rows = ceil(len(figureList)/config.columns), cols = config.columns))
        plot[-1].update_layout(title_text = caseDict[rank], height = 500 * ceil(len(figureList)/config.columns), width = 1920/config.columns) #type: ignore
    
    for result in resultList:
        if result.typ == ResultType.RMS:
            resultData : pd.DataFrame = pd.read_csv(result.fullpath, sep=';',decimal=',',header=[0,1]) #type: ignore
        elif result.typ == ResultType.EMT:
            resultData = loadEMT(result.fullpath)
        else:
            continue

        addResults(plot, result.typ, resultData, figureList, result.group, result.fullpath, colorMap, config.columns, config.pfFlatTIme, config.pscadInitTime)
   
    if config.genHTML:
        create_html(plot, figurePath, config, caseDict[rank])
        print(f'Exported plot for rank {rank} to {figurePath}.html')
        
    #if config.genImage: 
    #    plot.write_image(f'{figurePath}.{config.imageFormat}', height=500*nrows, width=500*config.columns) #type: ignore
    #    print(f'Exported plot for rank {rank} to {figurePath}.{config.imageFormat}')

    print(f'Plot for rank {rank} done.')

def create_html(plots : List[go.Figure], path : str, config : ReadConfig, title : str):

    additional_text =  """
            <div style="text-align: left; margin-top: 1px;">"""

    for group in config.simDataDirs:
        additional_text += f"<p>{group[0]} = {group[1]}</p>"

    additional_text += """<p><center><a href="https://github.com/Energinet-AIG/MTB">Generated with Energinets Model Testbench</a></center></p></div>"""

    if config.columns == 1:
        html_content = '<h1>' + title + '</h1>'
    else:
        html_content = ''
    for p in plots:
        html_content += p.to_html(full_html=False, include_plotlyjs='cdn') #type: ignore

    full_html_content = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
            </head>
            <body>
                {html_content}
                {additional_text}
            </body>
            </html>
            """
    
    with open(f'{path}.html', 'w') as file:
        file.write(full_html_content)

def readCasesheet(casesheetPath : str) -> Dict[int, str]:
    '''
    Reads optional casesheets and provides dict mapping rank to case title.
    '''

    try:
        pd.read_excel(casesheetPath, sheet_name='RfG cases', header=1) # type: ignore
    except FileNotFoundError:
        print(f'Casesheet not found at {casesheetPath}.')
        return dict()
    
    cases : List[Case] = list()
    for sheet in ['RfG', 'DCC', 'Unit', 'Custom']:
        dfc = pd.read_excel(casesheetPath, sheet_name=f'{sheet} cases', header=1) # type: ignore
        for _, case in dfc.iterrows(): # type: ignore
            cases.append(Case(case)) # type: ignore

    caseDict : Dict[int, str] = defaultdict(lambda: 'Unknown case')
    for case in cases:
        caseDict[case.rank] = case.Name
    return caseDict

def main() -> None:
    config = ReadConfig()
    
    print('Starting plotter')

    #Output config
    print('Configuration:')
    for setting in config.__dict__:
        print(f'\t{setting}: {config.__dict__[setting]}')

    print()

    resultDict = mapResultFiles(config)
    figureDict = readFigureSetup('figureSetup.csv')
    caseDict = readCasesheet(config.optionalCasesheet)

    uniqueGroups : Set[str] = set()

    for rank in resultDict.keys():
        for result in resultDict[rank]:
            uniqueGroups.add(result.group)

    colorSchemeMap = colorMap(list(uniqueGroups))

    if not exists(config.resultsDir):
        makedirs(config.resultsDir)

    threads : List[Thread] = list()

    for rank in resultDict.keys():
        if config.threads > 1:
            threads.append(Thread(target = drawPlot, args = (rank, resultDict, figureDict, caseDict, colorSchemeMap, config)))
        else:
            drawPlot(rank, resultDict, figureDict, caseDict, colorSchemeMap, config)
    
    NoT = len(threads)
    if NoT > 0:  
        sched = threads.copy()
        inProg : List[Thread] = []

        while len(sched) > 0:
            for t in inProg:
                if not t.is_alive():
                    print(f'Thread {t.native_id} finished')
                    inProg.remove(t)

            while len(inProg) < config.threads and len(sched) > 0:
                nextThread = sched.pop()
                nextThread.start()
                print(f'Started thread {nextThread.native_id}')
                inProg.append(nextThread)

            time.sleep(0.5)

    print('Finished plotter')

if __name__ == "__main__":
    main()

if LOG_FILE:
    LOG_FILE.close()