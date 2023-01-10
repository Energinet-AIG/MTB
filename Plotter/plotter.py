"""
Information concerning the use of Power Plant and Model Test Bench

Energinet provides the Power Plant and Model Test Bench (PP-MTB) for the purpose of developing a prequalification test bench for production facility and simulation performance which the facility owner may use in its own simulation environment in order to pre-test compliance with the applicable technical requirements for simulation models. The PP-MTB are provided under the following considerations:
-	Use of the PP-MTB and its results are indicative and for informational purposes only. Energinet may only in its own simulation environment perform conclusive testing, performance and compliance of the simulation models developed and supplied by the facility owner.
-	Downloading the PP-MTB and updating the PP-MTB must only be done through a Energinet provided link. Users of the PP-MTB must not share the PP-MTB with other facility owners. The facility owner should always use the latest version of the PP-MTB in order to get the most correct results. 
-	Use of the PP-MTB are at the facility owners and the users own risk. Energinet is not responsible for any damage to hardware or software, including simulation models or computers.
-	All intellectual property rights, including copyright to the PP-MTB remains at Energinet in accordance with applicable Danish law.
"""

from os import listdir, makedirs
from os.path import join, split, splitext, exists
import re
import csv
import pandas as pd
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from configparser import ConfigParser
from types import SimpleNamespace

def idFile(fileName : str, typ : int) -> tuple:
    h, f = split(fileName)
    pat = re.compile(r'^(\w+?)_([0-9]+).' + (r'inf' if typ else r'csv') + r'$')
    rem = re.match(pat, f.lower())
    if rem:
        projectName = rem.group(1)
        caseId = int(rem.group(2))
        return (projectName, caseId)
    return None, None

def mapResultFiles(rmsRootDir : str, emtRootDir : str) -> dict:
    resultFiles = []
    resultFiles.append(listdir(rmsRootDir))
    resultFiles.append(listdir(emtRootDir))

    projects = dict()

    for typ in [0,1]:
        for fn in resultFiles[typ]:
            projectName, caseId = idFile(fn, typ)
            if projectName is None:
                continue
            if not projectName in projects.keys():
                projects[projectName] = dict()
            if not caseId in projects[projectName].keys():
                projects[projectName][caseId] = [None,None]
            projects[projectName][caseId][typ] = join(rmsRootDir if not typ else emtRootDir, fn)

    return projects

def readFigureSetup(setupFile : str) -> list:
    setup = list()
    with open(setupFile, newline='') as setupFile:
        setupReader = csv.DictReader(setupFile, delimiter = ';')
        for row in setupReader:
            setup.append(row)
    return setup

def generateFigure(rmsResult : str, emtResult : str, figureSetup : list, figureName : str, startTime : float, ncolumns : int, genHtml : bool, genStatic : bool, consolPlotTyp : bool) -> None:
    print('Generating figure: {}'.format(figureName))
    
    nfig = len(figureSetup)
    colors = ['#ff0000', '#1bc41b', '#0000ff', '#808000','#00ffff','#ff00ff']
    nrows = (nfig + ncolumns - nfig%ncolumns)//ncolumns
    fig = make_subplots(rows=nrows,cols=ncolumns)

    emtEndTime = max(emtResult['time']) if emtResult is not None else -1.0
    rmsEndTime = max(rmsResult.iloc[:, 0]) if rmsResult is not None else -1.0

    ntraces = 0
    firstEMT = True
    firstRMS = True
    for fSetup in figureSetup:
        fid = int(fSetup['figure'])
        rowpos = (fid - 1)//ncolumns + 1
        colpos = (fid - 1) % ncolumns + 1
        for sig in range(1,4):
            emtSigName = fSetup['emt_signal_{}'.format(sig)]
            rmsSigName = fSetup['rms_signal_{}'.format(sig)]
            if emtResult is not None:
                if emtSigName in emtResult.columns:
                    fig.append_trace(
                        go.Scatter(
                        x=emtResult['time'],
                        y=emtResult[emtSigName],
                        line_color=colors[ntraces] if not consolPlotTyp else '#1313ad', 
                        name='EMT:{}'.format(emtSigName) if not consolPlotTyp else 'EMT',
                        legendgroup='EMT',
                        showlegend=firstEMT or (not consolPlotTyp)
                    ),
                    row=rowpos, col=colpos
                    )
                    ntraces += 1
                    firstEMT = False
                elif emtSigName != '':
                    print('Signal \'{}\'not recognized in emt resultfile'.format(emtSigName))
            if rmsResult is not None:
                if rmsSigName in rmsResult.columns:
                    fig.append_trace(
                        go.Scatter(
                        x=rmsResult.iloc[:, 0],
                        y=rmsResult[rmsSigName],
                        line_color=colors[ntraces] if not consolPlotTyp else '#ff6e00', 
                        name='RMS:{}'.format(rmsSigName) if not consolPlotTyp else 'RMS',
                        legendgroup='RMS',
                        showlegend=firstRMS or (not consolPlotTyp)
                    ),
                    row=rowpos, col=colpos
                    )
                    ntraces += 1
                    firstRMS = False
                elif rmsSigName != '':
                    print('Signal \'{}\' not recognized in rms resultfile'.format(rmsSigName))
        fig.update_xaxes(
            title_text='Time[s]',
            range=[startTime, max(rmsEndTime,emtEndTime)],  
            row=rowpos, col=colpos
        )
        fig.update_yaxes(
            title_text='{}[{}]'.format(fSetup['title'],fSetup['units']),  
            row=rowpos, col=colpos
        )
        ntraces = 0

    fig.update_layout(title_text=figureName)
    if genHtml:
        fig.write_html('{}.html'.format(figureName))
    if genStatic:
        fig.write_image('{}.jpeg'.format(figureName), width=500*nrows, height=500*ncolumns)

def readConfig() -> SimpleNamespace:
    cp = ConfigParser()
    cp.read('config.ini')
    parsedConf = cp['config']
    config = SimpleNamespace()
    config.rmsDir = parsedConf['rmsDir']
    config.emtDir = parsedConf['emtDir']
    config.resultsDir = parsedConf['resultsDir']
    config.figureSetupfilePath = parsedConf['figureSetupfilePath']
    config.columns = parsedConf.getint('columns')               
    config.startTime = parsedConf.getfloat('startTime')        
    config.genHTML = parsedConf.getboolean('genHTML')
    config.genJPEG = parsedConf.getboolean('genJPEG')
    return config

def emtColumns(file : str) -> dict:
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

def main() -> None:
    config = readConfig()
    figureSetup = readFigureSetup(config.figureSetupfilePath)
    projects = mapResultFiles(config.rmsDir, config.emtDir)
    
    for project in projects.keys():
        pPath = join(config.resultsDir, project)
        if not exists(pPath):
            makedirs(pPath)
         
        cases = list(projects[project].keys())
        cases.sort()

        for case in cases:
            figureName = join(pPath, '{}_{}'.format(project,case))
            rFiles = projects[project][case]
            compareMode = True
            if rFiles[0]: #RMS results present
                rmsResult = pd.read_csv(rFiles[0],sep=';',decimal=',',header=1)
            else:
                rmsResult = None
                compareMode = False

            if rFiles[1]: #EMT results present
                emtResult = loadEMT(rFiles[1])
            else:
                emtResult = None
                compareMode = False
                
            generateFigure(rmsResult, emtResult, figureSetup, figureName, config.startTime, config.columns, config.genHTML, config.genJPEG, compareMode)

if __name__ == "__main__":
    main()