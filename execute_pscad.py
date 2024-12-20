'''
Executes the Powerplant model testbench in PSCAD.
'''
from __future__ import annotations 
import os
import sys

try:
    LOG_FILE = open('execute_pscad.log', 'w')
except:
    print('Failed to open log file. Logging to file disabled.')
    LOG_FILE = None #type: ignore

def print(*args): #type: ignore
    '''
    Overwrites the print function to also write to a log file.
    '''
    outputString = ''.join(map(str, args)) + '\n' #type: ignore
    sys.stdout.write(outputString)
    if LOG_FILE:
        LOG_FILE.write(outputString)
        LOG_FILE.flush()

if __name__ == '__main__':
    print(sys.version)
    #Ensure right working directory
    executePath = os.path.abspath(__file__)
    executeFolder = os.path.dirname(executePath)
    os.chdir(executeFolder)
    if not executeFolder in sys.path:
        sys.path.append(executeFolder)
    print(f'CWD: {executeFolder}')
    print('sys.path:')
    for path in sys.path:
        if path != '':
            print(f'\t{path}')
    
from configparser import ConfigParser

class readConfig:
  def __init__(self) -> None:
    self.cp = ConfigParser(allow_no_value=True)
    self.cp.read('config.ini')
    self.parsedConf = self.cp['config']
    self.sheetPath = str(self.parsedConf['Casesheet path'])
    self.pythonPath = str(self.parsedConf['Python path'])
    self.volley = int(self.parsedConf['Volley'])

config = readConfig()
sys.path.append(config.pythonPath)

from datetime import datetime
import shutil
import psutil #type: ignore
from typing import List, Optional
import sim_interface as si
import case_setup as cs
from pscad_update_ums import updateUMs

import mhi.pscad

def connectPSCAD() -> mhi.pscad.PSCAD:
    pid = os.getpid()
    ports = [con.laddr.port for con in psutil.net_connections() if con.status == psutil.CONN_LISTEN and con.pid == pid] #type: ignore

    if len(ports) == 0: #type: ignore
        exit('No PSCAD listening ports found')
    elif len(ports) > 1: #type: ignore
        print('WARNING: Multiple PSCAD listening ports found. Using the first one.')
    
    return mhi.pscad.connect(port = ports[0]) #type: ignore

def outToCsv(srcPath : str, dstPath : str):
    """
    Converts PSCAD .out file into .csv file
    """
    with open(srcPath) as out, \
            open(dstPath, 'w') as csv:
        csv.writelines(','.join(line.split()) +'\n' for line in out)

def moveFiles(srcPath : str, dstPath : str, types : List[str], suffix : str = '') -> None:
    '''
    Moves files of the specified types from srcPath to dstPath.
    '''
    for file in os.listdir(srcPath):
        _, typ = os.path.splitext(file)
        if typ in types:
            shutil.move(os.path.join(srcPath, file), os.path.join(dstPath, file + suffix))

def taskIdToRank(csvPath : str, projectName : str, emtCases : List[cs.Case]):
    '''
    Changes task ID to rank in the .csv and .inf files in csvPath.
    '''
    for file in os.listdir(csvPath):
        _, fileName = os.path.split(file)
        root, typ = os.path.splitext(fileName)
        if typ == '.csv_taskid' or typ == '.inf_taskid' and root.startswith(projectName + '_'):
            suffix = root[len(projectName) + 1:]
            parts = suffix.split('_')
            if  len(parts) > 0 and parts[0].isnumeric():
                taskId = int(parts[0])
                if taskId - 1 < len(emtCases):
                    parts[0] = str(emtCases[taskId  - 1].rank)
                    newName = projectName + '_' + '_'.join(parts) + typ.replace('_taskid', '')
                    print(f'Renaming {fileName} to {newName}')
                    os.rename(os.path.join(csvPath, fileName), os.path.join(csvPath, newName))
                else:
                    print(f'WARNING: {fileName} has a task ID that is out of bounds. Ignoring file.')
            else:
                print(f'WARNING: {fileName} has an invalid task ID. Ignoring file.')

def cleanUpOutFiles(buildPath : str, projectName : str) -> str:
    '''
    Cleans up the build folder by moving .out and .csv files to an 'Output' folder in the current working directory.
    Return path to results folder.
    '''
    # Converting .out files to .csv files
    for file in os.listdir(buildPath):
        _, fileName = os.path.split(file)
        root, typ = os.path.splitext(fileName)
        if fileName.startswith(projectName) and typ == '.out':
            print(f'Converting {file} to .csv')
            outToCsv(os.path.join(buildPath, fileName), os.path.join(buildPath, f'{root}.csv'))

    # Move desired files to an 'Output' folder in the current working directory
    outputFolder = 'output'

    if not os.path.exists(outputFolder):
        os.mkdir(outputFolder)
    else:
        for dir in os.listdir(outputFolder):
            _dir = os.path.join(outputFolder, dir)
            if os.path.isdir(_dir) and dir.startswith('MTB_'):
                if os.listdir(_dir) == []:
                    shutil.rmtree(_dir)

    resultsFolder = f'MTB_{datetime.now().strftime(r"%d%m%Y%H%M%S")}'

    #Move .csv and .inf files away from build folder into output folder
    csvFolder = os.path.join(outputFolder, resultsFolder)
    os.mkdir(csvFolder)
    moveFiles(buildPath, csvFolder, ['.csv', '.inf'], '_taskid')

    #Move .out file away from build folder
    outFolder = os.path.join(buildPath, resultsFolder)
    os.mkdir(outFolder)
    moveFiles(buildPath, outFolder, ['.out'])

    return csvFolder

def cleanBuildfolder(buildPath : str):
    '''
    "Cleans" the build folder by trying to delete it.
    '''
    try:
        shutil.rmtree(buildPath)
    except FileNotFoundError:
        pass

def findMTB(pscad : mhi.pscad.PSCAD) -> mhi.pscad.UserCmp:
    '''
    Finds the MTB block in the project.
    '''
    projectLst = pscad.projects()
    MTBcand : Optional[mhi.pscad.UserCmp] = None
    for prjDic in projectLst:
        if prjDic['type'].lower() == 'case':
            project = pscad.project(prjDic['name'])
            MTBs : List[mhi.pscad.UserCmp]= project.find_all(Name_='$MTB_9124$') #type: ignore
            if len(MTBs) > 0:
                if MTBcand or len(MTBs) > 1:
                    exit('Multiple MTB blocks found in workspace.')
                else:
                    MTBcand = MTBs[0]

    if not MTBcand:
        exit('No MTB block found in workspace.')
    return MTBcand

def addInterfaceFile(project : mhi.pscad.Project):
    '''
    Adds the interface file to the project.
    '''
    resList = project.resources()
    for res in resList:
        if res.path == r'.\interface.f' or res.name == 'interface.f':
            return

    print('Adding interface.f to project')
    project.create_resource(r'.\interface.f')

def main():
    print()
    print('execute_pscad.py started at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'), '\n')
    pscad = connectPSCAD()

    plantSettings, channels, _, _, emtCases = cs.setup(config.sheetPath, pscad = True, pfEncapsulation = None)

    #Print plant settings from casesheet
    print('Plant settings:')
    for setting in plantSettings.__dict__:
        print(f'{setting} : {plantSettings.__dict__[setting]}')
    print()
    
    #Set MTB to volley mode
    MTB = findMTB(pscad)
    project = pscad.project(MTB.project_name)
    caseList = []
    for case in emtCases:
        caseList.append(case.rank)
    
    if MTB.parameters()['par_mode'] == 'MANUAL' and MTB.parameters()['par_manualrank'] in caseList:
        #Output rank in relation to task id
        singleRank = MTB.parameters()['par_manualrank']
        singleName = emtCases[MTB.parameters()['par_manualrank']].Name
        print(f'Excecuting only Rank {singleRank}: {singleName}')
    else:
        #Set MTB to volley mode if rank does not have a corresponding case
        if MTB.parameters()['par_mode'] == 'MANUAL':
            print(f'Setting MTB to volley mode since specified rank does not have a corresponding case in testcases.')
            MTB.parameters(par_mode = 1) #type: ignore
            print()
        #Output ranks in relation to task id
        print('Rank / Task ID / Casename:')
        for case in emtCases:
            print(f'{case.rank} / {emtCases.index(case) + 1} / {case.Name}')

    print()
    si.renderFortran('interface.f', channels)
    
    #Set executed flag
    MTB.parameters(executed = 1) #type: ignore  

    #Update pgb names for all unit measurement components
    updateUMs(pscad)

    #Add interface file to project
    addInterfaceFile(project)

    buildFolder : str = project.temp_folder #type: ignore
    cleanBuildfolder(buildFolder) #type: ignore

    project.parameters(time_duration = 999, time_step = plantSettings.PSCAD_Timestep, sample_step = '1000') #type: ignore
    project.parameters(PlotType = '1', output_filename = f'{plantSettings.Projectname}.out') #type: ignore
    project.parameters(SnapType='0', SnapTime='2', snapshot_filename='pannatest5us.snp') #type: ignore

    pscad.remove_all_simulation_sets()
    pmr = pscad.create_simulation_set('MTB')
    pmr.add_tasks(MTB.project_name)
    project_pmr = pmr.task(MTB.project_name)
    project_pmr.parameters(ammunition = len(emtCases) if MTB.parameters()['par_mode'] == 'VOLLEY' else 1 , volley = config.volley, affinity_type = '2') #type: ignore

    pscad.run_simulation_sets('MTB') #type: ignore ??? By sideeffect changes current working directory ???
    os.chdir(executeFolder)

    csvFolder = cleanUpOutFiles(buildFolder, plantSettings.Projectname)
    print()
    taskIdToRank(csvFolder, plantSettings.Projectname, emtCases)

    print('execute.py finished at: ', datetime.now().strftime('%m-%d %H:%M:%S'))

if __name__ == '__main__':
    main()

if LOG_FILE:
    LOG_FILE.close()
