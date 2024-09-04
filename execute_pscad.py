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
    sys.path.append(executeFolder)
    print(executeFolder)

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
import psutil
from typing import List
import sim_interface as si
import case_setup as cs
from pscad_update_ums import updateUMs

import mhi.pscad

def connectPSCAD() -> mhi.pscad.PSCAD:
    pid = os.getpid()
    ports = [con.laddr.port for con in psutil.net_connections() if con.status == psutil.CONN_LISTEN and con.pid == pid] #type: ignore

    if len(ports) == 0: #type: ignore
        RuntimeError('No PSCAD listening ports found')
    elif len(ports) > 1: #type: ignore
        Warning('Multiple PSCAD listening ports found. Using the first one.')
    
    return mhi.pscad.connect(port = ports[0]) #type: ignore

def outToCsv(srcPath : str, dstPath : str):
    """
    Converts PSCAD .out file into .csv file
    """
    with open(srcPath) as out, \
            open(dstPath, 'w') as csv:
        csv.writelines(','.join(line.split()) +'\n' for line in out)

def moveFiles(srcPath : str, dstPath : str, types : List[str]) -> None:
    '''
    Moves files of the specified types from srcPath to dstPath.
    '''
    for file in os.listdir(srcPath):
        _, typ = os.path.splitext(file)
        if typ in types:
            shutil.move(os.path.join(srcPath, file), os.path.join(dstPath, file))

def taskIdToRank(csvPath : str, projectName : str, emtCases : List[cs.Case]):
    '''
    Changes task ID to rank in the .csv and .inf files in csvPath.
    '''
    for file in os.listdir(csvPath):
        _, fileName = os.path.split(file)
        root, typ = os.path.splitext(fileName)
        if typ == '.csv' or typ == '.inf':
            parts = root.split('_')
            if len(parts) > 1 and parts[0] == projectName and parts[1].isnumeric():
                taskId = int(parts[1])
                if taskId - 1 < len(emtCases):
                    parts[1] = str(emtCases[int(parts[1]) - 1].rank)
                    newName = '_'.join(parts)
                    print(f'Renaming {fileName} to {newName + typ}')
                    os.rename(os.path.join(csvPath, fileName), os.path.join(csvPath, newName + typ))
                else:
                    Warning(f'{fileName} has a task ID that is out of bounds. Ignoring file.')

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
    moveFiles(buildPath, csvFolder, ['.csv', '.inf'])

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

def setMTBtoVolley(project : mhi.pscad.Project):
    '''
    Sets MTB block to volley mode.
    '''
    MTBs : List[mhi.pscad.UserCmp]= project.find_all(Name_ = '$MTB_9124$') #type: ignore
    for MTB in MTBs:
        print(f'Setting {MTB} to volley mode')
        MTB.parameters(par_mode = 1)

def addInterfaceFile(project : mhi.pscad.Project):
    '''
    Adds the interface file to the project.
    '''
    resList = project.resources()
    for res in resList:
        if res.path == '.\interface.f' or res.name == 'interface.f':
            return

    print('Adding interface.f to project')
    project.create_resource('.\interface.f')

def main():
    print('execute_pscad.py started at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'), '\n')
    pscad = connectPSCAD()

    plantSettings, channels, _, _, emtCases = cs.setup(config.sheetPath, pscad = True, pfEncapsulation = None)

    #Output ranks in relation to task id
    print('Rank / Task ID / Casename:')
    for case in emtCases:
        print(f'{case.rank} / {emtCases.index(case) + 1} / {case.Name}')

    print()
    si.renderFortran('interface.f', channels)

    #Print plant settings from casesheet
    print('Plant settings:')
    for setting in plantSettings.__dict__:
        print(f'{setting} : {plantSettings.__dict__[setting]}')
    print()
    project = pscad.project(plantSettings.PSCAD_Namespace)
    
    #Update pgb names for all unit measurement components
    updateUMs(pscad)

    #Set MTB to volley mode
    setMTBtoVolley(project)

    #Add interface file to project
    addInterfaceFile(project)

    buildFolder : str = project.temp_folder #type: ignore
    cleanBuildfolder(buildFolder) #type: ignore

    project.parameters(time_duration = 999, time_step = plantSettings.PSCAD_Timestep, sample_step = '1000') #type: ignore
    project.parameters(PlotType = '1', output_filename = f'{plantSettings.Projectname}.out') #type: ignore
    project.parameters(SnapType='0', SnapTime='2', snapshot_filename='pannatest5us.snp') #type: ignore

    pscad.remove_all_simulation_sets()
    pmr = pscad.create_simulation_set('MTB')
    pmr.add_tasks(plantSettings.PSCAD_Namespace)
    project_pmr = pmr.task(plantSettings.PSCAD_Namespace)
    project_pmr.parameters(ammunition = len(emtCases), volley = config.volley, affinity_type = '2') #type: ignore

    pscad.run_simulation_sets('MTB') #type: ignore ??? By sideeffect changes current working directory ???
    os.chdir(executeFolder)

    csvFolder = cleanUpOutFiles(buildFolder, plantSettings.Projectname)
    print()
    taskIdToRank(csvFolder, plantSettings.Projectname, emtCases)

    print('execute.py finished at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

if __name__ == '__main__':
    main()

if LOG_FILE:
    LOG_FILE.close()