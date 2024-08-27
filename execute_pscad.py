'''
Executes the Powerplant model testbench in PSCAD.
'''
from __future__ import annotations 
import os
import sys

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
    ports = [con.laddr.port for con in psutil.net_connections() if con.status == psutil.CONN_LISTEN and con.pid == pid]

    if len(ports) == 0:
        RuntimeError('No PSCAD listening ports found')
    elif len(ports) > 1:
        Warning('Multiple PSCAD listening ports found. Using the first one.')
    
    return mhi.pscad.connect(port = ports[0])

def outToCsv(srcPath : str, dstPath : str):
    """
    Converts PSCAD .out file into .csv file
    """
    with open(srcPath) as out, \
            open(dstPath, 'w') as csv:
        csv.writelines(','.join(line.split())+'\n' for line in out)

def moveFiles(srcPath : str, dstPath : str, types : List[str]) -> None:
    '''
    Moves files of the specified types from srcPath to dstPath.
    '''
    for file in os.listdir(srcPath):
        _, typ = os.path.splitext(file)
        if typ in types:
            shutil.move(os.path.join(srcPath, file), os.path.join(dstPath, file))

def cleanUpOutFiles(buildPath : str, projectName : str):
    '''
    Cleans up the build folder by moving .out and .csv files to an 'Output' folder in the current working directory.
    '''
    # Converting .out files to .csv files
    for file in os.listdir(buildPath):
        _, fileName = os.path.split(file)
        root, typ = os.path.splitext(fileName)
        if fileName.startswith(projectName) and typ == '.out':
            print('Converting {} to .csv'.format(file))
            outToCsv(os.path.join(buildPath, fileName), os.path.join(buildPath, f'{root}.csv'))

    # Move desired files to an 'Output' folder in the current working directory
    outputFolder = 'output'

    if not os.path.exists(outputFolder):
        os.mkdir(outputFolder)

    resultsFolder = f'MTB_{datetime.now().strftime(r"%d%m%Y%H%M%S")}'

    #Move .csv and .inf files away from build folder into output folder
    csvFolder = os.path.join(outputFolder, resultsFolder)
    os.mkdir(csvFolder)
    moveFiles(buildPath, csvFolder, ['.csv', '.inf'])

    #Move .out file away from build folder
    outFolder = os.path.join(buildPath, resultsFolder)
    os.mkdir(outFolder)
    moveFiles(buildPath, outFolder, ['.out'])

def cleanBuildfolder(buildPath : str):
    '''
    "Cleans" the build folder by trying to delete it.
    '''
    try:
        shutil.rmtree(buildPath)
    except FileNotFoundError:
        pass

def main():
    print('execute_pscad.py started at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    pscad = connectPSCAD()

    plantSettings, channels, _, maxRank = cs.setup(config.sheetPath, pscad = True, pfEncapsulation = None)

    si.renderFortran('interface.f', channels)

    #Print key plant settings
    for setting in plantSettings.__dict__:
        print(f'{setting} : {plantSettings.__dict__[setting]}')

    project = pscad.project(plantSettings.PSCAD_Namespace)

    #Update pgb names for all unit measurement components
    updateUMs()

    buildFolder : str = project.temp_folder #type: ignore
    cleanBuildfolder(buildFolder) #type: ignore

    project.parameters(time_duration = 999, time_step = plantSettings.PSCAD_Timestep, sample_step = '1000') #type: ignore
    project.parameters(PlotType = '1', output_filename = f'{plantSettings.Projectname}.out') #type: ignore
    project.parameters(SnapType='0', SnapTime='2', snapshot_filename='pannatest5us.snp') #type: ignore

    pscad.remove_all_simulation_sets()
    pmr = pscad.create_simulation_set('MTB')
    pmr.add_tasks(plantSettings.PSCAD_Namespace)
    project_pmr = pmr.task(plantSettings.PSCAD_Namespace)
    project_pmr.parameters(ammunition = maxRank, volley = config.volley, affinity_type = '2') #type: ignore

    pscad.run_simulation_sets('MTB') #type: ignore ??? By sideeffect changes current working directory ???
    os.chdir(executeFolder)

    cleanUpOutFiles(buildFolder, plantSettings.Projectname)

    print('execute.py finished at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

if __name__ == '__main__':
    main()
