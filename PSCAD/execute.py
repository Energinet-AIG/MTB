import os
abspath = os.path.abspath(__file__)
dname = os.path.dirname(abspath)
os.chdir(dname)

from configparser import ConfigParser
from types import SimpleNamespace

def readConfig() -> SimpleNamespace:
    cp = ConfigParser(allow_no_value=True)
    cp.read('config.ini')
    parsedConf = cp['config']
    config = SimpleNamespace()
    config.sheetPath = parsedConf['Casesheet path']
    config.pythonPath = parsedConf['Python path']
    config.kp = parsedConf['Kp']
    config.ti = parsedConf['Ti']
    return config

config = readConfig()
import sys
sys.path.append(config.pythonPath)

import pandas as pd 
from datetime import datetime
from warnings import warn
from os.path import join, split, splitext, exists
from os import listdir, mkdir
from shutil import move, rmtree
from PSCAD_interface import waveform, recordedWaveform, fortranFile
import mhi.pscad #type: ignore
from mhi.pscad.utilities.file import File as fileUtil #type: ignore

def moveFiles(srcFolder : str, dstFolder : str, types : list):
    for file in listdir(srcFolder):
        root, typ = splitext(file)
        if typ in types:
            move(join(srcFolder, file), join(dstFolder, file))

def readPlantSettings(sheetPath : str) -> SimpleNamespace:
    inputs = pd.read_excel(sheetPath, sheet_name='Input', header=None)
    inputs.index = inputs.iloc[:, 0]
    inputs = inputs.iloc[1:, 1]

    settings = SimpleNamespace()
    
    settings.Projectname = str(inputs['Projectname'])
    settings.PSCAD_Namespace = str(inputs['PSCAD Namespace'])
    settings.Pn = float(inputs['Pn'])
    settings.Uc = float(inputs['Uc'])
    settings.Un = float(inputs['Un'])
    settings.Area = str(inputs['Area'])
    settings.SCR_min = float(inputs['SCR min'])
    settings.SCR_tuning = float(inputs['SCR tuning'])
    settings.SCR_max = float(inputs['SCR max'])
    settings.XR_SCR_min = float(inputs['X/R SCR min'])
    settings.XR_SCR_tuning = float(inputs['X/R SCR tuning'])
    settings.XR_SCR_max = float(inputs['X/R SCR max'])
    settings.R0 = float(inputs['R0'])
    settings.X0 = float(inputs['X0'])
    settings.Default_Q_mode = str(inputs['Default Q mode'])
    settings.PSCAD_Timestep = float(inputs['PSCAD Timestep'])
    settings.PSCAD_Volley = int(inputs['PSCAD Volley'])
    settings.Initialization_time = float(inputs['Initialization time'])

    return settings

def readCases(sheetPath : str) -> pd.DataFrame:
    cases = pd.read_excel(sheetPath, sheet_name='Cases', header=1)
    cases.index = cases.iloc[:, 0]
    cases = cases.iloc[:,1:]
    return cases

def cleanBuildfolder(folder):
    try:
        rmtree(folder)
    except:
        pass

def prepareMeasFile(path : str, measFileFolder : str = '', columns : list = [], pfFormat : bool = False) -> list:
    pathFolder, pathFilename = split(path)
    pathName, pathExtension = splitext(pathFilename)
    
    reader = open(path, 'r')

    if pathExtension.lower() == '.meas' or pathExtension.lower() == '.out':    
        lineBuffer = reader.readlines()
        data = []

        def parseLine(line, linenr, file):
            floatBuffer = ''
            values = []
            line += '\n'

            for c in line:
                if not c in [',',' ','\t','\n']:
                    floatBuffer += c
                else:
                    if len(floatBuffer) > 0:
                        try:
                            values.append(float(floatBuffer))
                        except ValueError:
                            print(f'Could not parse line nr: {linenr} in "{file}". Value "{floatBuffer}" not understandable as float. Exiting.')
                            exit()
                        floatBuffer = ''
            return values
            
        i = 2
        for line in lineBuffer[1:]:
            data.append(parseLine(line, i, path))
            i += 1

        data = pd.DataFrame(data)

    elif pathExtension.lower() == '.csv':
        data = pd.read_csv(path, sep=';', decimal='.', header=None, skiprows=1)
    else:
        exit(f'Unknown filetype of: {path}.')
    
    data = data.set_index(0)         
    data.sort_index(ascending=True, inplace=True)

    if not exists(measFileFolder):
        mkdir(measFileFolder)

    recordingLength = float(data.index[-1])
    recordings = []

    if columns == []:
        columns = data.columns

    for c in columns:
        if not c in data.columns:
            exit(f'Column "{c}" not in {path}.')
        else:
            recFile = join(measFileFolder, f'{pathName}_{c}.out')
            measData = (data[c]).to_csv(None, sep = ' ', header=False, index_label=False).replace('\r\n','\n')
            if pfFormat:
                measData = '1\n' + measData
            else:
                measData = '\n' + measData
            f = open(recFile, 'w')
            f.write(measData)
            f.close()
            recordings.append((recFile, recordingLength))

    return recordings

def constructRecWf(path : str, measFileFolder : str, columns : list) -> list:
    preparedFiles = prepareMeasFile(path, measFileFolder, columns)
    rwfObjects = []
    for pf in preparedFiles:
        rwfObjects.append(recordedWaveform(pf[0], pf[1]))
    return rwfObjects    

def setupBaseVSC(rank : int, case : pd.Series, fortran : fortranFile, plantSettings : SimpleNamespace) -> None:
    ivrns = case['Inst. Voltage recording (EMT ONLY)'].replace(' ', '') != ''
    vrns = case['Voltage recording'].replace(' ', '') != ''
    prns = case['Phase recording'].replace(' ', '') != ''
    frns = case['Freq. recording'].replace(' ', '') != ''

    if ivrns:
        Vmode = 2

        if vrns or prns or frns:
            warn(f'Rank {rank}: Instantenous voltage recording given. Conflicting recordings ignored.')    

        rwfObjs = constructRecWf(case['Inst. Voltage recording (EMT ONLY)'], 'inputFiles', [1,2,3])

        if len(rwfObjs) != 3:
            exit(f'Rank {rank}: The Instantenous voltage recording has wrong number of columns (!=3).')

        fortran['mtb_s_vref_pu'][rank] = waveform(0)
        fortran['mtb_s_phref_deg'][rank] = waveform(0)
        fortran['mtb_s_fref_hz'][rank] = waveform(0) 
        fortran['mtb_s_varef_pu'][rank] = rwfObjs[0]
        fortran['mtb_s_vbref_pu'][rank] = rwfObjs[1]
        fortran['mtb_s_vcref_pu'][rank] = rwfObjs[2]

    elif vrns:
        Vmode = 1
        rwfObjs = constructRecWf(case['Voltage recording'], 'inputFiles', [1])

        if len(rwfObjs) != 1:
            exit(f'Rank {rank}: The voltage recording has wrong number of columns (!=1).')

        fortran['mtb_s_vref_pu'][rank] = rwfObjs[0]
        fortran['mtb_s_varef_pu'][rank] = waveform(0)
        fortran['mtb_s_vbref_pu'][rank] = waveform(0)
        fortran['mtb_s_vcref_pu'][rank] = waveform(0)
    else:
        Vmode = 0
        fortran['mtb_s_vref_pu'][rank] = waveform(-abs(case['U0']))
        fortran['mtb_s_varef_pu'][rank] = waveform(0)
        fortran['mtb_s_vbref_pu'][rank] = waveform(0)
        fortran['mtb_s_vcref_pu'][rank] = waveform(0)       

    if Vmode != 2:
        if prns:
            rwfObjs = constructRecWf(case['Phase recording'], 'inputFiles', [1])

            if len(rwfObjs) != 1:
                exit(f'Rank {rank}: The phase recording has wrong number of columns (!=1).')

            fortran['mtb_s_phref_deg'][rank] = rwfObjs[0]
        else:
            fortran['mtb_s_phref_deg'][rank] = waveform(0)

        if frns:
            rwfObjs = constructRecWf(case['Freq. recording'], 'inputFiles', [1])

            if len(rwfObjs) != 1:
                exit(f'Rank {rank}: The frequency recording has wrong number of columns (!=1).')

            fortran['mtb_s_fref_hz'][rank] = rwfObjs[0]
        else:
            fortran['mtb_s_fref_hz'][rank] = waveform(50)    

    fortran['mtb_t_vmode'][rank] = Vmode
    fortran['mtb_t_r0_ohm'][rank] = plantSettings.R0
    fortran['mtb_t_x0_ohm'][rank] = plantSettings.X0
    fortran['mtb_s_scr'][rank] = waveform(case['SCR0'])
    fortran['mtb_s_xr'][rank] = waveform(case['XR0'])
    fortran['mtb_s_dvref_pu'][rank] = waveform(0)

def setupBasePlantRef(rank : int, case : pd.Series, fortran : fortranFile, plantSettings : SimpleNamespace) -> None:
    prrns = case['Pref recording'].replace(' ', '') != ''
    qrrns = case['Qref recording'].replace(' ', '') != ''

    if prrns:
        rwfObjs = constructRecWf(case['Pref recording'], 'inputFiles', [1])

        if len(rwfObjs) != 1:
            exit(f'Rank {rank}: The pref recording has wrong number of columns (!=1).')

        fortran['mtb_s_pref_pu'][rank] = rwfObjs[0]
    else:
        fortran['mtb_s_pref_pu'][rank] = waveform(case['P0'],0)

    if qrrns:
        rwfObjs = constructRecWf(case['Qref recording'], 'inputFiles', [1])

        if len(rwfObjs) != 1:
            exit(f'Rank {rank}: The qref recording has wrong number of columns (!=1).')

        fortran['mtb_s_qref_pu'][rank] = rwfObjs[0]
    else:
        fortran['mtb_s_qref_pu'][rank] = waveform(case['Qref0'],0)

    if case['Qmode'] == 'PF' or case['Qmode'] == 'Default' and plantSettings.Default_Q_mode == 'PF':
        fortran['mtb_t_qmode'][rank] = 2 
    elif case['Qmode'] == 'Q(U)' or case['Qmode'] == 'Default' and plantSettings.Default_Q_mode == 'Q(U)':
        fortran['mtb_t_qmode'][rank] = 1
    else:
        fortran['mtb_t_qmode'][rank] = 0

    if case['FSM']:
        fortran['mtb_t_fsm'][rank] = 1
    else:
        fortran['mtb_t_fsm'][rank] = 0

def setupSimTime(rank : int, case : pd.Series, fortran : fortranFile, plantSettings : SimpleNamespace) -> None:
    if pd.isna(case['Simulationtime']):
        recCands = ['mtb_s_vref_pu', 'mtb_s_phref_deg', 'mtb_s_fref_hz', 'mtb_s_varef_pu', 'mtb_s_pref_pu', 'mtb_s_qref_pu']
        simTime = -1.0

        for signal in recCands:
            if fortran[signal][rank].type == 1:
                simTime = max(simTime, fortran[signal][rank].len)

        if simTime == -1.0:
            exit(f'Rank {rank}: Neither simulationtime or any recording given.')
        
        fortran['mtb_t_simtime_s'][rank] = simTime
    else:
        fortran['mtb_t_simtime_s'][rank] = case['Simulationtime'] + plantSettings.Initialization_time

def typeValidateCase(case : pd.Series) -> None:
    for signal in ['Inst. Voltage recording (EMT ONLY)', 'Voltage recording', 'Phase recording', 'Freq. recording', 'Pref recording', 'Qref recording']:
        case[signal] = str(case[signal])
        if case[signal] == 'nan':
            case[signal] = ''

def parseEvents(rank : int, case : pd.Series, fortran: fortranFile) -> None:
    initOffset = fortran['mtb_c_inittime_s'].value()
    
    i = 0
    while(True):
        typeLabel = f'type.{i}' if i > 0 else 'type'
        timeLabel = f'time.{i}' if i > 0 else 'time'
        x1Label = f'X1.{i}' if i > 0 else 'X1'
        x2Label = f'X2.{i}' if i > 0 else 'X2'

        if typeLabel in case.index and timeLabel in case.index and x1Label in case.index and x2Label in case.index:
            eventType = case[typeLabel]
            eventTime = float(case[timeLabel]) + initOffset
            eventX1 = float(case[x1Label])
            eventX2 = float(case[x2Label])
            
            if type(eventType) == float:
                pass
            elif eventType == 'Pctrl ref.':
                if fortran['mtb_s_pref_pu'][rank].type == 0:
                    fortran['mtb_s_pref_pu'][rank].add(eventTime, eventX1, eventX2)

            elif eventType == 'Qctrl ref.':
                if fortran['mtb_s_qref_pu'][rank].type == 0:
                    fortran['mtb_s_qref_pu'][rank].add(eventTime, eventX1, eventX2)

            elif eventType == 'Voltage':
                if fortran['mtb_t_vmode'][rank] == 0:
                    fortran['mtb_s_vref_pu'][rank].add(eventTime, eventX1, eventX2)

            elif eventType == 'dVoltage':
                if fortran['mtb_t_vmode'][rank] != 2:
                    fortran['mtb_s_dvref_pu'][rank].add(eventTime, eventX1, eventX2)

            elif eventType == 'Phase':
                if fortran['mtb_t_vmode'][rank] != 2 and fortran['mtb_s_phref_deg'][rank].type == 0:
                    fortran['mtb_s_phref_deg'][rank].add(eventTime, eventX1, eventX2)
            
            elif eventType == 'Frequency':
                if fortran['mtb_t_vmode'][rank] != 2 and fortran['mtb_s_fref_hz'][rank].type == 0:
                    fortran['mtb_s_fref_hz'][rank].add(eventTime, eventX1, eventX2)

            elif eventType == 'SCR':
                fortran['mtb_s_scr'][rank].add(eventTime, eventX1, 0.0)
                fortran['mtb_s_xr'][rank].add(eventTime, eventX2, 0.0)

            elif eventType == '3p fault':
                fortran['flt_s_type'][rank].add(eventTime, 7.0, 0.0)
                fortran['flt_s_type'][rank].add(eventTime + eventX2, 0.0, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX1, 0.0)

            elif eventType == '2p-g fault':
                fortran['flt_s_type'][rank].add(eventTime, 5.0, 0.0)
                fortran['flt_s_type'][rank].add(eventTime + eventX2, 0.0, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX1, 0.0)            

            elif eventType == '2p fault':
                fortran['flt_s_type'][rank].add(eventTime, 3.0, 0.0)
                fortran['flt_s_type'][rank].add(eventTime + eventX2, 0.0, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX1, 0.0)    

            elif eventType == '1p fault':
                fortran['flt_s_type'][rank].add(eventTime, 1.0, 0.0)
                fortran['flt_s_type'][rank].add(eventTime + eventX2, 0.0, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX1, 0.0)    

            elif eventType == '3p fault (ohm)':
                fortran['flt_s_type'][rank].add(eventTime, 8.0, 0.0)
                fortran['flt_s_rf_ohm'][rank].add(eventTime, eventX1, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX2, 0.0)

            elif eventType == '2p-g fault (ohm)':
                fortran['flt_s_type'][rank].add(eventTime, 6.0, 0.0)
                fortran['flt_s_rf_ohm'][rank].add(eventTime, eventX1, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX2, 0.0)          

            elif eventType == '2p fault (ohm)':
                fortran['flt_s_type'][rank].add(eventTime, 4.0, 0.0)
                fortran['flt_s_rf_ohm'][rank].add(eventTime, eventX1, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX2, 0.0)   

            elif eventType == '1p fault (ohm)':
                fortran['flt_s_type'][rank].add(eventTime, 2.0, 0.0)
                fortran['flt_s_rf_ohm'][rank].add(eventTime, eventX1, 0.0)
                fortran['flt_s_resxf'][rank].add(eventTime, eventX2, 0.0)

            elif eventType == 'Clear fault':
                fortran['flt_s_type'][rank].add(eventTime, 0.0, 0.0)
            
            else:
                exit(f'Unknown event type in rank {rank}: {eventType}.')
        
        else:
            return
        i += 1

def cleanUpOutFiles(buildFolder, projectName):
    # Converting .out files to .csv files
    for file in listdir(buildFolder):
        path, fileName = split(file)
        root, typ = splitext(fileName)
        if fileName.startswith(projectName) and typ == '.out':
            print('Converting {} to .csv'.format(file))
            fileUtil.convert_out_to_csv(buildFolder, fileName, '{}.csv'.format(root))

    # Move desired files to an 'Output' folder in the current working directory
    outputFolder = 'output'

    if not exists(outputFolder):
        mkdir(outputFolder)

    simoutFolder = join(outputFolder, 'MTB_{}'.format(datetime.now().strftime(r'%d%m%Y%H%M%S')))
    mkdir(simoutFolder)
    moveFiles(buildFolder, simoutFolder, ['.csv', '.inf'])

    #Move .out file away from build folder
    runFolder = join(buildFolder, 'MTB_{}'.format(datetime.now().strftime(r'%d%m%Y%H%M%S')))
    mkdir(runFolder)
    moveFiles(buildFolder, runFolder, ['.out'])

pscad = mhi.pscad.application() 

plantSettings = readPlantSettings(config.sheetPath)

project = pscad.project(plantSettings.PSCAD_Namespace)
buildFolder = project.temp_folder
cleanBuildfolder(buildFolder)

cases = readCases(config.sheetPath)

fortran = fortranFile('ppmtb.f')

# Voltage source control
fortran.newTimeInv('mtb_t_vmode', 0)
fortran.newSignal('mtb_s_vref_pu', waveform(1))
fortran.newSignal('mtb_s_dvref_pu', waveform(0))
fortran.newSignal('mtb_s_phref_deg', waveform(0))
fortran.newSignal('mtb_s_fref_hz', waveform(50))
fortran.newSignal('mtb_s_varef_pu', waveform(0))
fortran.newSignal('mtb_s_vbref_pu', waveform(0))
fortran.newSignal('mtb_s_vcref_pu', waveform(0))
fortran.newSignal('mtb_s_scr', plantSettings.SCR_min)
fortran.newSignal('mtb_s_xr', plantSettings.XR_SCR_min)
fortran.newTimeInv('mtb_t_r0_ohm', 0)
fortran.newTimeInv('mtb_t_x0_ohm', 0)

# Plant references and outputs
fortran.newSignal('mtb_s_pref_pu', waveform(1))
fortran.newSignal('mtb_s_qref_pu', waveform(0))
fortran.newTimeInv('mtb_t_qmode', 0)
fortran.newTimeInv('mtb_t_fsm', 0)

# Constants
fortran.newConstant('mtb_c_pn', plantSettings.Pn)
fortran.newConstant('mtb_c_vbase', plantSettings.Un)
fortran.newConstant('mtb_c_vc', plantSettings.Uc)
fortran.newConstant('mtb_c_inittime_s', plantSettings.Initialization_time)
fortran.newConstant('mtb_c_kp', config.kp)
fortran.newConstant('mtb_c_ti_s', config.ti)

# Time and rank control
fortran.newTimeInv('mtb_t_simtime_s', -1)

# Fault
fortran.newSignal('flt_s_type', waveform(0))
fortran.newSignal('flt_s_rf_ohm', waveform(0))
fortran.newSignal('flt_s_resxf', waveform(0))

for rank, case in cases.iterrows():
    if case['EMT']:
        typeValidateCase(case)
        setupBaseVSC(rank, case, fortran, plantSettings)
        setupBasePlantRef(rank, case, fortran, plantSettings)
        setupSimTime(rank, case, fortran, plantSettings)
        fortran['mtb_t_fsm'][rank] = float(case['FSM'])
        fortran['flt_s_type'][rank] = waveform(0)
        fortran['flt_s_rf_ohm'][rank] = waveform(0)
        fortran['flt_s_resxf'][rank] = waveform(0)
        parseEvents(rank, case, fortran)

fortran.render()

project.parameters(time_duration = 999, time_step = plantSettings.PSCAD_Timestep, sample_step = '1000')  

project.parameters(PlotType = '1', output_filename = '{}.out'.format(plantSettings.Projectname))
project.parameters(SnapType='0', SnapTime='2', snapshot_filename='pannatest5us.snp')
pscad.remove_all_simulation_sets()
pmr = pscad.create_simulation_set('PMR')
pmr.add_tasks(plantSettings.PSCAD_Namespace)
project_pmr = pmr.task(plantSettings.PSCAD_Namespace)
project_pmr.parameters(ammunition = max(cases.index), volley = plantSettings.PSCAD_Volley, affinity_type = '2') 
pscad.run_simulation_sets('PMR') #By sideeffect changes current working directory ? :S
os.chdir(dname)

cleanUpOutFiles(buildFolder, plantSettings.Projectname)

print('The simulation finished at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))