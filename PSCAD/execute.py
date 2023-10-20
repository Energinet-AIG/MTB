import sys
sys.path.append(r'C:\Program Files\Python37\Lib\site-packages')
import pandas as pd
import math
import shutil
from datetime import datetime
import pathlib
from os import listdir, mkdir
from os.path import split, splitext, join, exists
from types import SimpleNamespace
import mhi.pscad #type: ignore
from mhi.pscad.utilities.file import File as fileUtil #type: ignore

def moveFiles(srcFolder : str, dstFolder : str, types : list):
    for file in listdir(srcFolder):
        root, typ = splitext(file)
        if typ in types:
            shutil.move(join(srcFolder, file), join(dstFolder, file))

projectFolder = pathlib.Path(__file__).parent.resolve()
pscad = mhi.pscad.application() 
  
print('The simulation started at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

# Read the source excel file    
testCases = pd.ExcelFile(join(projectFolder, 'TestCases.xlsx'))

# Generate the TestCases.txt file from the TestCases.xlsx file
cases = pd.read_excel(testCases, sheet_name='Cases', usecols = 'C:AL')
caseMatrixPath = join(projectFolder, 'TestCases.txt')
cases.to_csv(caseMatrixPath, header=False, index=False)
print('{} saved'.format(caseMatrixPath))

# Read model-specific information from the TestCases.xlsx file
options = SimpleNamespace()

inputs = pd.read_excel(testCases, sheet_name='Input')
options.Pn = inputs['Pn (MW)'][0]                   # Nominal active power (MW)
options.Vn = inputs['Vn_PoC (kV)'][0]               # PoC nominal voltage (LL, RMS, kV)
options.SCR = inputs['SCR'][0]                      # Minimum SCR at PoC
options.XRratio = inputs['X/R'][0]               
options.QmodeQ = int(inputs['Qmode_Q'][0])          # Qref control mode
options.QmodeV = int(inputs['Qmode_V'][0])          # Voltage control mode
options.QmodePF = int(inputs['Qmode_PF'][0])        # Power factor control mode
options.projectname = inputs['ProjectName'][0]      # Project name
options.compiler = inputs['PSCAD compiler'][0]
options.timeStep = inputs['PSCAD TimeStep (us)'][0] # Simulation time step (us)
options.totalRuns = int(inputs['TotalCases'][0])    # Total number of simulation runs (test cases)
options.volleySize = int(inputs['Volley'][0])       # The number of simulations that are run in parallel

pscad.settings(fortran_version = options.compiler)

project = pscad.project(options.projectname)
project.parameters(time_duration = 300, time_step = options.timeStep, sample_step = '1000')   # Units: (time_duration [s], time_step [us], sample_step [us])
project.parameters(PlotType = '1', output_filename = '{}.out'.format(options.projectname))
project.parameters(SnapType='0', SnapTime='2', snapshot_filename='pannatest5us.snp')

# Find the blocks that need model-specific information
blocks = SimpleNamespace()
blocks.vs = project.find('VoltSource')
blocks.vn = project.find('Vn_PoC')
blocks.pn = project.find('Pn_model_block')
blocks.scr = project.find('SCR_block')
blocks.xrratio = project.find('XRratio_block')    
blocks.multimeter = project.find('MultiMeter_PoC')
blocks.QrefCtrl = project.find('QrefCtrl_block')
blocks.VoltCtrl = project.find('VoltCtrl_block')
blocks.PFCtrl = project.find('PFCtrl_block')

# Voltage Source
blocks.vs.parameters(MVA = options.Pn * options.SCR, Vm = options.Vn, Vbase = options.Vn)

# Grid Impedance
blocks.vn.parameters(Value = options.Vn, Dim = '1', )        # Nominal voltage
blocks.pn.parameters(Value = options.Pn, Dim = '1', )        # Nominal power
blocks.scr.parameters(Value = options.SCR, Dim = '1', )      # SCR
blocks.xrratio.parameters(Value=options.XRratio, Dim='1', )  # X/R ratio

# For PoC measurement  
blocks.multimeter.parameters(S = options.Pn, BaseV = options.Vn, BaseA = options.Pn/options.Vn * math.sqrt(2/3))    

# For different reactive power control modes
blocks.QrefCtrl.parameters(A = options.QmodeQ, )
blocks.VoltCtrl.parameters(A = options.QmodeV, )
blocks.PFCtrl.parameters(A = options.QmodePF, )    

# Disable all the output channels in the model
outputchannels = project.find_all('master:pgb')
for output in outputchannels:
    output.disable()

# Find all the PP-MTB output channels
outpchnls = SimpleNamespace()
enabledChannels = ['P_pos_pu_PoC_ENDK', 'Q_pos_pu_PoC_ENDK', 'Q_neg_pu_PoC_ENDK', 'V_A_RMS_pu_PoC_ENDK', 'V_B_RMS_pu_PoC_ENDK', 'V_C_RMS_pu_PoC_ENDK', 'V_pos_pu_PoC_ENDK', 'V_neg_pu_PoC_ENDK', 'I_A_RMS_pu_PoC_ENDK', 'I_B_RMS_pu_PoC_ENDK', 'I_C_RMS_pu_PoC_ENDK', 'Id_pos_pu_PoC_ENDK', 'Id_neg_pu_PoC_ENDK', 'Iq_pos_pu_PoC_ENDK', 'Iq_neg_pu_PoC_ENDK', 'f_ENDK', 'P_PoC_ENDK', 'Q_PoC_ENDK']

for c in enabledChannels:
    (project.find('master:pgb', c)).enable()

# Include the project in the simulation set and set PMR parameters
pscad.remove_all_simulation_sets()
pmr = pscad.create_simulation_set('PMR')
pmr.add_tasks(options.projectname)
project_pmr = pmr.task(options.projectname)
project_pmr.parameters(ammunition = options.totalRuns, volley = options.volleySize, affinity_type = '2') 

# ammunition: total number of simulations
# volley: number of simulations to run in parallel at once
# affinity_type ( =2: Trace all, =1: Trace single (last run only), =0: Disable tracing)

pscad.run_simulation_sets('PMR')    

# Reenable all the output channels in the model
outputchannels = project.find_all('master:pgb')
for output in outputchannels:
    output.enable()

# Get the build folder of the project
tempFolder = project.temp_folder

# Converting .out files to .csv files
for file in listdir(tempFolder):
    path, fileName = split(file)
    root, typ = splitext(fileName)
    if fileName.startswith(options.projectname) and typ == '.out':
        print('Converting {} to .csv'.format(file))
        fileUtil.convert_out_to_csv(tempFolder, fileName, '{}.csv'.format(root))

# Move desired files to an 'Output' folder in the current working directory
outputFolder = join(projectFolder, 'output')

if not exists(outputFolder):
    mkdir(outputFolder)

simoutFolder = join(outputFolder, 'MTB_{}'.format(datetime.now().strftime(r'%d%m%Y%H%M%S')))
mkdir(simoutFolder)
moveFiles(tempFolder, simoutFolder, ['.csv', '.inf'])

#Move .out file away from build folder
runFolder = join(tempFolder, 'MTB_{}'.format(datetime.now().strftime(r'%d%m%Y%H%M%S')))
mkdir(runFolder)
moveFiles(tempFolder, runFolder, ['.out'])

print('The simulation finished at:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))