import os
import sys
sys.path.append(r'C:\Program Files\Python37\Lib\site-packages')
import pandas as pd
import math
import shutil
import datetime
import pathlib
from os import listdir
from os.path import split, splitext, join
import mhi.pscad #type: ignore
from mhi.pscad.utilities.file import File as fileUtil #type: ignore

def moveFiles(srcFolder : str, dstFolder : str, types : list):
    for file in listdir(srcFolder):
        root, typ = splitext(file)
        if typ in types:
            shutil.move(join(srcFolder, file), join(dstFolder, file))

projectFolder = pathlib.Path(__file__).parent.resolve()
pscad = mhi.pscad.application() 
  
print('The simulation starts at:', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

# Read the source excel file    
testCases = pd.ExcelFile(os.path.join(projectFolder, 'TestCases.xlsx'))

# Generate the TestCases.txt file from the TestCases.xlsx file
Cases = pd.read_excel(testCases, sheet_name='Cases', usecols = 'C:AL')
Cases.to_csv(os.path.join(projectFolder, 'TestCases.txt'), header=False, index=False)

# Read model-specific information from the TestCases.xlsx file
Inputs = pd.read_excel(testCases, sheet_name='Input')
Pn_model = Inputs['Pn (MW)'][0]                 # Nominal active power (MW)
Vn = Inputs['Vn_PoC (kV)'][0]                   # PoC nominal voltage (LL, RMS, kV)
SCR = Inputs['SCR'][0]                          # Minimum SCR at PoC
XRratio = Inputs['X/R'][0]               
Qmode_Q = int(Inputs['Qmode_Q'][0])             # Qref control mode
Qmode_V = int(Inputs['Qmode_V'][0])             # Voltage control mode
Qmode_PF = int(Inputs['Qmode_PF'][0])           # Power factor control mode
projectname = Inputs['ProjectName'][0]          # Project name
compiler = Inputs['PSCAD compiler'][0]
TimeStep = Inputs['PSCAD TimeStep (us)'][0]     # Simulation time step (us)
TotalSmltRun = int(Inputs['TotalCases'][0])     # Total number of simulation runs (test cases)
Volley = int(Inputs['Volley'][0])               # The number of simulations that are run in parallel

# Change to the right complier
pscad.settings(fortran_version=compiler)

# Find the model
project = pscad.project(projectname)
# Modify 'Description'
describe = project.parameters()['description']
project.parameters(description= describe + '_Energinet_' + datetime.datetime.now().strftime('%Y%m%d'))
# Simulation time and time step
MaxSmltTime = 300       # Simulation time for the whole project, no shorter than the longest run
project.parameters(time_duration=MaxSmltTime, time_step=TimeStep, sample_step='1000')   # Units: (time_duration [s], time_step [us], sample_step [us])
# Data output
project.parameters(PlotType='1', output_filename=projectname+'.out')
# Snapshot
project.parameters(SnapType='0', SnapTime='2', snapshot_filename='pannatest5us.snp')

# Find the blocks that need model-specific information
voltsource = project.find('VoltSource')
vn = project.find('Vn_PoC')
pn = project.find('Pn_model_block')
scr = project.find('SCR_block')
xrratio = project.find('XRratio_block')    
multimeter = project.find('MultiMeter_PoC')
QrefCtrl = project.find('QrefCtrl_block')
VoltCtrl = project.find('VoltCtrl_block')
PFCtrl = project.find('PFCtrl_block')

## Initialize components (input the model-specific information into the corresponding blocks)

# Calculate the capacity of the voltage source (MVA)    
Pn_VS = Pn_model*SCR
# Calculate the base current for PoC measurement
Ib_model = Pn_model/Vn*math.sqrt(2/3)

# For Voltage Source
voltsource.parameters(MVA=Pn_VS, Vm=Vn, Vbase=Vn)
# For Grid Impedance
vn.parameters(Value=Vn, Dim='1', )              # Nominal voltage
pn.parameters(Value=Pn_model, Dim='1', )        # Nominal power
scr.parameters(Value=SCR, Dim='1', )            # SCR
xrratio.parameters(Value=XRratio, Dim='1', )    # X/R ratio
# For PoC measurement  
multimeter.parameters(S=Pn_model, BaseV=Vn, BaseA=Ib_model)    
# For different reactive power control modes
QrefCtrl.parameters(A=Qmode_Q, )
VoltCtrl.parameters(A=Qmode_V, )
PFCtrl.parameters(A=Qmode_PF, )    

# Disable all the output channels in the model
outputchannels = project.find_all('master:pgb')
for output in outputchannels:
    output.disable()

# Find all the output channels defined by Energinet
P_pos_PoC = project.find('master:pgb', 'P_pos_pu_PoC_ENDK')
Q_pos_PoC = project.find('master:pgb', 'Q_pos_pu_PoC_ENDK')
Q_neg_PoC = project.find('master:pgb', 'Q_neg_pu_PoC_ENDK')
U_PoC_A_RMS = project.find('master:pgb', 'V_A_RMS_pu_PoC_ENDK')
U_PoC_B_RMS = project.find('master:pgb', 'V_B_RMS_pu_PoC_ENDK')
U_PoC_C_RMS = project.find('master:pgb', 'V_C_RMS_pu_PoC_ENDK')
U_PoC_pos = project.find('master:pgb', 'V_pos_pu_PoC_ENDK')
U_PoC_neg = project.find('master:pgb', 'V_neg_pu_PoC_ENDK')
I_PoC_A_RMS = project.find('master:pgb', 'I_A_RMS_pu_PoC_ENDK')
I_PoC_B_RMS = project.find('master:pgb', 'I_B_RMS_pu_PoC_ENDK')
I_PoC_C_RMS = project.find('master:pgb', 'I_C_RMS_pu_PoC_ENDK')
I_PoC_d_pos = project.find('master:pgb', 'Id_pos_pu_PoC_ENDK')
I_PoC_d_neg = project.find('master:pgb', 'Id_neg_pu_PoC_ENDK')
I_PoC_q_pos = project.find('master:pgb', 'Iq_pos_pu_PoC_ENDK')
I_PoC_q_neg = project.find('master:pgb', 'Iq_neg_pu_PoC_ENDK')
f = project.find('master:pgb', 'f_ENDK')
P_PoC = project.find('master:pgb', 'P_PoC_ENDK')
Q_PoC = project.find('master:pgb', 'Q_PoC_ENDK')

# Enable only the output channels from Energinet
P_pos_PoC.enable()
Q_pos_PoC.enable()
Q_neg_PoC.enable()
U_PoC_A_RMS.enable()
U_PoC_B_RMS.enable()
U_PoC_C_RMS.enable()
U_PoC_pos.enable()
U_PoC_neg.enable()
# I_PoC_A_RMS.enable()
# I_PoC_B_RMS.enable()
# I_PoC_C_RMS.enable()
I_PoC_d_pos.enable()
I_PoC_d_neg.enable()
I_PoC_q_pos.enable()
I_PoC_q_neg.enable()
f.enable()
P_PoC.enable()
Q_PoC.enable()
    
# Include the project in the simulation set and set PMR parameters
pscad.remove_all_simulation_sets()          # Firstly, remove the previous simulation set
pmr = pscad.create_simulation_set('PMR')    # Create a new simulation set
pmr.add_tasks(projectname)
project_pmr = pmr.task(projectname)
project_pmr.parameters(ammunition=TotalSmltRun, volley=Volley, affinity_type='2') 

# ammunition: total number of simulations
# volley: number of simulations to run in parallel at once
# affinity_type ( =2: Trace all, =1: Trace single (last run only), =0: Disable tracing)

pscad.run_simulation_sets('PMR')    

# Get the build folder of the project
tempFolder = project.temp_folder

# Converting .out files to .csv files
for file in listdir(tempFolder):
    path, fileName = split(file)
    root, typ = splitext(fileName)
    if fileName.startswith(projectname) and typ == '.out':
        print('Converting {} to .csv'.format(file))
        fileUtil.convert_out_to_csv(tempFolder, fileName, '{}.csv'.format(root))

# Move desired files to an 'Output' folder in the current working directory
outputFolder = join(projectFolder, 'output')

if os.path.exists(outputFolder):
    shutil.rmtree(outputFolder)

os.mkdir(outputFolder)
moveFiles(tempFolder, outputFolder, ['.csv', '.inf'])

#Move .out file away from build folder
runFolder = join(tempFolder, 'MTB_{}'.format(datetime.datetime.now().strftime(r'%d%m%Y%H%M%S')))
os.mkdir(runFolder)
moveFiles(tempFolder, runFolder, ['.out'])

print('The simulation finishes at:', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))