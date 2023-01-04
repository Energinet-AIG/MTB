#!/usr/bin/env python3

import sys, os
sys.path.append(r"C:\Program Files\Python37\Lib\site-packages")
import numpy as np
import pandas as pd
import math as m
import shutil
import datetime
import runpy
import pathlib
project_folder = pathlib.Path(__file__).parent.resolve()
plot_folder = os.path.join(project_folder, "Plot")
sys.path.append(plot_folder)

import mhi.pscad
import mhi.pscad.utilities.file 
from mhi.pscad.utilities.file import File
from mhi.pscad.utilities.file import OutFile

with mhi.pscad.application() as pscad:

    # Start the timer
    t_start = datetime.datetime.now()
print('The simulation starts at:', t_start.strftime("%Y-%m-%d %H:%M:%S"))
print('\n')

# Read the source excel file    
TestCases = pd.ExcelFile(os.path.join(project_folder, "TestCases.xlsx"))
# Generate the TestCases.txt file from the TestCases.xlsx file
Cases = pd.read_excel(TestCases, sheet_name='Cases', usecols = "C:AL")
Cases.to_csv(os.path.join(project_folder, 'TestCases.txt'), header=False, index=False)


# Read model-specific information from the TestCases.xlsx file
Inputs = pd.read_excel(TestCases, sheet_name='Input')
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


## 
## Based on the input information to change the model settings and write the information into the model
##

## Project settings

# Change to the right complier
pscad.settings(fortran_version=compiler)

# Find the model
project = pscad.project(projectname)
# Modify 'Description'
describe = project.parameters()["description"]
project.parameters(description= describe + "_Energinet_" + t_start.strftime("%Y%m%d"))
# Simulation time and time step
MaxSmltTime = 300       # Simulation time for the whole project, no shorter than the longest run
project.parameters(time_duration=MaxSmltTime, time_step=TimeStep, sample_step="1000")   # Units: (time_duration [s], time_step [us], sample_step [us])
# Data output
project.parameters(PlotType="1", output_filename=projectname+".out")
# Snapshot
project.parameters(SnapType="0", SnapTime="2", snapshot_filename="pannatest5us.snp")

# Find the blocks that need model-specific information
voltsource = project.find("VoltSource")
vn = project.find("Vn_PoC")
pn = project.find("Pn_model_block")
scr = project.find("SCR_block")
xrratio = project.find("XRratio_block")    
multimeter = project.find("MultiMeter_PoC")
QrefCtrl = project.find("QrefCtrl_block")
VoltCtrl = project.find("VoltCtrl_block")
PFCtrl = project.find("PFCtrl_block")

## Initialize components (input the model-specific information into the corresponding blocks)

# Calculate the capacity of the voltage source (MVA)    
Pn_VS = Pn_model*SCR
# Calculate the base current for PoC measurement
Ib_model = Pn_model/Vn*m.sqrt(2/3)

# For Voltage Source
voltsource.parameters(MVA=Pn_VS, Vm=Vn, Vbase=Vn)
# For Grid Impedance
vn.parameters(Value=Vn, Dim="1", )              # Nominal voltage
pn.parameters(Value=Pn_model, Dim="1", )        # Nominal power
scr.parameters(Value=SCR, Dim="1", )            # SCR
xrratio.parameters(Value=XRratio, Dim="1", )    # X/R ratio
# For PoC measurement  
multimeter.parameters(S=Pn_model, BaseV=Vn, BaseA=Ib_model)    
# For different reactive power control modes
QrefCtrl.parameters(A=Qmode_Q, )
VoltCtrl.parameters(A=Qmode_V, )
PFCtrl.parameters(A=Qmode_PF, )    


## Plotting settings

# Disable all the output channels in the model
outputchannels = project.find_all("master:pgb")
for output in outputchannels:
    output.disable()

# Find all the output channels defined by Energinet
P_pos_PoC = project.find("master:pgb", "P_pos_pu_PoC_ENDK")
Q_pos_PoC = project.find("master:pgb", "Q_pos_pu_PoC_ENDK")
Q_neg_PoC = project.find("master:pgb", "Q_neg_pu_PoC_ENDK")
U_PoC_A_RMS = project.find("master:pgb", "V_A_RMS_pu_PoC_ENDK")
U_PoC_B_RMS = project.find("master:pgb", "V_B_RMS_pu_PoC_ENDK")
U_PoC_C_RMS = project.find("master:pgb", "V_C_RMS_pu_PoC_ENDK")
U_PoC_pos = project.find("master:pgb", "V_pos_pu_PoC_ENDK")
U_PoC_neg = project.find("master:pgb", "V_neg_pu_PoC_ENDK")
I_PoC_A_RMS = project.find("master:pgb", "I_A_RMS_pu_PoC_ENDK")
I_PoC_B_RMS = project.find("master:pgb", "I_B_RMS_pu_PoC_ENDK")
I_PoC_C_RMS = project.find("master:pgb", "I_C_RMS_pu_PoC_ENDK")
I_PoC_d_pos = project.find("master:pgb", "Id_pos_pu_PoC_ENDK")
I_PoC_d_neg = project.find("master:pgb", "Id_neg_pu_PoC_ENDK")
I_PoC_q_pos = project.find("master:pgb", "Iq_pos_pu_PoC_ENDK")
I_PoC_q_neg = project.find("master:pgb", "Iq_neg_pu_PoC_ENDK")
f = project.find("master:pgb", "f_ENDK")
P_PoC = project.find("master:pgb", "P_PoC_ENDK")
Q_PoC = project.find("master:pgb", "Q_PoC_ENDK")

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



##
## Simulation
##
    
# Include the project in the simulation set and set PMR parameters
pscad.remove_all_simulation_sets()          # Firstly, remove the previous simulation set
pmr = pscad.create_simulation_set("PMR")    # Create a new simulation set
pmr.add_tasks(projectname)
project_pmr = pmr.task(projectname)
project_pmr.parameters(ammunition=TotalSmltRun, volley=Volley, affinity_type="2")  
# ammunition: total number of simulations
# volley: number of simulations to run in parallel at once
# affinity_type ( =2: Trace all, =1: Trace single (last run only), =0: Disable tracing)

# 1st round of simulation runs
pscad.run_simulation_sets("PMR")    

# Stop the timer 1
t_stop = datetime.datetime.now()
print('The simulation finishes at:', t_stop.strftime("%Y-%m-%d %H:%M:%S"))

t_diff = t_stop - t_start

total_seconds = t_diff.total_seconds()
hour = int(total_seconds//3600)
minute = int(total_seconds//60-60*hour)
second = int(total_seconds-3600*hour-60*minute)

print(f'The time spent in simulation is: {hour} h {minute} min {second} s')
print('\n')
print('Please wait for the plotting of results.')
print('\n')


##
## Post-simulation processing
##

# Get the build folder of the project
src_folder = project.temp_folder

# Converting .out files to .csv files
for x in range(TotalSmltRun):
    y = x+1;
    if y < 10:
        File.convert_out_to_csv(src_folder, projectname+"_0"+str(y)+"_01"+".out", projectname+"_0"+str(y)+"_01"+".csv")
        File.convert_out_to_csv(src_folder, projectname+"_0"+str(y)+"_02"+".out", projectname+"_0"+str(y)+"_02"+".csv")
    else:
        File.convert_out_to_csv(src_folder, projectname+"_"+str(y)+"_01"+".out", projectname+"_"+str(y)+"_01"+".csv")
        File.convert_out_to_csv(src_folder, projectname+"_"+str(y)+"_02"+".out", projectname+"_"+str(y)+"_02"+".csv")

# Move desired files to an 'Output' folder in the current working directory
Output = os.path.join(project_folder, "Output")

if os.path.exists(Output):
    shutil.rmtree(Output)

File.move_files(src_folder, Output, ".csv", ".inf")


## Plotting
plot_main = os.path.join(plot_folder, "Main.py")
runpy.run_path(plot_main)

result_folder = os.path.join(plot_folder, "Data\\Result_emt")

# Stop the timer 2
t_end = datetime.datetime.now()

print('The plotting finishes at:', t_end.strftime("%Y-%m-%d %H:%M:%S"))
print('\n')

t_diff = t_end - t_start

total_seconds = t_diff.total_seconds()
hour = int(total_seconds//3600)
minute = int(total_seconds//60-60*hour)
second = int(total_seconds-3600*hour-60*minute)


print('The execution is finished completely.')
print(f'The total time spent is: {hour} h {minute} min {second} s')
print('\n')
print('Please check the results here: ', result_folder)
















