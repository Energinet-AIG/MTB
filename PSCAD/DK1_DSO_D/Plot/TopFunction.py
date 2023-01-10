import os
import pathlib
import shutil
import mhi.pscad
import mhi.pscad.utilities.file 
from mhi.pscad.utilities.file import File

from datetime import date
import pandas as pd
from plotly_graphs import run_all
from create_report import create_document
from merge_csv import merge_emt_csv_files


def execute(plot_flag, RMS_flag, EMT_flag, report_flag):

    script_folder = pathlib.Path(__file__).parent.resolve()     # Get the directory of the folder where the script lies
    project_wd = os.path.dirname(script_folder)                 # Get the project working directory, (which is the parent folder of script_folder), where 'TestCases.xlsx' lies.
    output_EMT = project_wd + "\\Output"                        # This is the folder where EMT results (both csv and inf files) are stored
    plot_folder = script_folder
    data_folder = os.path.join(plot_folder, "Data")                        # This is the path to the csv files, both RMS and EMT.
    
    
    if not os.path.exists(output_EMT):
        print("Error: No 'Output' folder in the working folder")
    
    emt_folder = os.path.join(data_folder, "emt")       # Where the EMT results are moved to for plotting
    
    if os.path.exists(emt_folder):
        shutil.rmtree(emt_folder)

    File.move_files(output_EMT, emt_folder, ".csv", ".inf")
    
    test_cases = project_wd +'\\'+'TestCases.txt'        # The name of the text file which contains the start time and stop time of each case.
    criteria = os.path.join(plot_folder, 'Criterions.txt')                         # The name of the text file which contains the criterions for each case.
    report_content = os.path.join(plot_folder, 'text_for_report.txt')
    
    TestCases = pd.ExcelFile(project_wd +"\\"+"TestCases.xlsx")
    Inputs = pd.read_excel(TestCases, sheet_name='Input')
    
    project_name = Inputs['ProjectName'][0]             # This is the name of the project.
    total_cases = int(Inputs['TotalCases'][0])          # This is the number of cases you want to run the code for.
    Pn = Inputs['Pn (MW)'][0]                           # (MW) Nominal active power
    
        
    merge_emt_csv_files(
        main_path = data_folder + "\\emt",          # This is the path to where the emt data is located.
        num_cases = total_cases  
        )
    
    
    if plot_flag is True:
        
        result_folder = run_all(
            number_of_cases = total_cases,  
            src_folder = data_folder,  
            projectname = project_name,   
            txt_time = test_cases,  
            txt_cases = criteria,  
            Pn_model = Pn,    
            RMS = RMS_flag,  
            EMT = EMT_flag,  
            )
    
    
    if report_flag is True:
        
        today = date.today().strftime("%Y%m%d")
        
        if RMS_flag is True and EMT_flag is True:
            result_folder_name = 'Result_RMS_and_EMT' 
        elif RMS_flag is True:
            result_folder_name = 'Result_RMS' 
        elif EMT_flag is True:
            result_folder_name = 'Result_EMT' 
        
        result_folder = os.path.join(data_folder, result_folder_name)
        
        create_document(
            header_text = 'Model Validation Report',  # The text at the top of the document.
            num_cases = total_cases,  # This is the number of cases you want to run the code for.
            txt_file = report_content,  # The name of the text file with text for the report.
            path_png_files = result_folder,  # The path to the png files to use for the report generation
            project_name = project_name,  # This is the name of the project. 
            save_path = data_folder,  # The path to where the report should be saved.
            document_name = project_name+'_Energinet_'+today  # The name of the saved report.
            )
