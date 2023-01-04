import re
from unicodedata import name
import pandas as pd



def dict_columns_EMT(input_file):
    """
    This function generates a dictionary where the key is the column number
    of where the specific column is located in the file.
    
    :param input file: An inf file with the specifications to use.
    :return: A dictionary with column number as key and column name as value.
    """
    
    dict_names = {}

    with open(input_file, "r") as input_file:
        for line in input_file:
            
            list_line = line.split()

            list_column_num = [string for string in list_line if "PGB" in string]
            column_num = re.sub("[^0-9]", "", list_column_num[0])
            
            list_name = [string for string in list_line if "Desc" in string]
            name = list_name[0].split('=')[1].replace('"', '')
            
            dict_names[name] = column_num
    
    return dict_names


def columns_EMT(input_file):
    """
    
    :param input file: An inf file with the specifications to use.
    :return: A list with the column numbers for the columns to plot in the right order.
    """
    
    dict_names = dict_columns_EMT(input_file)

    list_names = ['P_pos_pu_PoC_ENDK', 'f_ENDK', 'Q_pos_pu_PoC_ENDK', 'Q_neg_pu_PoC_ENDK',
                  'Id_pos_pu_PoC_ENDK', 'Id_neg_pu_PoC_ENDK', 'Iq_pos_pu_PoC_ENDK', 'Iq_neg_pu_PoC_ENDK',
                  ['V_A_RMS_pu_PoC_ENDK', 'V_B_RMS_pu_PoC_ENDK', 'V_C_RMS_pu_PoC_ENDK'], 'V_pos_pu_PoC_ENDK', 'V_neg_pu_PoC_ENDK']

    list_columns = []

    for name in list_names:
        if type(name) is list:
            name_list = []
            for each_name in name:
                name_list.append(int(dict_names.get(each_name)))
            list_columns.append(name_list)
        else:
            list_columns.append(int(dict_names.get(name)))
    
    return list_columns
    

def columns_RMS(csv_file):
    """
    This function takes the csv file as input and generates a list of 
    columns to plot based on specified column names.

    :param csv_file:    The csv file with RMS data.
    :return:            A list with column numbers to plot 
    """
    
    df = pd.read_csv(csv_file, sep=';', decimal=",", header=1)

    list_columns = []
    
    list_columns.append(df.columns.get_loc('s:p in p.u.'))
    list_columns.append(df.columns.get_loc('m:fehz in Hz'))
    list_columns.append(df.columns.get_loc('s:q in p.u.'))
    list_columns.append(df.columns.get_loc('s:q2 in p.u.'))
    list_columns.append(df.columns.get_loc('m:i1P:bus2 in p.u.'))
    list_columns.append(df.columns.get_loc('m:i2P:bus2 in p.u.'))
    list_columns.append(df.columns.get_loc('m:i1Q:bus2 in p.u.'))
    list_columns.append(df.columns.get_loc('m:i2Q:bus2 in p.u.'))
    list_columns.append([df.columns.get_loc('m:u:bus2:A in p.u.'), df.columns.get_loc('m:u:bus2:B in p.u.'), df.columns.get_loc('m:u:bus2:C in p.u.')])
    
    list_columns.append(df.columns.get_loc('m:u1:bus2 in p.u.'))
    list_columns.append(df.columns.get_loc('m:u2:bus2 in p.u.'))

    return list_columns


def df_columns_rms(df_rms, csv_file, Pn_model):
        
    list_columns = columns_RMS(csv_file)

    df_rms.iloc[:, int(list_columns[10])] = df_rms.iloc[:, int(list_columns[10])]
    
    return df_rms, list_columns
