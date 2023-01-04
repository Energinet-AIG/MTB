import pandas as pd
import numpy as np

import os


def load_csv_files(first_file_path, second_file_path):
    """
    This function loads in the paths to the csv files into pandas dataframes.

    :param first_file_path: The path to the first csv file.
    :param second_file_path: The path to the second csv file.
    :return: Two pandas dataframes.
    """
    
    df1 = pd.read_csv(first_file_path, skiprows = 1)
    df2 = pd.read_csv(second_file_path, skiprows = 1)

    return df1, df2


def splice_data(df1, df2):
    """
    This function merges the two dataframes, such that the first dataframe is 
    the first columns and the second dataframe is added as new columns.

    :param df1: The first dataframe with data from the first csv file.
    :param df2: The second dataframe with data from the second csv file.
    :return: A merged dataframe.
    """

    df2 = df2.iloc[: , 1:]

    df_result = pd.concat([df1, df2], axis=1)

    return df_result


def save_csv(df, save_path, file_name):
    """
    This function saves the given dataframes to a csv files at the given location.

    :param df: The dataframe to save.
    :param save_path: The path to where the dataframe should be saved.
    :param file_name: The file name when saving the dataframe
    """

    df.to_csv(os.path.join(save_path, '{}.csv'.format(file_name)), index=False)


def remove_csvs(first_file_path, second_file_path):
    """
    This function removes the the given files.

    :param first_file_path: The path to the first csv file.
    :param second_file_path: The path to the second csv file.
    """

    os.remove(first_file_path)
    os.remove(second_file_path)


def merge_delete(main_path, first_file_name, second_file_name):
    """
    This function reads the csv files into dataframes, merges them,
    saves the merged dataframe and removes to other two dataframes.

    :param main_path: The path to where all the files are located.
    :param first_file_name: The name of the first csv file.
    :param second_file_name: The name of the second csv file.
    """

    first_file_path = os.path.join(main_path, first_file_name)
    second_file_path = os.path.join(main_path, second_file_name)

    df1, df2 = load_csv_files(first_file_path, second_file_path)

    df_result = splice_data(df1, df2)

    file_name = first_file_name.strip('.csv')[:-3]
    save_csv(df_result, main_path, file_name)

    remove_csvs(first_file_path, second_file_path)


def list_cases(num_cases):
    """
    This function generates a list with the case numbers on the form '_10_'.

    :param num_cases: Number of test cases.
    :return: A list with the all the numbers of the test cases.
    """

    return ['_0{}_'.format(case_num) if case_num < 10 else '_{}_'.format(case_num) for case_num in range(1, num_cases + 1)]


def create_files_list(main_path):
    """
    This function generates a list with all csv files in the specified path.

    :param main_path: The path to where all the files are located.
    :return: A list with the file names of all csv files at the given path.
    """

    list_files = []
    for file in os.listdir(main_path):
        if file.endswith('.csv'):
            list_files.append(file)
    
    return list_files


def dict_cases(list_files, list_num_cases):
    """
    This function generates a dictionary with the dictionary as the key
    and the files belonging to that case as a list of values.

    :param list_files: A list of all csv files in the directory.
    :param list_num_cases: A list of the case numbers.
    :return: A dictionary with case number as key and a list of the 
    files belonging to this case as value.
    """

    return {case_num: [file for file in list_files if case_num in file] for case_num in list_num_cases}


def merge_emt_csv_files(main_path, num_cases):
    """
    This function runs through all the cases and merges the two csv files
    into one csv file. After merging the previous two files are deleted.

    :param main_path: The path to the folder where the csv files to merge is placed.
    :param num_cases: The number of cases to merge the files for.
    """

    list_num_cases = list_cases(num_cases)

    list_files = create_files_list(main_path)
    
    cases_dict = dict_cases(list_files, list_num_cases)

    for value_list in list(cases_dict.values()):
        
        [first_file] = [file for file in value_list if '_01.csv' in file]
        [second_file] = [file for file in value_list if '_02.csv' in file]
        
        merge_delete(
            main_path=main_path,
            first_file_name=first_file,
            second_file_name=second_file
        )

