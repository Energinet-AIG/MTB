# from turtle import color
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px

import numpy as np
import pandas as pd
import re
import os

from read_txt_inputs import get_lines_dict
from get_columns import columns_EMT, df_columns_rms

def add_graph(fig, data_xaxis, data_yaxis, row_num, col_num, line_color, rms_or_emt, showlegend):
    """
    This function adds a graph to the figure at the specified sublocation.
    The data for the graph is specified as inputs for the x- and y-axis.

    :param fig: The figure to add a graph to.
    :param data_xaxis: The data to use on the x axis.
    :param data_yaxis: The data to use on the y axis.
    :param row_num: The row number for the placement of the graph. 
    :param col_num: The column number for the placement of the graph.
    :param line_color: The color of the line to draw.
    :param rms_or_emt: A string indicating if it is EMT or RMS data.
    """

    fig.append_trace(
        go.Scatter(
            x=data_xaxis,
            y=data_yaxis,
            line_color=line_color, 
            name=rms_or_emt,
            legendgroup='group',
            showlegend=showlegend
        ),
        row=row_num, col=col_num
        )


def update_axis(fig, x_title, y_title, x_range, y_range, row_num, col_num):
    """
    This function updates the x- and y-axis to have the specified titles
    and the specified ranges.

    :param fig: The figure to update axes for.
    :param x_title: The title on the x-axis.
    :param y_title: The title on the y-axis.
    :param x_range: The range of the x-axis on the plot, this should be a list like [x_min, x_max]
    :param y_range: The range of the y-axis on the plot, this should be a list like [y_min, y_max]
    :param row_num: The row number for the placement of the graph.
    :param col_num: The column number for the placement of the graph.
    """

    fig.update_xaxes(
        title_text=x_title, 
        range=x_range, 
        row=row_num, 
        col=col_num
        )
    
    fig.update_yaxes(
        title_text=y_title, 
        range=y_range, 
        row=row_num, 
        col=col_num
        )


def one_graph(fig, csv_file, file_col_num, row_num, col_num, x_title, y_title, x_range, y_range, colors, rms_or_emt, showlegend):
    """
    This function draws one graph with the specified data and the specified location.

    :param fig: The figure to add a graph to.
    :param csv_file: The csv file to draw the graphs from.
    :param file_col_num: The number of the column which contains the data to plot.
    :param row_num: The row number for the placement of the graph.
    :param col_num: The column number for the placement of the graph.
    :param x_title: The title on the x-axis.
    :param y_title: The title on the y-axis.
    :param x_range: The range of the x-axis on the plot, this should be a list like [x_min, x_max]
    :param y_range: The range of the y-axis on the plot, this should be a list like [y_min, y_max]
    :param colors: A list of colors to use for the plotting.
    :param rms_or_emt: A string indicating if it is EMT or RMS data.
    :param showlegend: Boolean (True/False) indicating if the legend should be shown for that graph.
    """

    if type(file_col_num) is list:
        for num, color in zip(file_col_num, colors):
            add_graph(
                fig=fig, 
                data_xaxis=csv_file.iloc[:, 0], 
                data_yaxis=csv_file.iloc[:, num], 
                row_num=row_num, 
                col_num=col_num,
                line_color=color,
                rms_or_emt=rms_or_emt,
                showlegend=showlegend
                )
    
    else:
        add_graph(
            fig=fig, 
            data_xaxis=csv_file.iloc[:, 0], 
            data_yaxis=csv_file.iloc[:, file_col_num], 
            row_num=row_num, 
            col_num=col_num,
            line_color=colors[0],
            rms_or_emt=rms_or_emt,
            showlegend=showlegend
            )

    update_axis(
        fig=fig, 
        x_title=x_title, 
        y_title=y_title, 
        x_range=x_range,
        y_range=y_range, 
        row_num=row_num, 
        col_num=col_num)


def initialize_figure():
    """
    This function initialises the figure, i.e. it creates the figure with
    nine subplots (3 by 4). 

    :return: A figure ready to add nine graphs to.
    """

    fig = make_subplots(
        rows=3,
        cols=4
    )

    return fig


def generate_figure(fig, csv_file, time_start, time_stop, list_col_y_axis, colors, rms_or_emt):
    """
    This function is generating all the figures in one plot.

    :param fig:             The figure to add the graphs to.
    :param csv_file:        Csv file to use for plotting.
    :param time_start:      The start time for the x axis on the plot.
    :param time_stop:       The end time for the x axis on the plot.
    :param list_col_y_axis: A list of the column numbers to use for each graph.
    :param colors:          A list of colors to use for the plotting.
    :param rms_or_emt:      A string indicating if it is EMT or RMS data.
    """

    one_graph(
        fig=fig, 
        csv_file=csv_file,
        file_col_num=list_col_y_axis[0], 
        row_num=1, 
        col_num=1, 
        x_title='', 
        y_title='P_pos (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.5, 1.2],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )
    
    one_graph(
        fig=fig, 
        csv_file=csv_file,
        file_col_num=list_col_y_axis[1], 
        row_num=1, 
        col_num=2, 
        x_title='', 
        y_title='f (Hz)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[47, 52],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=True
        )
    
    one_graph(
        fig=fig, 
        csv_file=csv_file, 
        file_col_num=list_col_y_axis[2], 
        row_num=1, 
        col_num=3, 
        x_title='', 
        y_title='Q_pos (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.5, 1.2],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )

    one_graph(
        fig=fig, 
        csv_file=csv_file, 
        file_col_num=list_col_y_axis[3], 
        row_num=1, 
        col_num=4, 
        x_title='', 
        y_title='Q_neg (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.5, 1.2],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )

    one_graph(
        fig=fig, 
        csv_file=csv_file,
        file_col_num=list_col_y_axis[4], 
        row_num=2, 
        col_num=1, 
        x_title='', 
        y_title='Id_pos (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.6, 1.4],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )

    one_graph(
        fig=fig, 
        csv_file=csv_file, 
        file_col_num=list_col_y_axis[5], 
        row_num=2, 
        col_num=2, 
        x_title='', 
        y_title='Id_neg (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.6, 1.4],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )


    one_graph(
        fig=fig, 
        csv_file=csv_file,
        file_col_num=list_col_y_axis[6], 
        row_num=2, 
        col_num=3, 
        x_title='', 
        y_title='Iq_pos (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.6, 1.4],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )

    one_graph(
        fig=fig, 
        csv_file=csv_file, 
        file_col_num=list_col_y_axis[7], 
        row_num=2, 
        col_num=4, 
        x_title='', 
        y_title='Iq_neg (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.6, 1.4],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )


    one_graph(
        fig=fig, 
        csv_file=csv_file,
        file_col_num=list_col_y_axis[8], 
        row_num=3, 
        col_num=1, 
        x_title='Time (s)', 
        y_title='V_abc_RMS_PSCAD (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.25, 1.5],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )
    
    one_graph(
        fig=fig, 
        csv_file=csv_file, 
        file_col_num=list_col_y_axis[8], 
        row_num=3, 
        col_num=2, 
        x_title='Time (s)', 
        y_title='V_abc_RMS_PF (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.25, 1.5],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )
    
    one_graph(
        fig=fig, 
        csv_file=csv_file,
        file_col_num=list_col_y_axis[9], 
        row_num=3, 
        col_num=3, 
        x_title='Time (s)', 
        y_title='V_pos (pu)', 
        x_range=[time_start-0.2, time_stop],
        y_range=[-0.25, 1.5],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )
    
    one_graph(
        fig=fig, 
        csv_file=csv_file, 
        file_col_num=list_col_y_axis[10], 
        row_num=3, 
        col_num=4, 
        x_title='Time (s)', 
        y_title='V_neg (pu)',
        x_range=[time_start-0.2, time_stop], 
        y_range=[-0.25, 1.5],
        colors=colors,
        rms_or_emt=rms_or_emt,
        showlegend=False
        )
    

    return fig


def add_criterion_line(fig, coords_list, row_num, col_num):
    """
    This function adds a criterion line to the graph in the figure on the location (row_num, col_num). 

    :param fig:         The figure to make changes to.
    :param coords_list: A list with the coordinates of the ends of the line to draw, on the form [x1, x2, y1, y2].
    :param row_num:     The row number of the graph to add a line to.
    :param col_num:     The column number of the graph to add a line to.
    """

    fig.append_trace(
        go.Scatter(
            x=[float(coords_list[0]), float(coords_list[1])],
            y=[float(coords_list[2]), float(coords_list[3])],
            mode="lines",
            line={'dash': 'dot', 'color': 'red'},
            showlegend=False
            ), 
        row=row_num, col=col_num
        )


def add_criteria(fig, graph_dict):
    """
    This function draws the criteria on top of the subfigures based on the Criterions.txt file.
    
    :param fig:         The figure to add the criterion lines to.
    :param graph_dict:  A dictionary with graph name and line coordinates.
    :return:            The figure with criterions added.
    """
    
    keys = graph_dict.keys()
    
    if any('P_pos (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'P_pos (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 1, 1)
    
    if any('f (Hz)' in key for key in keys):
        [key_name] = [key for key in keys if 'f (Hz)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 1, 2)

    if any('Q_pos (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'Q_pos (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 1, 3)

    if any('Q_neg (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'Q_neg (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 1, 4)

    if any('Id_pos (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'Id_pos (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 2, 1)
    
    if any('Id_neg (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'Id_neg (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 2, 2)
    
    if any('Iq_pos (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'Iq_pos (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 2, 3)

    if any('Iq_neg (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'Iq_neg (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 2, 4)    
    
    if any('V_abc_RMS_PSCAD (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'V_abc_RMS_PSCAD (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 3, 1)
    
    if any('V_abc_RMS_PF (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'V_abc_RMS_PF (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 3, 2)
    
    if any('V_pos (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'V_pos (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 3, 3)

    if any('V_neg (pu)' in key for key in keys):
        [key_name] = [key for key in keys if 'V_neg (pu)' in key]
        list_lines = graph_dict.get(key_name)
        for line in list_lines:
            add_criterion_line(fig, line, 3, 4)    
   
    return fig
    

def run_all(number_of_cases, src_folder, projectname, txt_time, txt_cases, 
            Pn_model, EMT=True, RMS=True):
    """
    This function loops through all the cases and generates the graphs for each case
    and saves the result as html and png files.

    :param number_of_cases: The total number of cases.
    :param src_folder: The path to the folder with the csv files with data to use.
    :param projectname: The name of the project - this should match the first part
    of the name of the csv files.
    :param txt_time_start: The text file containing the start time.
    :param txt_time_stop: The text file containing the end time.
    :param txt_cases: The path to the text_file containing the coordinates
    for the line to draw in each case for each graph.
    :param Pn_model: Parameter to use for RMS.
    :param EMT: Boolean which indicates if the EMT data should be plotted or not (True or False).
    :param RMS: Boolean which indicates if the RMS data should be plotted or not (True or False).
    """

    for case_num in range(1, number_of_cases + 1):
        
        fig = initialize_figure()
        
        time_stop = np.loadtxt(txt_time, delimiter=',')[(case_num - 1), 14]
        time_start = np.loadtxt(txt_time, delimiter=',')[(case_num - 1), 15]
        
        
        if EMT is True:
            if case_num < 10:
                df_emt = pd.read_csv(
                    os.path.join(src_folder, 'emt', projectname + "_0" + str(case_num) + ".csv")
                    )
            else:
                df_emt = pd.read_csv(
                    os.path.join(src_folder, 'emt', projectname + "_" + str(case_num) + ".csv")
                    )

            list_col_y_axis = columns_EMT(
                os.path.join(src_folder, 'emt', projectname + "_01.inf"))
            
            generate_figure(
                fig=fig,
                csv_file=df_emt, 
                time_start=time_start, 
                time_stop=time_stop, 
                list_col_y_axis=list_col_y_axis, 
                colors=['#1313ad', '#c98a1c', '#14a627'] if RMS is False else ['#1313ad', '#1313ad', '#1313ad'],
                rms_or_emt='EMT'
                )
            
        
        if RMS is True:
            if case_num < 10:
                df_rms = pd.read_csv(
                    os.path.join(src_folder, 'rms', projectname + "_0" + str(case_num) + ".csv"), 
                    sep=';',
                    decimal=',',
                    engine='python', 
                    header=1
                    )
                df_rms, list_col_y_axis = df_columns_rms(
                    df_rms,
                    os.path.join(src_folder, 'rms', projectname + "_0" + str(case_num) + ".csv"),
                    Pn_model=Pn_model
                    )
           
            else:
                df_rms = pd.read_csv(
                    os.path.join(src_folder, 'rms', projectname + "_" + str(case_num) + ".csv"), 
                    sep=';',
                    decimal=',',
                    engine='python', 
                    header=1
                    )
                df_rms, list_col_y_axis = df_columns_rms(
                    df_rms,
                    os.path.join(src_folder, 'rms', projectname + "_" + str(case_num) + ".csv"),
                    Pn_model=Pn_model
                    )
            
            generate_figure(
                fig=fig,
                csv_file=df_rms, 
                time_start=time_start, 
                time_stop=time_stop, 
                list_col_y_axis=list_col_y_axis,
                colors=['#ff6e00'],
                rms_or_emt='RMS'
                )
            

        add_criteria(fig, graph_dict=get_lines_dict(case_num, txt_cases))
        
        
        if RMS is True and EMT is True:
            fig.update_layout()
        else:
            fig.update_layout(showlegend=False)
        
        
        if RMS is True and EMT is True:
            result_folder_name = 'Result_RMS_and_EMT' 
        elif RMS is True:
            result_folder_name = 'Result_RMS' 
        elif EMT is True:
            result_folder_name = 'Result_EMT' 
            
        result_folder = os.path.join(src_folder, result_folder_name)
        if not os.path.exists(result_folder):
            os.mkdir(result_folder)

        if case_num < 10:
            fig.write_image(os.path.join(result_folder, projectname + "_0" + str(case_num) + '.png'), width=2000, height=1000)
            fig.write_html(os.path.join(result_folder, projectname + "_0" + str(case_num) + '.html'))
            continue
        else:
            fig.write_image(os.path.join(result_folder, projectname + "_" + str(case_num) + '.png'), width=2000, height=1000)
            fig.write_html(os.path.join(result_folder, projectname + "_" + str(case_num) + '.html'))
            continue
        
    return result_folder
        
