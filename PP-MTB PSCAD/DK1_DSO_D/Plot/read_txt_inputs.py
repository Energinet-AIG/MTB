import re


def case_start_list(input_file):
    """
    This function creates a list with the line number of where each case
    starts in the text document.

    :param input_file: The text file containing the lines to add to the graphs.
    :return: A list with the line numbers of where each case starts.
    """
    
    line_num = 0
    list_case_start = []

    with open(input_file, "r") as input_file:
        for line in input_file:
            line_num += 1
            if 'Case' in line:
                list_case_start.append(line_num)

    return list_case_start


def dict_cases(input_file):
    """
    This function creates a dictionary where the key is specifying the case and
    the value is a list of all the lines in the text file belonging to that case.

    :param input_file: The text file containing the lines to add to the graphs.
    :return: A list with key-value-pairs, where keys are the case number
    and the values is a list of the lines in the text file belonging to that case. 
    """

    list_case_start = case_start_list(input_file)

    case_dict = {}

    with open(input_file, "r") as input_file:
        lines_list = input_file.readlines()
        lines_list_stripped = [string.rstrip() for string in lines_list]

    for list_num, line_num in enumerate(list_case_start):

        key = lines_list_stripped[line_num - 1]
        
        end_line = -1 if list_num == (len(list_case_start) - 1) else list_case_start[list_num + 1]

        value = [line for line in lines_list_stripped[line_num - 1 : end_line]]

        case_dict[key] = value

    return case_dict


def locate_lines(case_num, case_dict):
    """
    This function is locating the case number in the dictionary and
    retuning the list of lines from the text file belonging to that case.

    :param case_num: The number of cases.
    :param case_dict: The dictionary containing case numbers as keys and
    a list of lines from the text file belonging to that case.
    :return: The list of lines from the text file belonging to the specified case.
    """

    key = 'Case {}:'.format(case_num)

    return case_dict.get(key)


def read_coordinates(list_lines):
    """
    This function sprippes the given list such that it only contains 
    numbers, . and -

    :param list_lines: A list of lines from the text file for a specific case.
    :return: The input list stripped to only include numbers, . and -
    """

    list_of_coords = []

    for line in list_lines:
        coords = [coord for coord in re.sub("[^0-9|.|-]", " ", line).split()]
        list_of_coords.append(coords)

    return list_of_coords



def bundling(case_list):
    """
    This function creates a dictionary with the title of the y-axis as
    key and a list of of the line coordinates from the text file as value.
    Only the titles of the y-axis which contains coordinates is added to 
    the dictionary.

    :param case_list: A list of lines from the text file for a specific case.
    :return: A dictionary containing the coordinates for the lines to add
    in specific graphs. The key is the title of the y-axis in the graph
    and the value is then a list of coordinates to draw the line(s) from.
    """

    graph_names = [graph_name for graph_name in case_list if '--' in graph_name]

    graph_dict = {}

    for index, string in enumerate(case_list[:-1]):
        if any(graph_name in string for graph_name in graph_names):
            if case_list[index + 1] == '':
                continue
            else:
                extra_line = 1
                list_lines = []
                while 'line' in case_list[index + extra_line]:
                    list_lines.append(case_list[index + extra_line])
                    extra_line += 1
                graph_dict[string] = read_coordinates(list_lines)

    return graph_dict


def get_lines_dict(case_num, txt_input_file):
    """
    This function creates a dictionary for the specified case with the
    titles on the y-axis as key and a list of coordinates for the lines
    to draw as value. 

    :param case_num: The case number to generate the dictionary with line
    coordinates for.
    :param txt_input_file: The text file containing the lines to add to the graphs.
    :return: A dictionary for a specific case containing the coordinates for the 
    lines to add in specific graphs. The key is the title of the y-axis in the graph
    and the value is then a list of coordinates to draw the line(s) from.
    """

    case_list = locate_lines(case_num, dict_cases(txt_input_file))

    graph_dict = bundling(case_list)

    return graph_dict
