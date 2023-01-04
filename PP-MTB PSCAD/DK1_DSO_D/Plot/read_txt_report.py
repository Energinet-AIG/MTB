

def create_case_list(input_file):
    """
    This function generates a list with line numbers of when each case begins in the text document.

    :param input_file: The text document with the text to add to each case in the document.
    :return: A list with line numbers on when each case starts.
    """
    
    list_case_start = []
    line_num = 0
    with open(input_file, "r") as input_file:
        for line in input_file:
            line_num += 1
            if '#' in line:
                list_case_start.append(line_num)

    return list_case_start


def create_case_dict(input_file):
    """
    This function generates a dictionary with case number as key and the corresponding
    text from the text file for that case as value.
    
    :param input_file: The text document with the text to add to each case in the document.
    :return: A dictionary with case number as key and a list of the text from the 
    text file belonging to that case as value.
    """

    list_case_start = create_case_list(input_file)

    case_dict = {}

    with open(input_file, "r") as input_file:
        lines_list = input_file.readlines()
        lines_list_stripped = [string.rstrip() for string in lines_list]

    for list_num, line_num in enumerate(list_case_start):

        key = int(''.join(filter(str.isdigit, lines_list_stripped[line_num])))
        
        end_line = -1 if list_num == (len(list_case_start) - 1) else list_case_start[list_num + 1]

        value = [line for line in lines_list_stripped[line_num + 1 : end_line] if not (line == '' or '#' in line)]
        
        index_text_above = value.index('Text above figure:')
        index_text_below = value.index('Text below figure:')
        
        value = [[string for string in value[index_text_above:index_text_below]], [string for string in value[index_text_below:]]]

        case_dict[key] = value

    return case_dict


def text_to_figure(list_text):
    """
    This function merges the text together and returns the text to be before and 
    after the figure.
    
    :param list_text: A list with text to go before and after the graph for 
    a specific case.
    :return: Two strings where the first string is the text to place before the 
    figure in the document and the second string is the text to place after the 
    figure in the document.
    """
    
    list_above, list_below = list_text[0], list_text[1]
    
    text_above = ' '.join(list_above[1:])
    text_below = ' '.join(list_below[1:])

    return text_above, text_below


def get_text(case_num, case_dict):
    """
    This function extracts and return the text for the specified case.
    
    :param case_num: The case number to get the text for.
    :param case_dict: The dictionary with the text for each case.
    :param return: Two strings where the first string is the text to place before the 
    figure in the document and the second string is the text to place after the 
    figure in the document.
    """

    list_txt = case_dict.get(case_num)

    text_above, text_below = text_to_figure(list_txt)

    return text_above, text_below


