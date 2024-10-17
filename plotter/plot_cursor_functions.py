import plotly.graph_objects as go
from typing import List
from cursor_type import CursorType


def min_max_value_text(x, y, time_ranges):
    if len(time_ranges) > 0:
        mask = (x >= time_ranges[0]) & (x < time_ranges[1]) if len(time_ranges) == 2 else (x >= time_ranges[0])
        y = y[mask]
        x = x[mask]
    # Find the min and max of y
    min_y = y.min()
    max_y = y.max()

    # Find the corresponding x-values
    min_x = x[y.idxmin()]  # x-value where y is minimum
    max_x = x[y.idxmax()]  # x-value where y is maximum

    # Construct the text
    annotation_text = (f"Max: {max_y:.2f} at x = {max_x}<br>"
                       f"Min: {min_y:.2f} at x = {min_x}<br>")
    return annotation_text


def mean_value_text(x, y, time_ranges):
    if len(time_ranges) > 0:
        mask = (x >= time_ranges[0]) & (x < time_ranges[1]) if len(time_ranges) == 2 else (x >= time_ranges[0])
        y = y[mask]
        x = x[mask]
    mean_y = sum(y) / len(y)
    annotation_text = f"Mean: {mean_y:.2f} <br>"
    return annotation_text


def signals_text(rawSigNames):
    rawSigNames_text = ""
    for rawSigName in rawSigNames:
        rawSigNames_text += f"t{rawSigName} "
    return f"Calculated for the signals " + rawSigNames_text


def time_ranges_text(time_ranges):
    time_ranges_text = ""
    for i in range(len(time_ranges)):
        time_ranges_text += f"t{i} is {time_ranges[i]} "
    return f"Time ranges provided are " + time_ranges_text


# Function to append the text as a scatter trace to the provided figure
def add_text_subplot(fig: go.Figure, x, y, cursor_types: List[CursorType], index_number, time_ranges, rawSigNames):
    table_data = fig.data[index_number].cells.values

    # Access the values for the specific cell in the table
    # The values are arranged in a way that we can access them based on rowPos and colPos
    cursor_type = table_data[0]
    signals = table_data[1]
    time_values = table_data[2]
    values = table_data[3]  # Assuming the second entry contains the values

    # Append the annotation text to the corresponding value
    updated_values = values[:]
    updated_signals = signals[:]
    updated_time_values = time_values[:]
    updated_cursor_type = cursor_type[:]
    index = 0
    if CursorType.MIN_MAX in cursor_types:
        set_or_append_cursor_data(updated_cursor_type, updated_signals, updated_time_values, updated_values, index,
                              rawSigNames, time_ranges, "Min and Max values", min_max_value_text(x, y, time_ranges))
        index += 1

    if CursorType.AVERAGE in cursor_types:
        set_or_append_value(updated_cursor_type, index, "Average values")
        set_or_append_cursor_data(updated_cursor_type, updated_signals, updated_time_values, updated_values, index,
                                  rawSigNames, time_ranges, "Average values", mean_value_text(x, y, time_ranges))
        index += 1

    # Update the table with the modified values
    fig.data[index_number].cells.values = [updated_cursor_type, updated_signals, updated_time_values, updated_values]

    return fig


def set_or_append_cursor_data(updated_cursor_type, updated_signals, updated_time_values, updated_values, index,
                              rawSigNames, time_ranges, cursor_type_text, cursor_value_text):
    set_or_append_value(updated_cursor_type, index, cursor_type_text)
    set_or_append_value(updated_signals, index, signals_text(rawSigNames))
    set_or_append_value(updated_time_values, index, time_ranges_text(time_ranges))
    set_or_append_value(updated_values, index, cursor_value_text)


def set_or_append_value(list_to_update, index, value):
    """
    Update a list by setting the value at a specific index if within bounds,
    or appending the value if the index is out of bounds.

    Args:
        list_to_update (list): The list to be updated.
        index (int): The index at which the value should be set or appended.
        value: The value to be set or appended.
    """
    if index >= len(list_to_update):
        # Append if index is out of bounds
        list_to_update.append(value)
    else:
        # Set the value at the specified index
        list_to_update[index] = value
