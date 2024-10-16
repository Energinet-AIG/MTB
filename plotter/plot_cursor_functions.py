import plotly.graph_objects as go
from typing import List
from cursor_type import CursorType


def min_max_annotation_text(x, y):
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


def mean_annotation_text(y):
    mean_y = sum(y) / len(y)
    annotation_text = f"Mean: {mean_y:.2f} <br>"
    return annotation_text


# Function to append the text as a scatter trace to the provided figure
def add_text_subplot(fig: go.Figure, x, y, cursor_types: List[CursorType], index_number):
    print("length is: ", len(fig.data))
    print("index is: ", index_number)

    # Update the table cell to include the annotation text
    # Assuming the table data is stored in the first data trace of the figure
    table_data = fig.data[index_number].cells.values

    # Access the values for the specific cell in the table
    # The values are arranged in a way that we can access them based on rowPos and colPos
    cursor_type = table_data[0]
    values = table_data[1]  # Assuming the second entry contains the values

    # Append the annotation text to the corresponding value
    updated_values = values[:]
    updated_cursor_type = cursor_type[:]
    index = 0
    if CursorType.MIN_MAX in cursor_types:
        set_or_append_value(updated_cursor_type, index, "Min and Max values")
        set_or_append_value(updated_values, index, min_max_annotation_text(x, y))
        index += 1

    if CursorType.AVERAGE in cursor_types:
        set_or_append_value(updated_cursor_type, index, "Average values")
        set_or_append_value(updated_values, index, mean_annotation_text(y))
        index += 1

    # Update the table with the modified values
    fig.data[index_number].cells.values = [updated_cursor_type, updated_values]

    return fig


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
