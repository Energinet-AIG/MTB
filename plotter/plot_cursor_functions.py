import plotly.graph_objects as go
from typing import List, Dict
from cursor_type import CursorType


# Create Min, Max, and Mean Annotations
def create_annotations(x, y, cursor_types : List[CursorType]) -> Dict:
    annotation_text = ""
    if CursorType.MIN_MAX in cursor_types:
        annotation_text += min_max_annotation_text(x, y)
    if CursorType.AVERAGE in cursor_types:
        annotation_text += mean_annotation_text(y)

    # Append the text as an annotation
    return dict(
        text=annotation_text,  # Text for the annotation
        x=0.5, y=0.5,  # Central positioning within the annotation subplot
        showarrow=False,  # No arrow
        font=dict(size=14),
        bordercolor=None,
        borderwidth=0,
        borderpad=0,
        bgcolor=None,  # Background color for the annotation box
        opacity=1.0  # Transparency
    )


def min_max_annotation_text(x, y):
    # Find the min and max of y
    min_y = y.min()
    max_y = y.max()

    # Find the corresponding x-values
    min_x = x[y.idxmin()]  # x-value where y is minimum
    max_x = x[y.idxmax()]  # x-value where y is maximum

    # Construct the annotation text
    annotation_text = (f"Max: {max_y:.2f} at x = {max_x}<br>"
                       f"Min: {min_y:.2f} at x = {min_x}<br>")

    return annotation_text



def mean_annotation_text(y):
    mean_y = sum(y) / len(y)
    annotation_text = f"Mean: {mean_y: .2f} <br>"
    return annotation_text


# Function to append the annotations subplot to the provided figure
def add_annotations_subplot(fig: go.Figure, x, y, rowPos, colPos, cursor_types : List[CursorType]):
    # Generate the annotations
    annotation = create_annotations(x, y, cursor_types)

    fig.add_annotation(annotation, row=rowPos, col=colPos, x=0.01, y=0.99)

    # Hide the x and y axes for the annotations subplot
    fig.update_xaxes(visible=False, row=rowPos, col=colPos)
    fig.update_yaxes(visible=False, row=rowPos, col=colPos)

    return fig
