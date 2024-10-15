import plotly.graph_objects as go
from plotly.subplots import make_subplots
from typing import List, Dict


# Create Min, Max, and Mean Annotations
def create_min_max_annotations(x, y, existing_annotations: List[Dict] = None) -> List[Dict]:
    if existing_annotations is None:
        annotations = []
    else:
        annotations = existing_annotations

    # Calculate Min, Max, Mean
    min_y = y.min()
    max_y = y.max()
    mean_y = sum(y) / len(y)

    # Construct the annotation text
    annotation_text = f"Max: {max_y:.2f}<br>Min: {min_y:.2f}<br>Mean: {mean_y:.2f}"

    # Append the text as an annotation
    annotations.append(dict(
        text=annotation_text,  # Text for the annotation
        x=0.5, y=0.5,  # Central positioning within the annotation subplot
        showarrow=False,  # No arrow
        font=dict(size=14),
        bordercolor=None,
        borderwidth=0,
        borderpad=0,
        bgcolor=None,  # Background color for the annotation box
        opacity=1.0  # Transparency
    ))

    return annotations


# Function to append the annotations subplot to the provided figure
def add_annotations_subplot(fig: go.Figure, x, y, rowPos, colPos):
    # Generate the annotations
    annotations = create_min_max_annotations(x, y)

    # Add each annotation to the second row, first column of the subplot
    for annotation in annotations:
        fig.add_annotation(annotation, row=rowPos, col=colPos, x=0.01, y=0.99)

    # Hide the x and y axes for the annotations subplot
    fig.update_xaxes(visible=False, row=rowPos, col=colPos)
    fig.update_yaxes(visible=False, row=rowPos, col=colPos)

    return fig
