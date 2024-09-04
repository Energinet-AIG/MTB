import enum
from typing import List, Dict


def create_min_max_annotations(x, y, subplot_ref, existing_annotations: List[Dict] = None) -> List[Dict]:
    if existing_annotations is None:
        annotations = []
    else:
        annotations = existing_annotations
    min_y = y.min()
    max_y = y.max()
    min_x = x[y.idxmin()]
    max_x = x[y.idxmax()]
    annotations.extend([
        dict(
            x=min_x,
            y=min_y,
            xref=subplot_ref['xref'],
            yref=subplot_ref['yref'],
            text=f"Min: ({min_x}, {min_y})",
            showarrow=True,
            arrowhead=7,
            ax=0,
            ay=-40
        ),
        dict(
            x=max_x,
            y=max_y,
            xref=subplot_ref['xref'],
            yref=subplot_ref['yref'],
            text=f"Max: ({max_x}, {max_y})",
            showarrow=True,
            arrowhead=7,
            ax=0,
            ay=-40
        )
    ])
    return annotations

def add_annotations(x, y, figure, row_number, column_number):
    subplot_ref = {
        'xref': f'x{column_number}',
        'yref': f'y{row_number}'
    }
    annotations = create_min_max_annotations(x, y, subplot_ref)
    for i in annotations:
        figure.add_annotation(i)