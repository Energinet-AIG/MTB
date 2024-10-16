from __future__ import annotations
import pandas as pd
from plotly.subplots import make_subplots  # type: ignore
import plotly.graph_objects as go  # type: ignore
from typing import List, Tuple
import plot_cursor_functions
from Result import ResultType
from Rank import Rank
from math import ceil


def addCursors(htmlPlots: List[go.Figure],
               typ: ResultType,
               data: pd.DataFrame,
               rankDict: List[Rank],
               pfFlatTIme: float,
               pscadInitTime: float,
               rank: int,
               nColumns: int):
    rank_settings = [i for i in rankDict if i.id == rank]
    if len(rank_settings) == 0:
        return list()

    signalKey = typ.name.lower()

    # Initialize subplot positions
    fi = -1  # Start index from -1 as it is incremented before use

    # Loop through rank settings
    for rank_setting in rank_settings:
        rawSigNames = getattr(rank_setting, f'{signalKey}_signals')
        threshold = getattr(rank_setting, 'cursor_time_ranges')
        cursor_options = getattr(rank_setting, 'cursor_options')
        # Increment plot index
        fi += 1

        if len(rawSigNames) == 0:
            continue

        # Select the correct plot
        plot = htmlPlots[fi] if nColumns == 1 else htmlPlots[0]

        x = []
        y = []

        for rawSigName in rawSigNames:
            if typ == ResultType.RMS:
                # Remove hash and split signal name
                while rawSigName.startswith('#'):
                    rawSigName = rawSigName[1:]
                splitSigName = rawSigName.split('\\')

                if len(splitSigName) == 2:
                    sigColumn = ('##' + splitSigName[0], splitSigName[1])
                else:
                    sigColumn = rawSigName
            else:
                sigColumn = rawSigName

            # Determine the time column and offset based on the type
            timeColName = 'time' if typ == ResultType.EMT else data.columns[0]
            timeoffset = pfFlatTIme if typ == ResultType.RMS else pscadInitTime

            if sigColumn in data.columns:
                # Get the signal data and time values
                x.extend(data[timeColName] - timeoffset)  # type: ignore
                y.extend(data[sigColumn])  # type: ignore

        # Filter the data based on the threshold
        if len(y) != 0:
            x = pd.Series(x)
            y = pd.Series(y)

            if len(threshold) > 0:
                assert len(threshold) <= 2
                mask = (x >= threshold[0]) & (x < threshold[1]) if len(threshold) == 2 else (x >= threshold[0])
                y_filtered = y[mask]
                x_filtered = x[mask]
            else:
                y_filtered = y
                x_filtered = x
            index_number = fi if nColumns != 1 else 0
            plot_cursor_functions.add_text_subplot(plot, x_filtered, y_filtered, cursor_options, index_number)

    return htmlPlots


def setupPlotLayoutCursors(config, ranksCursor: List, htmlPlots: List[go.Figure],
                           imagePlots: List[go.Figure]):
    lst: List[Tuple[int, List[go.Figure]]] = []

    if config.genHTML:
        lst.append((config.htmlColumns, htmlPlots))
    if config.genImage:
        lst.append((config.imageColumns, imagePlots))

    for columnNr, plotList in lst:
        if columnNr == 1:
            for rankCursor in ranksCursor:
                # Prepare cursor data for the table
                cursor_data = rankCursor.cursor_data if hasattr(rankCursor, 'cursor_data') else [
                    {'type': 'None', 'value': 'None'}]

                # Prepare data for the table
                table_data = [
                    [cursor['type'], cursor['value']] for cursor in cursor_data
                ]

                # Create the table
                table = go.Table(
                    header=dict(values=["Cursor type", "Values"], fill_color='paleturquoise', align='left'),
                    cells=dict(values=list(zip(*table_data)), fill_color='lavender', align='left')
                )

                # Create a figure to contain the table
                fig_table = go.Figure(data=[table])
                fig_table.update_layout(title=rankCursor.title, height=500)  # Set height for each table
                plotList.append(fig_table)

        elif columnNr > 1:
            num_rows = ceil(len(ranksCursor) / columnNr)
            titles = [rankCursor.title for rankCursor in ranksCursor]  # Gather titles for each table

            # Create subplots specifically for tables
            fig_subplots = make_subplots(rows=num_rows, cols=columnNr,
                                         subplot_titles=titles,
                                         specs=[[{'type': 'table'} for _ in range(columnNr)] for _ in
                                                range(num_rows)])  # Define all as table subplots

            for i, rankCursor in enumerate(ranksCursor):
                # Prepare cursor data for the table
                cursor_data = rankCursor.cursor_data if hasattr(rankCursor, 'cursor_data') else [
                    {'type': 'None', 'value': 'None'}]

                # Prepare data for the table
                table_data = [
                    [cursor['type'], cursor['value']] for cursor in cursor_data
                ]

                # Create the table
                table = go.Table(
                    header=dict(values=["Cursor type", "Values"], fill_color='paleturquoise', align='left'),
                    cells=dict(values=list(zip(*table_data)), fill_color='lavender', align='left')
                )

                # Add table to the subplot layout
                fig_subplots.add_trace(table, row=i // columnNr + 1, col=i % columnNr + 1)

            # Update the layout of the subplot figure
            fig_subplots.update_layout(height=500 * num_rows)  # Adjust height based on number of rows

            plotList.append(fig_subplots)