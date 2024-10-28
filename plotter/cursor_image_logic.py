from __future__ import annotations
import pandas as pd
from plotly.subplots import make_subplots  # type: ignore
import plotly.graph_objects as go  # type: ignore
from typing import List, Tuple
import plot_cursor_functions
from Result import ResultType
from Cursor import Cursor
from math import ceil
from Result import Result
from read_and_write_functions import loadEMT


def addCursors(htmlPlots: List[go.Figure],
               resultList: List[Result],
               cursorDict: List[Cursor],
               pfFlatTIme: float,
               pscadInitTime: float,
               rank: int,
               nColumns: int):
    cursor_settings = [i for i in cursorDict if i.id == rank]
    if len(cursor_settings) == 0:
        return list()

    # Initialize subplot positions
    fi = -1  # Start index from -1 as it is incremented before use
    for cursor_setting in cursor_settings:
        # Loop through rank settings
        totalRawSigNames = []
        time_ranges = getattr(cursor_setting, 'time_ranges')
        cursor_options = getattr(cursor_setting, 'cursor_options')
        # Increment plot index
        fi += 1

        # Select the correct plot
        plot = htmlPlots[fi] if nColumns == 1 else htmlPlots[0]

        x = []
        y = []
        for result in resultList:
            signalKey = result.typ.name.lower()
            rawSigNames = getattr(cursor_setting, f'{signalKey}_signals')
            totalRawSigNames.extend(rawSigNames)
            data = None
            if result.typ == ResultType.RMS:
                data: pd.DataFrame = pd.read_csv(result.fullpath, sep=';', decimal=',',
                                                       header=[0, 1])  # type: ignore
            elif result.typ == ResultType.EMT:
                data: pd.DataFrame = loadEMT(result.fullpath)
            if len(rawSigNames) == 0:
                continue
            for rawSigName in rawSigNames:
                if result.typ == ResultType.RMS:
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
                timeColName = 'time' if result.typ == ResultType.EMT else data.columns[0]
                timeoffset = pfFlatTIme if result.typ == ResultType.RMS else pscadInitTime

                if sigColumn in data.columns:
                    # Get the signal data and time values
                    x.extend(data[timeColName] - timeoffset)  # type: ignore
                    y.extend(data[sigColumn])  # type: ignore

        # Filter the data based on the time_ranges
        if len(y) != 0:
            x = pd.Series(x)
            y = pd.Series(y)
            index_number = fi if nColumns != 1 else 0
            plot_cursor_functions.add_text_subplot(plot, x, y, cursor_options, index_number, time_ranges, totalRawSigNames)

    return htmlPlots


def setupPlotLayoutCursors(config, ranksCursor: List, htmlPlots: List[go.Figure],
                           imagePlots: List[go.Figure]):
    lst: List[Tuple[int, List[go.Figure]]] = []

    if config.genHTML:
        lst.append((config.htmlCursorColumns, htmlPlots))
    if config.genImage:
        lst.append((config.imageCursorColumns, imagePlots))

    for columnNr, plotList in lst:
        if columnNr == 1:
            for rankCursor in ranksCursor:
                # Prepare cursor data for the table
                table = create_cursor_table()

                # Create a figure to contain the table
                fig_table = go.Figure(data=[table])
                fig_table.update_layout(title=rankCursor.title, height=140*max(len(rankCursor.cursor_options), 1))
                plotList.append(fig_table)

        elif columnNr > 1:
            num_rows = ceil(len(ranksCursor) / columnNr)
            titles = [rankCursor.title for rankCursor in ranksCursor]  # Gather titles for each table

            # Create subplots specifically for tables
            fig_subplots = make_subplots(rows=num_rows, cols=columnNr,
                                         subplot_titles=titles,
                                         specs=[[{'type': 'table'} for _ in range(columnNr)] for _ in
                                                range(num_rows)])  # Define all as table subplots
            height_to_use = 500
            for i, rankCursor in enumerate(ranksCursor):
                # Prepare cursor data for the table
                table = create_cursor_table()

                # Add table to the subplot layout
                fig_subplots.add_trace(table, row=i // columnNr + 1, col=i % columnNr + 1)

                # Update the layout of the subplot figure
                height_to_use = max(500*len(rankCursor.cursor_options), height_to_use)
            fig_subplots.update_layout(height=height_to_use)

            plotList.append(fig_subplots)


def create_cursor_table():
    cursor_data = [{'type': 'None', 'signals': 'None', 'time_values': 'None', 'value': 'None'}]
    # Prepare data for the table, including two additional placeholder columns
    table_data = [
        [cursor['type'], cursor['signals'], cursor['time_values'], cursor['value']] for cursor in cursor_data
    ]
    # Create the table with additional columns in the header and cells
    table = go.Table(
        header=dict(values=["Cursor type", "Signals", "Cursor time points", "Values"],
                    fill_color='paleturquoise', align='left'),
        cells=dict(values=list(zip(*table_data)), fill_color='lavender', align='left')
    )
    return table
