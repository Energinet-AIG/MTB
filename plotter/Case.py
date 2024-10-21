from typing import List, Tuple, Union


class Case:
    def __init__(self, case: 'pd.Series[Union[str, int, float, bool]]') -> None:
        self.rank: int = int(case['Rank'])
        self.RMS: bool = bool(case['RMS'])
        self.EMT: bool = bool(case['EMT'])
        self.Name: str = str(case['Name'])
        self.U0: float = float(case['U0'])
        self.P0: float = float(case['P0'])
        self.Pmode: str = str(case['Pmode'])
        self.Qmode: str = str(case['Qmode'])
        self.Qref0: float = float(case['Qref0'])
        self.SCR0: float = float(case['SCR0'])
        self.XR0: float = float(case['XR0'])
        self.Simulationtime: float = float(case['Simulationtime'])
        self.Events : List[Tuple[str, float, Union[float, str], Union[float, str]]] = []

        index : pd.Index[str] = case.index # type: ignore
        i = 0
        while(True):
            typeLabel = f'type.{i}' if i > 0 else 'type'
            timeLabel = f'time.{i}' if i > 0 else 'time'
            x1Label = f'X1.{i}' if i > 0 else 'X1'
            x2Label = f'X2.{i}' if i > 0 else 'X2'

            if typeLabel in index and timeLabel in index and x1Label in index and x2Label in index:
                try:
                    x1value = float(str(case[x1Label]).replace(' ',''))
                except ValueError:
                    x1value = str(case[x1Label])

                try:
                    x2value = float(str(case[x2Label]).replace(' ',''))
                except ValueError:
                    x2value = str(case[x2Label])

                self.Events.append((str(case[typeLabel]), float(case[timeLabel]), x1value, x2value))
                i += 1
            else:
                break