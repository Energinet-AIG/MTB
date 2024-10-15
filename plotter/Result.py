from enum import Enum


class ResultType(Enum):
    RMS = 0
    EMT = 1


class Result:
    def __init__(self, typ : ResultType, rank : int, projectName : str, bulkname : str, fullpath : str, group : str) -> None:
        self.typ = typ
        self.rank = rank
        self.projectName = projectName
        self.bulkname = bulkname
        self.fullpath = fullpath
        self.group = group
        self.shorthand = f'{group}\\{projectName}'
