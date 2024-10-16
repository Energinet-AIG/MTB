from enum import Enum

class CursorType(Enum):
    MIN_MAX = 1
    AVERAGE = 2

    @classmethod
    def from_string(cls, string : str):
        try:
            return cls[string.upper()]
        except KeyError:
            raise ValueError(f"{string} is not a valid {cls.__name__}")