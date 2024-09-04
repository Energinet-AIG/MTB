from enum import Enum


class DownSamplingMethod(Enum):
    GRADIENT = 1
    AMOUNT = 2
    NO_DOWN_SAMPLING = 3

    @classmethod
    def from_string(cls, string):
        try:
            return cls[string.upper()]
        except KeyError:
            raise ValueError(f"{string} is not a valid {cls.__name__}")