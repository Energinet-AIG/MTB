from down_sampling_method import DownSamplingMethod
from typing import List


class Figure:
    def __init__(self,
                 id: int,
                 title: str,
                 units: str,
                 emt_signal_1: str,
                 emt_signal_2: str,
                 emt_signal_3: str,
                 rms_signal_1: str,
                 rms_signal_2: str,
                 rms_signal_3: str,
                 gradient_threshold: float,
                 down_sampling_method: DownSamplingMethod,
                 include_in_case: List[int],
                 exclude_in_case: List[int]) -> None:
        self.id = id
        self.title = title
        self.units = units
        self.emt_signal_1 = emt_signal_1
        self.emt_signal_2 = emt_signal_2
        self.emt_signal_3 = emt_signal_3
        self.rms_signal_1 = rms_signal_1
        self.rms_signal_2 = rms_signal_2
        self.rms_signal_3 = rms_signal_3
        self.gradient_threshold = float(gradient_threshold)
        self.down_sampling_method = down_sampling_method
        self.include_in_case: List[int] = include_in_case
        self.exclude_in_case: List[int] = exclude_in_case