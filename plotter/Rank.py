from typing import List


class Rank:
    def __init__(self,
                 id: int,
                 title: str,
                 cursor_options: List[str],
                 emt_signals: List[str],
                 rms_signals: List[str],
                 time_ranges: List[int]) -> None:
        self.id = id
        self.title = title
        self.cursor_options = cursor_options
        self.emt_signals = emt_signals
        self.rms_signals = rms_signals
        self.time_ranges = time_ranges