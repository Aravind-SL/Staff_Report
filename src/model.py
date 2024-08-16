from dataclasses import dataclass


@dataclass
class WorkBookFile:
    file_name: str
    slow_learners_threshold: int
    advanced_learners_threshold: int

    def __str__(self):
        return self.file_name.split("/")[-1]

    def __hash__(self):
        return hash(self.file_name)

    def set_slow_leaners_threshold(self, val: int):
        self.slow_learners_threshold = val

    def set_advanced_learners_threshold(self, val: int):
        self.advanced_learners_threshold = val


