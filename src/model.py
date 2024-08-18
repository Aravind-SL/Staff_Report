from typing import *
import pandas as pd
from dataclasses import dataclass


@dataclass
class CourseFile:
    file_name: str
    slow_learners_threshold: int
    advanced_learners_threshold: int

    def __str__(self):
        return self.file_name.split("/")[-1]

    def __hash__(self):
        return hash(self.file_name)

    def set_slow_learners_threshold(self, val: int):
        self.slow_learners_threshold = val

    def set_advanced_learners_threshold(self, val: int):
        self.advanced_learners_threshold = val


@dataclass
class ClassReportInput:
    name: str
    course_reports_input: List[CourseFile]

    def __str__(self):
        return self.name

    def __hash__(self):
        return hash(self.name)

@dataclass
class CombinedReport:
    student_summary: pd.DataFrame
    slow_learners: pd.DataFrame
    fast_learners: pd.DataFrame

    def to_excel(path: str):
        pass
