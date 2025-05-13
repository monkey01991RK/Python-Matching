import json
from collections import defaultdict
import pandas as pd
import os
from dataclasses import dataclass, asdict
from typing import List, Optional


@dataclass
class Subject:
    name: str
    teacher: Optional[str]
    regular_classes: Optional[int]
    special_classes: Optional[int]


@dataclass
class Student:
    grade: str
    student_name: str
    subjects: List[Subject]


class match_basic:
    def __init__(self, match_file_path) -> None:
        self.match_file_path = match_file_path  # Single file path
        self.match_main()

    def _read_path(self, path) -> None:
        self.xls = pd.ExcelFile(path)
        self.sheet_names = self.xls.sheet_names

    def extract_schedule_match_blocks(self, file_path):
        results = []
        self._read_path(file_path)

        for sheet_name in self.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            for _, row in df.iterrows():
                student_name = row.get("生徒名")
                grade = row.get("学年", "不明")

                if pd.isna(student_name):
                    continue

                subjects = []
                subject_columns = [
                    ("教科", "講師", "通常コマ数", "講習コマ数"),
                    ("教科.1", "講師.1", "通常コマ数.1", "講習コマ数.1"),
                    ("教科.2", "講師.2", "通常コマ数.2", "講習コマ数.2"),
                    ("教科.3", "講師.3", "通常コマ数.3", "講習コマ数.3"),
                    ("教科.4", "講師.4", "通常コマ数.4", "講習コマ数.4"),
                ]

                for cols in subject_columns:
                    sub_name = row.get(cols[0])
                    if pd.notna(sub_name):
                        subject = Subject(
                            name=sub_name,
                            teacher=row.get(cols[1]) if pd.notna(row.get(cols[1])) else None,
                            regular_classes=int(row.get(cols[2]) or 0),
                            special_classes=int(row.get(cols[3]) or 0)
                        )
                        subjects.append(subject)

                student = Student(
                    grade=grade,
                    student_name=student_name,
                    subjects=subjects
                )
                results.append(student)

        return results

    def match_main(self):
        if os.path.exists(self.match_file_path):
            students = self.extract_schedule_match_blocks(self.match_file_path)
            all_data = [asdict(student) for student in students]
            with open("all_students_schedule.json", "w", encoding="utf-8") as f:
                json.dump(all_data, f, ensure_ascii=False, indent=2)
