import pandas as pd
import json
from datetime import datetime
import re


class Student_data:
    def __init__(self, stu_file_paths) -> None:
        self.stu_file_paths = stu_file_paths  # List of Excel file paths
        self.stu_main()

    def _read_path(self, path) -> None:
        self.xls = pd.ExcelFile(path)
        self.sheet_names = self.xls.sheet_names

    def extract_schedule_calendar_blocks(self, file_path):
        results = []
        self._read_path(file_path)

        for sheet in self.sheet_names:
            df = self.xls.parse(sheet, header=None)
            student_name = sheet

            for date_idx in range(2, 49, 8): 
                month_row = df.iloc[date_idx]
                day_row = df.iloc[date_idx]

                for col in range(1, len(day_row), 2):
                    class_date = None
                    # Extract and normalize month from column 0
                    month_cell = str(month_row[0]) if not pd.isna(month_row[0]) else ""
                    month_match = re.search(r"([0-9０-９]{1,2})月", month_cell)
                    if not month_match:
                        continue
                    month_str = month_match.group(1).translate(str.maketrans("０１２３４５６７８９", "0123456789"))
                    current_month = int(month_str)

                    # Extract day from column
                    raw_day = str(day_row[col]) if not pd.isna(day_row[col]) else ""
                    if "日" not in raw_day:
                        continue
                    try:
                        day = int(re.sub(r"[^\d]", "", raw_day))
                        class_date = datetime(2025, current_month, day).date()
                    except ValueError:
                        continue
                        # Skip processing if class_date is invalid (e.g., raw_day is empty)
                    if not class_date:
                        continue

                    # Check rows 4 to 8 for time and class info
                    for row_idx in range(date_idx + 2, date_idx + 7):

                        if row_idx >= len(df):
                            continue

                        time_val = df.iloc[row_idx, 0]
                        if pd.isna(time_val):
                            continue
                        time_match = re.match(r"\d{1,2}:\d{2}", str(time_val))
                        if not time_match:
                            continue
                        time_str = time_match.group(0)
                        subject_val = df.iloc[row_idx, col] if col < df.shape[1] else None
                        form_val = df.iloc[row_idx, col + 1] if (col + 1) < df.shape[1] else None

                        if pd.isna(subject_val) or str(subject_val).strip() == "×":
                            continue

                        subject = str(subject_val).strip()
                        subject_form = str(form_val).strip() if not pd.isna(form_val) else ""

                        if subject and subject_form != "×":
                            results.append({
                                "Student name": student_name,
                                "Date": class_date.isoformat(),
                                "Time": time_str,
                                "Subject": subject,
                                "Subject form": subject_form
                            })

        return results

    def stu_main(self):
        all_data = []
        for path in self.stu_file_paths:
            data = self.extract_schedule_calendar_blocks(path)
            all_data.extend(data)

        with open("student_schedules.json", "w", encoding="utf-8") as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)


# Example usage:
# Student_data(["path_to_your_excel_file.xlsx"])
