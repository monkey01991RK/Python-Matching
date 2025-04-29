import pandas as pd
from datetime import datetime
import re
import json

class Teacher_data:
    def __init__(self, stu_file_path) -> None:
        self.stu_file_path = stu_file_path
        self.teach_main()

    def _read_path(self, path) -> None:
        self.xls = pd.ExcelFile(path)
        self.sheet_names = self.xls.sheet_names

    def extract_schedule_calendar_blocks(self, file_path):
        raw_results = []
        self._read_path(file_path)

        for sheet in self.sheet_names:
            df = self.xls.parse(sheet, header=None)
            if df.empty or df.shape[0] < 10:
                continue

            # Extract teacher name
            teach_name = sheet if pd.isna(df.iloc[0, 7]) else str(df.iloc[0, 7]).strip()

            # For each block of 8 rows
            for date_idx in range(2, len(df), 8):
                if date_idx + 6 >= len(df):
                    continue

                # Get month from the leftmost cell
                month_cell = str(df.iloc[date_idx, 0])
                month_match = re.search(r"([0-9０-９]{1,2})月", month_cell)
                if not month_match:
                    continue

                month_str = month_match[1].translate(str.maketrans("０１２３４５６７８９", "0123456789"))
                current_month = int(month_str)

                # Loop through date columns (2, 4, 6, ...)
                for col in range(2, df.shape[1], 2):
                    raw_day = str(df.iloc[date_idx, col])
                    if "日" not in raw_day and not re.search(r"\d{1,2}", raw_day):
                        continue
                    try:
                        day = int(re.sub(r"[^\d]", "", raw_day))
                        class_date = datetime(2025, current_month, day).date()
                    except ValueError:
                        continue
                    # Time and schedule rows (rows +2 to +6)
                    for row_idx in range(date_idx + 2, date_idx + 7):
                        if row_idx >= len(df):
                            continue

                        time_cell = df.iloc[row_idx, 0]
                        if pd.isna(time_cell):
                            continue

                        time_match = re.match(r"\d{1,2}[:：]\d{2}", str(time_cell))
                        if not time_match:
                            continue
                        time_str = time_match[0].replace("：", ":")

                        # Check all adjacent columns for multiple entries
                        for offset in range(2):
                            check_col = col + offset
                            if check_col >= df.shape[1]:
                                continue

                            subject_cell = df.iloc[row_idx, check_col]
                            if pd.isna(subject_cell):
                                continue

                            parts = str(subject_cell).split("/")
                            if len(parts) != 2:
                                continue

                            student_name, subject = parts[0].strip(), parts[1].strip()

                            raw_results.append({
                                "Teacher name": teach_name,
                                "Student name": student_name,
                                "Subject": subject,
                                "Date": class_date.isoformat(),
                                "Time": time_str
                            })

        # Now group the raw results
        grouped = {}
        for entry in raw_results:
            key = (entry["Teacher name"], entry["Date"], entry["Time"])
            if key not in grouped:
                grouped[key] = {
                    "Teacher name": entry["Teacher name"],
                    "Student name": [],
                    "Subject": [],
                    "Date": entry["Date"],
                    "Time": entry["Time"]
                }
            grouped[key]["Student name"].append(entry["Student name"])
            grouped[key]["Subject"].append(entry["Subject"])

        return list(grouped.values())

    def teach_main(self):
        all_schedules = self.extract_schedule_calendar_blocks(self.stu_file_path)
        output_path = "teacher_schedule.json"
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(all_schedules, f, ensure_ascii=False, indent=2)
        print(f"Schedule exported to {output_path} with {len(all_schedules)} entries.")
