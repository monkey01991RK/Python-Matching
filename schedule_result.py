import json
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
from datetime import datetime
from collections import defaultdict
import os

class Schedule_result:
    def __init__(self):
        with open('student_schedules.json', 'r', encoding='utf-8') as f:
            self.student_data = json.load(f)

        with open('teacher_schedule.json', 'r', encoding='utf-8') as f:
            self.teacher_data = json.load(f)

        self.times = ["13:10", "14:40", "16:30", "18:00", "19:30"]
        self.booths = list(range(1, 13))
        self.dates = sorted(set(s['Date'] for s in self.student_data))
        self.teacher_map = {}
        self.build_teacher_map()
        self.assign_booths()

    def normalize_time(self, time_str):
        return time_str.split('～')[0].strip() if '～' in time_str else time_str.strip()

    def build_teacher_map(self):
        for entry in self.teacher_data:
            student_name = entry['Student name']
            subject = entry.get('Subject', '').strip()
            first_char = subject[0] if subject else ''
            date = entry['Date']
            time = self.normalize_time(entry['Time'])
            key = (student_name, date, time, first_char)
            self.teacher_map[key] = entry['Teacher name']

    def assign_booths(self):
        assigned = defaultdict(set)
        for s in self.student_data:
            date = s['Date']
            time = self.normalize_time(s['Time'])
            key = (date, time)
            if 'Booth' not in s or s['Booth'] is None:
                for booth in self.booths:
                    if booth not in assigned[key]:
                        s['Booth'] = booth
                        assigned[key].add(booth)
                        break
            else:
                assigned[key].add(s['Booth'])

    def generate(self, output_path="formatted_schedule_output.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Schedule"

        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
        align = Alignment(wrap_text=True, vertical="center", horizontal="center")
        gray_fill = PatternFill("solid", fgColor="DDDDDD")

        ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1)
        ws.merge_cells(start_row=1, start_column=2, end_row=3, end_column=2)
        ws.cell(row=1, column=1, value="時間帯").alignment = align
        ws.cell(row=1, column=2, value="ブース番号").alignment = align

        col = 3
        for date in self.dates:
            dt = datetime.strptime(date, "%Y-%m-%d")
            formatted = f"{dt.day} 日\n({dt.strftime('%a')})"
            ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col+6)
            ws.cell(row=1, column=col, value=formatted).alignment = align

            ws.merge_cells(start_row=2, start_column=col, end_row=3, end_column=col)
            ws.cell(row=2, column=col, value="担当講師").alignment = align

            ws.merge_cells(start_row=2, start_column=col+1, end_row=2, end_column=col+3)
            ws.cell(row=2, column=col+1, value="生徒1").alignment = align
            ws.cell(row=3, column=col+1, value="名前").alignment = align
            ws.cell(row=3, column=col+2, value="学年 科目").alignment = align
            ws.cell(row=3, column=col+3, value="通/講").alignment = align

            
            ws.merge_cells(start_row=2, start_column=col+4, end_row=2, end_column=col+6)
            ws.cell(row=2, column=col+4, value="生徒2").alignment = align
            ws.cell(row=3, column=col+4, value="名前").alignment = align
            ws.cell(row=3, column=col+5, value="学年 科目").alignment = align
            ws.cell(row=3, column=col+6, value="通/講").alignment = align

            col += 7

        current_row = 4
        for time in self.times:
            time_start_row = current_row
            for booth in self.booths:
                ws.cell(row=current_row, column=2, value=booth).alignment = align
                col = 3
                for date in self.dates:
                    booth_data = [s for s in self.student_data
                                  if s.get('Date') == date and self.normalize_time(s.get('Time', '')) == time and s.get('Booth') == booth]
                    students = []
                    teacher = ''
                    for s in booth_data:
                        sub = s.get('Subject', '').strip()
                        first_char = sub[0] if sub else ''
                        t_key = (s['Student name'], s['Date'], self.normalize_time(s['Time']), first_char)
                        teacher = self.teacher_map.get(t_key, '')
                        students.append(s)

                    ws.cell(row=current_row, column=col, value=teacher).alignment = align
                    if len(students) > 0:
                        
                        ws.cell(row=current_row, column=col+1, value=students[0].get('Student name', '')).alignment = align
                        ws.cell(row=current_row, column=col+2, value=students[0].get('Grade', '') + " " + students[0].get('Subject', '')).alignment = align
                        ws.cell(row=current_row, column=col+3, value=students[0].get('Type', '')).alignment = align
                    if len(students) > 1:
                      
                        ws.cell(row=current_row, column=col+4, value=students[1].get('Student name', '')).alignment = align
                        ws.cell(row=current_row, column=col+5, value=students[1].get('Grade', '') + " " + students[1].get('Subject', '')).alignment = align
                        ws.cell(row=current_row, column=col+6, value=students[1].get('Type', '')).alignment = align
                    for i in range(col, col+7):
                        cell = ws.cell(row=current_row, column=i)
                        cell.border = border
                        cell.fill = gray_fill
                    col += 7
                current_row += 1
            ws.merge_cells(start_row=time_start_row, start_column=1, end_row=current_row - 1, end_column=1)
            ws.cell(row=time_start_row, column=1, value=time).alignment = align

        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            if col_letter == 'B':
                ws.column_dimensions[col_letter].width = 6
            elif col_letter in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws.column_dimensions[col_letter].width = 12
            else:
                ws.column_dimensions[col_letter].width = 10

        for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
            ws.row_dimensions[row[0].row].height = 16

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        wb.save(output_path)
        print(f"✅ Excel saved to: {os.path.abspath(output_path)}")