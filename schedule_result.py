import json
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
from datetime import datetime
from collections import defaultdict
import os
import traceback
traceback.print_exc()

class Schedule_result:
    def __init__(self, file_path):
        self.file_path = file_path
        with open('student_schedules.json', 'r', encoding='utf-8') as f:
            self.student_data = json.load(f)
        with open('teacher_schedule.json', 'r', encoding='utf-8') as f:
            self.teacher_data = json.load(f)
        
        self.times = ["13:10", "14:40", "16:30", "18:00", "19:30"]
        self.booths = list(range(1, 13))
        self.dates = sorted({s['Date'] for s in self.student_data})
        self.matched = defaultdict(lambda: {'teacher': None, 'students': []})
        
        self.match_students_and_teachers()
        self.month_dates = defaultdict(list)
        for date in self.dates:
            dt = datetime.strptime(date, "%Y-%m-%d")
            month_key = dt.strftime("%Y-%m")
            self.month_dates[month_key].append(date)

    def normalize_time(self, time_str):
        return time_str.split('～')[0].strip() if '～' in time_str else time_str.strip()

    def match_students_and_teachers(self):
        student_groups = defaultdict(list)
        teacher_groups = defaultdict(list)

        for s in self.student_data:
            key = (s['Date'], self.normalize_time(s['Time']))
            student_groups[key].append(s)

        for t in self.teacher_data:
            key = (t['Date'], self.normalize_time(t['Time']))
            teacher_groups[key].append(t)

        for key, students in student_groups.items():
            teachers = teacher_groups.get(key, [])
            unmatched_students = students.copy()
            booth_number = 1

            for teacher in teachers:
                matched_students = []
                teacher_students = teacher.get('Student name', [])
                teacher_subjects = teacher.get('Subject', [])
                if not isinstance(teacher_students, list):
                    teacher_students = teacher_students.split('、')
                if not isinstance(teacher_subjects, list):
                    teacher_subjects = [teacher_subjects]

                for student_name, teacher_subject in zip(teacher_students, teacher_subjects):
                    for student in unmatched_students:
                        if student['Student name'] != student_name:
                            continue

                        subject_match = False
                        if student['Subject'] == '理社' and teacher_subject == '理社通':
                            subject_match = True
                        elif student['Subject'] == '文系' and teacher_subject in ('文系通', '文系講'):
                            subject_match = True
                        elif student['Subject'][0] == teacher_subject[0]:
                            subject_match = True

                        form_match = False
                        student_form = student.get('Subject form', '')
                        if teacher_subject and student_form:
                            form_match = teacher_subject[-1] == student_form[0]

                        if subject_match and form_match:
                            matched_students.append(student)
                            unmatched_students.remove(student)
                            break  # Found, so break inner loop

                if matched_students:
                    self.matched[(key, teacher['Teacher name'])] = {
                        'teacher': teacher,
                        'students': matched_students,
                        'booth': booth_number
                    }
                    booth_number += 1

            # Secondary matching for remaining students
            for teacher in teachers:
                if any((key, teacher['Teacher name']) == k for k in self.matched.keys()):
                    continue  # already matched in first pass

                matched_students = []
                teacher_students = teacher.get('Student name', [])
                teacher_subjects = teacher.get('Subject', [])
                if not isinstance(teacher_students, list):
                    teacher_students = teacher_students.split('、')
                if not isinstance(teacher_subjects, list):
                    teacher_subjects = [teacher_subjects]

                for student in unmatched_students[:]:
                    for teacher_subject in teacher_subjects:
                        subject_match = False
                        if student['Subject'] == '理社' and teacher_subject == '理社通':
                            subject_match = True
                        elif student['Subject'] == '文系' and teacher_subject in ('文系通', '文系講'):
                            subject_match = True
                        elif student['Subject'][0] == teacher_subject[0]:
                            subject_match = True

                        form_match = False
                        student_form = student.get('Subject form', '')
                        if teacher_subject and student_form:
                            form_match = teacher_subject[-1] == student_form[0]

                        if subject_match and form_match:
                            matched_students.append(student)
                            unmatched_students.remove(student)
                            break  # Break inner loop after matching

                if matched_students:
                    self.matched[(key, teacher['Teacher name'])] = {
                        'teacher': teacher,
                        'students': matched_students,
                        'booth': booth_number
                    }
                    booth_number += 1

    def generate(self):
        try:
            red_border = Border(right=Side(style='medium', color='FF0000'))
            wb = Workbook()
            wb.remove(wb.active)
            align = Alignment(wrap_text=True, vertical="center", horizontal="center")
            gray_fill = PatternFill(fill_type="solid", fgColor="DDDDDD")
            green_font = Font(color="008000")
            blue_font = Font(color="0000FF")
            # New additional fonts and settings
            header_font = Font(color="000080", size=12, bold=True)  # Dark navy, bigger, bold
            header_height = 30  # Height for header rows
            time_booth_font = Font(color="000080", size=12, bold=True)

            for month, dates_in_month in self.month_dates.items():
                ws = wb.create_sheet(title=month.replace("-", "_"))
                ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1)
                ws.merge_cells(start_row=1, start_column=2, end_row=3, end_column=2)
                ws.cell(row=1, column=1, value="時間帯").alignment = align
                ws.cell(row=1, column=1).font = time_booth_font
                ws.cell(row=1, column=2, value="ブース番号").alignment = align
                ws.cell(row=1, column=2).font = time_booth_font
                col = 3
                for date in dates_in_month:
                    dt = datetime.strptime(date, "%Y-%m-%d")
                    formatted = f"{dt.day} 日\n({dt.strftime('%a')})"
                    ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col+6)
                    ws.cell(row=1, column=col, value=formatted).alignment = align

                    ws.merge_cells(start_row=2, start_column=col, end_row=3, end_column=col)
                    ws.cell(row=2, column=col, value="担当講師").alignment = align
                    ws.cell(row=2, column=col).font = blue_font

                    ws.merge_cells(start_row=2, start_column=col+1, end_row=2, end_column=col+3)
                    ws.cell(row=2, column=col+1, value="生徒1").alignment = align
                    ws.cell(row=2, column=col+1).font = blue_font

                    ws.cell(row=3, column=col+1, value="名前").alignment = align
                    ws.cell(row=3, column=col+2, value="学年 科目").alignment = align
                    ws.cell(row=3, column=col+3, value="通/講").alignment = align

                    ws.merge_cells(start_row=2, start_column=col+4, end_row=2, end_column=col+6)
                    ws.cell(row=2, column=col+4, value="生徒2").alignment = align
                    ws.cell(row=2, column=col+4).font = blue_font

                    ws.cell(row=3, column=col+4, value="名前").alignment = align
                    ws.cell(row=3, column=col+5, value="学年 科目").alignment = align
                    ws.cell(row=3, column=col+6, value="通/講").alignment = align

                    # Apply header font and header height
                    for r in range(1, 4):
                        for c in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c).font = header_font
                        ws.row_dimensions[r].height = header_height

                    for row in range(1, 4):
                        date_separator_cell = ws.cell(row=row, column=col+6)
                        date_separator_cell.border = red_border
                    col += 7

                current_row = 4
                for time in self.times:
                    time_start_row = current_row
                    for booth in self.booths:
                        cell_time = ws.cell(row=current_row, column=1, value=time)
                        cell_time.alignment = align
                        cell_time.font = time_booth_font  # <<< add this

                        cell_booth = ws.cell(row=current_row, column=2, value=booth)
                        cell_booth.alignment = align
                        cell_booth.font = time_booth_font  # <<< add this

                        col = 3
                        for date in dates_in_month:
                            key = (date, time)
                            booth_info = [v for k, v in self.matched.items() if k[0] == key and v['booth'] == booth]
                            cell_teacher = ws.cell(row=current_row, column=col)
                            cell_teacher.alignment = align
                            if booth_info:
                                for info in booth_info:
                                    if teacher := info['teacher']:
                                        students = info['students']

                                        if teacher:
                                            cell_teacher.value = teacher['Teacher name']

                                        if students:
                                            ws.cell(row=current_row, column=col+1, value=students[0].get('Student name', '')).alignment = align
                                            ws.cell(row=current_row, column=col+2, value=students[0].get('Grade', '') + " " + students[0].get('Subject', '')).alignment = align
                                            ws.cell(row=current_row, column=col+3, value=students[0].get('Subject form', '')).alignment = align

                                        if len(students) > 1:
                                            ws.cell(row=current_row, column=col+4, value=students[1].get('Student name', '')).alignment = align
                                            ws.cell(row=current_row, column=col+5, value=students[1].get('Grade', '') + " " + students[1].get('Subject', '')).alignment = align
                                            ws.cell(row=current_row, column=col+6, value=students[1].get('Subject form', '')).alignment = align

                            for i in range(col, col+7):
                                cell = ws.cell(row=current_row, column=i)
                                cell.fill = gray_fill
                                cell.border = Border(
                                    left=Side(style='thin', color='000000'),
                                    right=Side(style='thin', color='000000'),
                                    top=Side(style='thin', color='000000'),
                                    bottom=Side(style='thin', color='000000')
                                )
                                if (i - 3) % 7 == 6:
                                    cell.border = Border(
                                        left=Side(style='thin', color='000000'),
                                        right=Side(style='medium', color='FF0000'),
                                        top=Side(style='thin', color='000000'),
                                        bottom=Side(style='thin', color='000000')
                                    )
                                elif (i - 3) % 7 == 3:
                                    cell.border = Border(
                                        left=Side(style='thin', color='000000'),
                                        right=Side(style='medium', color='000000'),
                                        top=Side(style='thin', color='000000'),
                                        bottom=Side(style='thin', color='000000')
                                    )
                            col += 7
                        current_row += 1

                    for c in range(1, ws.max_column + 1):
                        cell = ws.cell(row=time_start_row, column=c)
                        existing_border = cell.border
                        cell.border = Border(
                            top=Side(style='medium', color='0000FF'),
                            left=existing_border.left,
                            right=existing_border.right,
                            bottom=existing_border.bottom
                        )
                    ws.merge_cells(start_row=time_start_row, start_column=1, end_row=current_row-1, end_column=1)

                for col_idx in range(1, ws.max_column + 1):
                    col_letter = get_column_letter(col_idx)
                    ws.column_dimensions[col_letter].width = 21
                    ws.column_dimensions['A'].width = 20  # 時間帯
                    ws.column_dimensions['B'].width = 10  # ボース番号

                for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
                    ws.row_dimensions[row[0].row].height = 30

            wb.save(self.file_path)
            print(f"Excel saved to: {os.path.abspath(self.file_path)}")
        except Exception as e:
            print("Error occurred during Excel generation:")
            traceback.print_exc()