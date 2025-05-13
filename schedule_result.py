import json
import os
from openpyxl import load_workbook, Workbook
import unicodedata
import re
from copy import copy
from openpyxl.cell.cell import Cell
from datetime import datetime, timedelta
from openpyxl import Workbook as OpenpyxlWorkbook
from openpyxl.utils import get_column_letter

def _normalize_name(name: str) -> str:
    return unicodedata.normalize("NFKC", name.replace("　", " ")).strip()
TIME_ROW_MAP = {
    "13:10": 5,
    "14:40": 6,
    "16:30": 7,
    "18:00": 8,
    "19:30": 9,
}

def copy_worksheet_template(target_wb, template_ws, new_title):
    new_ws = target_wb.create_sheet(title=new_title)
    merged_ranges = set()
    for merged_range in template_ws.merged_cells.ranges:
        merged_ranges.update(merged_range.cells)
    for row in template_ws.iter_rows():
        for cell in row:
            if not isinstance(cell, Cell):
                continue
            is_merged = (cell.coordinate in merged_ranges)
            new_cell = new_ws.cell(row=cell.row, column=cell.column)
            if not is_merged:
                new_cell.value = cell.value
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    for merged_range in template_ws.merged_cells.ranges:
        new_ws.merge_cells(str(merged_range))
    return new_ws


def get_top_left_if_merged(ws, row, col):
    cell_coord = f"{get_column_letter(col)}{row}"
    for merged_range in ws.merged_cells.ranges:
        if cell_coord in merged_range:
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return ws.cell(row=row, column=col)


class Schedule_result:
    def __init__(self, student_data, teacher_data, match_data, student_template, teacher_template, date_list):
        self.student_data = {
            _normalize_name(k): v for k, v in student_data.items()
        }
        self.teacher_data = {
            _normalize_name(k): v for k, v in teacher_data.items()
        }
        self.subject_data = match_data
        self.student_template = student_template
        self.teacher_template = teacher_template
        self.date_list= date_list
        self.teacher_output_path = "output/teachers_schedule.xlsx"
        self.student_output_dirs = {
            'elementary': "output/students_elementary.xlsx",
            'middle': "output/students_middle.xlsx",
            'high': "output/students_high.xlsx",
        }
        self.schedule_data = []
        self.date_order = []
        self.used_teacher_slots = {}  # (teacher, date, time, booth_index) -> True
    def run(self):
        self.generate_schedule()
        self.generate_teacher_excel()
        self.generate_student_excels()
    @staticmethod
    def parse_date(date_str):
        return datetime.strptime(date_str, "%Y-%m-%d")
    def generate_schedule(self):
        matched_pairs = set()
        teacher_daily_count = {}
        student_daily_count = {}
        last_scheduled = {}  # (student, teacher, subject) -> [datetime objects]
        for entry in self.subject_data:
            student_name = _normalize_name(entry.get('student_name', ''))
            grade = entry.get('grade', 'other')
            
            if student_name not in self.student_data:
                continue
            student_avail = self.student_data[student_name].get('schedule', {})
            # print(student_name)
            # print(student_avail)
            for subj in entry.get('subjects', []):
                # print(subj)
                if not isinstance(subj, dict):
                    continue
                subject = subj.get('name')
                raw_teacher = subj.get('teacher', '')
                # print(raw_teacher)
                teacher = _normalize_name(raw_teacher.get('name') if isinstance(raw_teacher, dict) else raw_teacher)
                if teacher not in self.teacher_data:
                    continue
                teacher_avail = self.teacher_data[teacher].get('schedule', {})
                subject_type = '特別' if subj.get('special_classes', 0) > 0 else '通常'
                pair_key = (student_name, teacher, subject)
                last_scheduled.setdefault(pair_key, [])
                total_classes = subj.get('regular_classes', 0) + subj.get('special_classes', 0)
                # print(total_classes)
                if total_classes == 0:
                    continue
                available_dates = sorted(set(student_avail.keys()) & set(teacher_avail.keys()))
              
                for date_str in available_dates:
                    date_obj = self.parse_date(date_str)
                    # Enforce spacing between sessions

                    if any(abs((date_obj - prev).days) < 2 for prev in last_scheduled[pair_key]):
                        continue
                    
                    if total_classes > 12 and any(abs((date_obj - prev).days) < 4 for prev in last_scheduled[pair_key]):
                        continue
                    teacher_daily = teacher_daily_count.setdefault(teacher, {})
                    student_daily = student_daily_count.setdefault(student_name, {})
                    if teacher_daily.get(date_str, 0) >= 2 or student_daily.get(date_str, 0) >= 2:
                        continue
                    for time_slot in TIME_ROW_MAP:
                        if not (student_avail[date_str].get(time_slot) and teacher_avail[date_str].get(time_slot)):
                            continue
                        t_free_list = teacher_avail[date_str][time_slot]
                        for booth_index in range(len(t_free_list)):
                            key = (teacher, date_str, time_slot, booth_index)
                            if t_free_list[booth_index] and not self.used_teacher_slots.get(key):

                                self.schedule_data.append({
                                    'date': date_str,
                                    'time': time_slot,
                                    'student': student_name,
                                    'teacher': teacher,
                                    'subject': subject,
                                    'type': subject_type,
                                    'grade': grade
                                })
                                # Mark slot used
                                student_avail[date_str][time_slot] = False
                                teacher_avail[date_str][time_slot][booth_index] = False
                                self.used_teacher_slots[key] = True
                                # Update counts
                                teacher_daily[date_str] = teacher_daily.get(date_str, 0) + 1
                                student_daily[date_str] = student_daily.get(date_str, 0) + 1
                                last_scheduled[pair_key].append(date_obj)
                                break
                        else:
                            continue
                        break
                    if len(last_scheduled[pair_key]) >= total_classes:
                        break
        self.date_order = sorted({entry['date'] for entry in self.schedule_data})
    def generate_teacher_excel(self):
        if not self.schedule_data:
            print("⚠️ No teacher data found — no file saved.")
            return
        if isinstance(self.teacher_template, str):
            self.teacher_template = load_workbook(self.teacher_template)
        template_wb = self.teacher_template
        template_sheetnames = template_wb.sheetnames
        name_map = {}

        for entry in self.schedule_data:
            full_name = entry['teacher']
            match = None
            for sheet_name in template_sheetnames:
                normalized_sheet = _normalize_name(sheet_name)
                if full_name == normalized_sheet or normalized_sheet in full_name:
                    match = sheet_name
                    break

            if match:
                name_map[full_name] = match
            else:
                print(f"⚠️ No matching sheet found for teacher: {full_name}")
        output_wb = Workbook()
        output_wb.remove(output_wb.active)
        handled_teachers = set()
        for entry in self.schedule_data:
            teacher = entry['teacher']
            if teacher not in name_map:
                continue
            template_sheetname = name_map[teacher]
            if teacher not in handled_teachers:
                template_ws = self.teacher_template[template_sheetname]
                new_ws = copy_worksheet_template(output_wb, template_ws, teacher[:30])
                handled_teachers.add(teacher)
            else:
                # Always access the sheet by name, not by previously cached variable
                new_ws = output_wb[teacher[:30]]

            row = int(TIME_ROW_MAP.get(entry['time'])) + 8*self._date_to_row(entry['date']) 
            col = self._date_to_col(entry['date'])
            if col < 1 or row < 1:
                continue

            entry_text = f"{entry['student']}:{entry['subject']}"
            cell_primary = get_top_left_if_merged(new_ws, row, col)
            cell_secondary = get_top_left_if_merged(new_ws, row, col + 1)

            if not cell_primary.value:
                cell_primary.value = entry_text
            elif not cell_secondary.value:
                cell_secondary.value = entry_text
        os.makedirs(os.path.dirname(self.teacher_output_path), exist_ok=True)
        output_wb.save(self.teacher_output_path)
        
    def generate_student_excels(self):
        def normalize_grade(raw_grade):
            raw = raw_grade.lower()
            if any(k in raw for k in ['小', 'elementary', '小学']):
                return 'elementary'
            elif any(k in raw for k in ['中', 'middle', '中学']):
                return 'middle'
            elif any(k in raw for k in ['高', 'high', '高校']):
                return 'high'
            return None
        grouped_students = {'elementary': [], 'middle': [], 'high': []}
        for entry in self.schedule_data:
            grade_raw = entry.get('grade', '')
            grade = normalize_grade(grade_raw)
            if grade in grouped_students:
                grouped_students[grade].append(entry)

        for category, entries in grouped_students.items():
            if not entries:
                continue

            wb = Workbook()
            wb.remove(wb.active)
            students = sorted({
                _normalize_name(entry['student']) 
                for entry in entries 
                if entry.get('student')
            })
            
            for student in students:
                student_entries = [e for e in entries if _normalize_name(e['student']) == student]
         
                matched = False
                for student_templ in self.student_template:
                    if isinstance(student_templ, str):
                        template_wb = load_workbook(student_templ)
                    elif isinstance(student_templ, OpenpyxlWorkbook):
                        template_wb = student_templ
                    else:
                        raise TypeError("student_templ must be a path or openpyxl Workbook")

                    for sheet_name in template_wb.sheetnames:
                        student_norm = _normalize_name(student)
                        sheet_norm = _normalize_name(sheet_name)
                        print(student_norm, "==", sheet_norm)
                        # Direct full or partial match
                        if (
                            student_norm == sheet_norm or
                            student_norm in sheet_norm or
                            sheet_norm in student_norm or
                            any(part in sheet_norm for part in student_norm.split())
                        ):
    
                            template_ws = template_wb[sheet_name]
                            new_sheet_name = f"{sheet_name[:9]}"
                            new_ws = copy_worksheet_template(wb, template_ws, new_sheet_name)

                            for entry in student_entries:
                                row = int(TIME_ROW_MAP.get(entry['time'])) + 8 * self._date_to_row(entry['date']) 
                                col = self._date_to_col(entry['date'])
                                if row and col:
                                    new_ws.cell(row=row, column=col-1).value = entry['subject']
                                    new_ws.cell(row=row, column=col).value = entry['type']
                                else:
                                    print(f"⚠️ Skipping invalid date/time in entry: {entry}")

                            matched = True
                            break  # Stop after first matching sheet

                if not matched:
                    print(f"⚠️ No matching sheet found for student: {student}")

            # Ensure workbook has at least one visible sheet
            if not wb.sheetnames:
                fallback_ws = wb.create_sheet("NoMatches")
                fallback_ws["A1"] = "⚠️ No matching sheets found for any student in this group."

            output_path = self.student_output_dirs[category]
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            wb.save(output_path)
            print(f"💾 Saved: {output_path}")

    def _date_to_col(self, date_str):
        base_col = 2
        try:
            index = self.date_list.index(date_str) %5
        except ValueError:
            index = 0
        return base_col + index * 2 + 1

    def _date_to_row(self, date_str):
        try:
            index = int(self.date_list.index(date_str) /5)
        except ValueError:
            index = 0
        return index
