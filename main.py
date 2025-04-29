from student_data import Student_data
from teacher_schedule import Teacher_data
from schedule_result import Schedule_result
import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
from tkinter import filedialog
import os
from enum import Enum

class MainDisplay:
    BACKGROUND_COLOR = "#f5f5f5"  # Lighter background for contrast

    def __init__(self):
        self.root = tk.Tk()
        self.loading_label = None
        self.style = ttk.Style()
        self.scrape_data = []
        self.set_layout()

    def __del__(self):
        pass

    def show_loading(self):
        if not self.loading_label:
            self.loading_label = ttk.Label(
                self.root, text="Loading...",
                font=("Meiryo UI", 10),
                foreground="#ff6600",
                background=self.BACKGROUND_COLOR
            )
            self.loading_label.place(relx=0.5, y=180, anchor="center")
        self.root.update_idletasks()

    def hide_loading(self):
        if self.loading_label:
            self.loading_label.destroy()
            self.loading_label = None
        self.root.update_idletasks()

    def disp_input_student(self):
        self.show_loading()
        self.stu_file_path = filedialog.askopenfilenames(
            initialdir=os.path.join(os.getcwd(), "input"),
            title="Select Student Schedule",
            filetypes=[("Excel Files", "*.xlsx*")],
        )
        if self.stu_file_path:
            print(f"Selected file: {self.stu_file_path}")
            Student_data(self.stu_file_path)
        self.hide_loading()

    def disp_input_teacher(self):
        self.show_loading()
        self.teach_file_path = filedialog.askopenfilename(
            initialdir=os.path.join(os.getcwd(), "input"),
            title="Select Teacher Schedule",
            filetypes=[("Excel Files", "*.xlsx*")],
        )
        if self.teach_file_path:
            print(f"Selected file: {self.teach_file_path}")
            Teacher_data(self.teach_file_path)
        self.hide_loading()

    def schedule_result(self, file_path):
        sr = Schedule_result(file_path=file_path)
        sr.generate()

    def disp_output_folder(self):
        self.show_loading()
        if file_path := filedialog.asksaveasfilename(
            initialdir=os.path.join(os.getcwd(), "output"),
            title="Save Schedule File",
            defaultextension=".csv",
            filetypes=[("Excel Files", "*.xlsx")],
        ):
            try:
                self.schedule_result(file_path)
            except Exception as e:
                print(e)
            self.root.update_idletasks()
        else:
            self.set_result("File save cancelled.")
            self.state = Enum("State", "EXPORT")
        self.hide_loading()

    def set_center_position(self, window_width: int = 750, window_height: int = 250):
        ww = self.root.winfo_screenwidth()
        wh = self.root.winfo_screenheight()
        lw, lh = window_width, window_height
        self.root.geometry(
            f"{lw}x{lh}+{int(ww/2 - lw/2)}+{int(wh/2 - lh/2)}"
        )

    def disp_window_title(self):
        self.root.title("スケジュール管理ツール")

    def configure_style(self):
        self.style.theme_use('default')
        self.style.configure(
            "RoundedButton.TButton",
            font=("Meiryo UI", 11),
            foreground="#ffffff",  # Button text color
            background="#0078D7",  # Deep blue
            borderwidth=0,
            padding=10,
            relief="flat"
        )
        self.style.map(
            "RoundedButton.TButton",     
            background=[('active', '#3399ff')],
            foreground=[('active', '#ffffff')],
        )

    def set_layout(self):
        self.set_center_position()
        self.root.resizable(width=False, height=False)
        self.root.configure(bg=self.BACKGROUND_COLOR)
        self.disp_window_title()
        self.configure_style()

        font = Font(family="Meiryo UI", size=11)

        self.btn_input_student = ttk.Button(
            self.root, text="生徒スケジュール選択",
            command=self.disp_input_student, style="RoundedButton.TButton"
        )
        self.btn_input_student.place(x=80, y=50, width=220, height=50)

        self.btn_input_teacher = ttk.Button(
            self.root, text="講師スケジュール選択",
            command=self.disp_input_teacher, style="RoundedButton.TButton"
        )
        self.btn_input_teacher.place(x=80, y=120, width=220, height=50)

        self.btn_execute = ttk.Button(
            self.root, text="スケジュール作成実行",
            command=self.disp_output_folder, style="RoundedButton.TButton"
        )
        self.btn_execute.place(x=400, y=85, width=250, height=60)

        self.root.mainloop()

    def set_result(self, msg=""):
        print(msg)
        self.root.update_idletasks()

def main():
    MainDisplay()

if __name__ == "__main__":
    main()
