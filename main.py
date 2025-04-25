from student_data import Student_data
from teacher_schedule import Teacher_data
from schedule_result import Schedule_result
import csv
import tkinter as tk
from tkinter.font import Font
from tkinter import filedialog
from tkinter import ttk
import os


class MainDisplay:
    BACKGROUND_COLOR = "#ffe4c4"

    def __init__(self):
        self.root = tk.Tk()
        self.set_layout()

    def __del__(self):
        pass

    def disp_input_student(self):
        self.stu_file_path = filedialog.askopenfilenames(
            initialdir=os.path.join(os.getcwd(), "input"),
            title="Select a file",
            filetypes=[("Excel Files", "*.xlsx*")],
        )
        if self.stu_file_path:
            print(f"Selected file: {self.stu_file_path}")
            Student_data(self.stu_file_path)

    def disp_input_teacher(self):
        self.teach_file_path = filedialog.askopenfilename(
            initialdir=os.path.join(os.getcwd(), "input"),
            title="Select a file",
            filetypes=[("Excel Files", "*.xlsx*")],
        )
        if self.teach_file_path:
            print(f"Selected file: {self.teach_file_path}")
            Teacher_data(self.teach_file_path)

    def schedule_result(self):
        # self.teach_file_path = filedialog.askopenfilename(
        #     initialdir=os.path.join(os.getcwd(), "input"),
        #     title="Select a file",
        #     filetypes=[("Excel Files", "*.xlsx*")],
        # )
        # if self.teach_file_path:
        #     print(f"Selected file: {self.teach_file_path}")
        sr = Schedule_result()
        sr.generate("E:/japanese_tak/formatted_schedule_output.xlsx")  

    def set_center_position(self, window_width: int = 700, window_height: int = 400):
        # sourcery skip: move-assign-in-block, use-fstring-for-concatenation

        ww = self.root.winfo_screenwidth()
        wh = self.root.winfo_screenheight()
        lh = window_height
        lw = window_width
        self.root.geometry(
            (f"{lw}x{lh}+{int(ww / 2 - lw / 2)}" + "+") + str(int(wh / 2 - lh / 2))
        )

    def disp_window_title(self):
        self.root.title("Schedule")

    def set_layout(self):
        self.set_center_position(window_width=680, window_height=130)
        self.root.resizable(width=False, height=False)
        self.root.configure(bg=self.BACKGROUND_COLOR)
        self.disp_window_title()

        font = Font(family="Meiryo UI", size=10)
        self.btn_input_files = tk.Button(
            text="Select Student_Schedule", command=lambda: self.disp_input_student()
        )
        self.btn_input_files.configure(font=font, bg="floralwhite", relief=tk.GROOVE)
        self.btn_input_files.place(x=50, y=50, width=200, height=30)

        font = Font(family="Meiryo UI", size=10)
        self.btn_input_files = tk.Button(
            text="Select Teacher_Schedule", command=lambda: self.disp_input_teacher()
        )
        self.btn_input_files.configure(font=font, bg="floralwhite", relief=tk.GROOVE)
        self.btn_input_files.place(x=300, y=50, width=200, height=30)

        self.btn_excute = tk.Button(text="Execute", command=lambda: self.schedule_result())
        self.btn_excute.configure(font=font)
        self.btn_excute.place(x=550, y=35, width=80, height=60)

      
        self.root.mainloop()

    def set_result(self): 

        self.root.update_idletasks()


def main():
    MainDisplay()  # Instantiate the class
    


if __name__ == "__main__":
    main()
