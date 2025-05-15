import os
import json
import threading
import traceback
from tkinter import filedialog, messagebox
import customtkinter as ctk  # modern Tk replacement

from student_data import Student_data
from teacher_schedule import Teacher_data
from match_module import match_basic
from schedule_result import Schedule_result


class MainDisplay:
    """GUI for the scheduling tool, styled with customtkinter.

    – Modern muted palette inspired by current‑gen Japanese web apps
    – Indeterminate progress bar runs **smoothly** while long tasks run
      thanks to background threading
    – All user‑facing errors surface in message boxes (no hidden prints)
    """

    # ------------------------------------------------------------------ palette
    BG             = "#f3f6f9"
    CARD_BG        = "#ffffff"
    PRIMARY        = "#0066ff"
    PRIMARY_HOVER  = "#0050cc"
    CTA            = "#28a745"      # call‑to‑action (green)
    CTA_HOVER      = "#218838"

    def __init__(self) -> None:
        # ------------------------------ root window
        self.root = ctk.CTk()
        self.root.title("スケジュール管理ツール")
        self.root.configure(fg_color=self.BG)
        self._set_window_size(860, 290)
        self.root.resizable(False, False)
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # ------------------------------ fonts
        self.jp_font = ctk.CTkFont(family="Yu Gothic UI", size=12)
        self.h1_font = ctk.CTkFont(family="Yu Gothic UI", size=20, weight="bold")

        # ------------------------------ paths & state
        self.student_json_path = "student_schedules.json"
        self.teacher_json_path = "teacher_diagonal_schedule.json"
        self.match_json_path   = "all_students_schedule.json"
        self.stu_file_path: str | None = None
        self.teach_file_path: str | None = None

        # ------------------------------ build UI
        self._build_layout()
        self.root.mainloop()

    # ================================================================== layout
    def _build_layout(self) -> None:
        # banner ------------------------------------------------------
        ctk.CTkLabel(self.root, text="スケジュール管理ツール", font=self.h1_font).pack(pady=(24, 12))

        # card container ---------------------------------------------
        card = ctk.CTkFrame(self.root, fg_color=self.CARD_BG, corner_radius=16)
        card.pack(padx=32, pady=8, fill="both")

        btn_cfg: dict = {
            "master": card,
            "font": self.jp_font,
            "fg_color": self.PRIMARY,
            "hover_color": self.PRIMARY_HOVER,
            "text_color": "#ffffff",
            "corner_radius": 12,
            "height": 48,
        }

        # buttons ----------------------------------------------------
        self.btn_student = ctk.CTkButton(**btn_cfg, text="生徒スケジュール取込", command=self._on_student_click)
        self.btn_teacher = ctk.CTkButton(**btn_cfg, text="講師スケジュール取込", command=self._on_teacher_click)
        self.btn_match   = ctk.CTkButton(**btn_cfg, text="科目マッチデータ取込", command=self._on_match_click)

        self.btn_exec = ctk.CTkButton(
            master=card,
            text="🟢 スケジュール作成 ▶",
            font=ctk.CTkFont(family="Yu Gothic UI", size=16, weight="bold"),
            fg_color=self.CTA,
            hover_color=self.CTA_HOVER,
            text_color="white",
            width=540,
            height=50,
            corner_radius=10,
            command=self._on_execute_click,
        )

        # grid placement
        card.columnconfigure((0, 1), weight=1, uniform="col")
        card.rowconfigure((0, 1), weight=1)
        self.btn_student.grid(row=0, column=0, padx=16, pady=12, sticky="ew")
        self.btn_teacher.grid(row=0, column=1, padx=16, pady=12, sticky="ew")
        self.btn_match.grid(row=1, column=0, padx=16, pady=12, sticky="ew")
        self.btn_exec.grid(row=1, column=1, padx=16, pady=12, sticky="ew")

        # progress bar (indeterminate) --------------------------------
        self.progress = ctk.CTkProgressBar(self.root, width=600, height=8, corner_radius=4, mode="indeterminate")
        self.progress.pack(pady=(12, 0))
        self.progress.pack_forget()

    # ------------------------------------------------ progress helpers
    def _show_progress(self, msg: str | None = None):
        if msg:
            print(msg)
        self.progress.set(0)
        self.progress.pack(pady=(12, 0))
        self.progress.start()
        self.root.update_idletasks()

    def _hide_progress(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.root.update_idletasks()

    # ================================================= button callbacks
    def _on_student_click(self):
        self._show_progress("生徒スケジュール取込中 …")
        paths = filedialog.askopenfilenames(initialdir=os.path.join(os.getcwd(), "input"),
                                            title="生徒スケジュールを選択", filetypes=[("Excel Files", "*.xlsx")])
        if paths:
            self.stu_file_path = paths
            Student_data(paths)
        self._hide_progress()

    def _on_teacher_click(self):
        self._show_progress("講師スケジュール取込中 …")
        path = filedialog.askopenfilename(initialdir=os.path.join(os.getcwd(), "input"),
                                          title="講師スケジュールを選択", filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.teach_file_path = path
            Teacher_data(path)
        self._hide_progress()

    def _on_match_click(self):
        self._show_progress("マッチデータ読込中 …")
        path = filedialog.askopenfilename(initialdir=os.path.join(os.getcwd(), "input"),
                                          title="科目マッチ用Excelを選択", filetypes=[("Excel Files", "*.xlsx")])
        if path:
            match_basic(path)
        self._hide_progress()

    # ------------------------------------------------ long‑running task
    def _on_execute_click(self):
        # UI: show progress and spawn worker thread
        self._show_progress("スケジュール生成中 …")
        worker = threading.Thread(target=self._execute_schedule_task, daemon=True)
        worker.start()

    def _execute_schedule_task(self):
        """Runs in a background thread; schedules Excel generation."""
        try:
            # prerequisite JSONs
            if not all(map(os.path.exists, (self.student_json_path, self.teacher_json_path, self.match_json_path))):
                raise FileNotFoundError("必要なJSONファイルが見つかりません。まず各取込を行ってください。")

            with open(self.student_json_path, "r", encoding="utf-8") as f:
                student_data = json.load(f)
            with open(self.teacher_json_path, "r", encoding="utf-8") as f:
                teacher_data = json.load(f)
            with open(self.match_json_path, "r", encoding="utf-8") as f:
                match_data = json.load(f)

            lecture_dates = []
            if os.path.exists("lecture_dates.json"):
                with open("lecture_dates.json", "r", encoding="utf-8") as f:
                    lecture_dates = json.load(f)

            sr = Schedule_result(
                student_data=student_data,
                teacher_data=teacher_data,
                match_data=match_data,
                student_template=self.stu_file_path or "./templete/student_templete.xlsx",
                teacher_template=self.teach_file_path or "./templete/teacher_templete.xlsx",
                date_list=lecture_dates,
            )
            sr.run()

            # success message via main thread
            self.root.after(0, lambda: messagebox.showinfo("完了", "✅ スケジュール作成が成功しました！出力ファイルが保存されました。"))
        except Exception as e:
            err = f"❌ エラーが発生しました:\n{str(e)}\n\n詳細:\n{traceback.format_exc()}"
            self.root.after(0, lambda: messagebox.showerror("エラー", err))
        finally:
            # stop progress bar in main thread
            self.root.after(0, self._hide_progress)

    # ================================================================= helpers
    def _set_window_size(self, w: int, h: int):
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = int(sw / 2 - w / 2), int(sh / 2 - h / 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")


# ------------------------------------------------------------------ run GUI
if __name__ == "__main__":
    MainDisplay()
