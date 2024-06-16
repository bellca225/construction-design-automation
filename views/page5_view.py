import tkinter as tk
from tkinter import ttk

class Page5View(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="엑셀 내 코멘트를 남겨주세요.")
        lb1.place(relx=0.05, rely=0.05)
        
        entry = ttk.Entry(self)
        entry.place(relx=0.05, rely=0.1, width=200, height=80)
        
        self.button_yes = ttk.Button(self, text="저장", command=self.on_yes)
        self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="취소", command=self.on_no)
        self.button_no.place(relx=0.25, rely=0.3)
    

    def on_yes(self):
        # self.controller.selection = "예"
        self.controller.set_comment_decision("예")
        self.controller.show_frame("Page6View")

    def on_no(self):
        # self.controller.selection = "아니오"
        self.controller.set_comment_decision("아니오")
        self.controller.show_frame("Page6View")