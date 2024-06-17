import tkinter as tk
from tkinter import ttk

class Page6View(ttk.Frame):
    ynCommnet = ''
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="8. 엑셀 다운로드 하시겠습니까?")
        lb1.place(relx=0.05, rely=0.05)

        self.button_yes = ttk.Button(self, text="예", command=self.on_yes)
        self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="아니오 (처음으로)", command=lambda: controller.show_frame("HomeView"))
        self.button_no.place(relx=0.24, rely=0.3)

    def update_data(self, decision):
        self.ynCommnet = decision
        
    def on_yes(self):
        # 코멘트 여부 확인
        print('엑셀 다운로드 함수 실행!')
        print('코멘트 : ' + self.ynCommnet)
        
  