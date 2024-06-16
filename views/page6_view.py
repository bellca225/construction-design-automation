import tkinter as tk
from tkinter import ttk



class Page6View(ttk.Frame):
    ynCommnet = ''
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="엑셀 다운로드 하시겠습니까?")
        lb1.place(relx=0.05, rely=0.05)
        lb1 = ttk.Label(self, text="공사금액 표시")
        lb1.place(relx=0.05, rely=0.1)

        self.button_yes = ttk.Button(self, text="예", command=self.on_yes)
        self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="아니오 (처음으로)", command=lambda: controller.show_frame("HomeView"))
        self.button_no.place(relx=0.25, rely=0.3)
        
        # self.decision_label = ttk.Label(self, text="코멘트 남기기 여부: ")
        # self.decision_label.pack()

    def update_data(self, decision):
        self.ynCommnet = decision
        # self.decision_label.config(text=f"코멘트 남기기 여부: {decision}")
        
    def on_yes(self):
        # 코멘트 여부 확인
        print('엑셀 다운로드 함수 실행!')
        if self.ynCommnet =='예':
            print('코멘트 남기기여부 : ' + self.ynCommnet)
        
  