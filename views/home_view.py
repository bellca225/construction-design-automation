import tkinter as tk
from tkinter import ttk

class HomeView(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        relyNum = 0.1
        chkbtnNm = controller.chkbtnNm
        lb1 = ttk.Label(self, text="작업을 선택하세요.")
        lb1.place(relx=0.05, rely=0.05)

        self.var_list = []  # 체크버튼 변수를 저장할 리스트

        for bNm in chkbtnNm:
            var = tk.IntVar()
            chkbtn = ttk.Checkbutton(self, text=bNm, variable=var)
            chkbtn.place(relx=0.05, rely=relyNum)
            relyNum += 0.03
            self.var_list.append(var)  # 변수를 리스트에 추가

        button1 = ttk.Button(self, text="선택 초기화", command=self.reset_selection)
        button1.place(relx=0.05, rely=0.3)
        button2 = ttk.Button(self, text="완료", command=self.on_complete)
        button2.place(relx=0.25, rely=0.3)

    def reset_selection(self):
        for var in self.var_list:
            var.set(0)  # 각 체크버튼 변수를 0으로 설정하여 체크 해제

    def on_complete(self):
        selected_indices = [i for i, var in enumerate(self.var_list) if var.get() == 1]
        selected_texts = [self.controller.chkbtnNm[i] for i in selected_indices]
        self.controller.show_frame("Page1View", selected_texts)
