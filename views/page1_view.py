
import tkinter as tk
from tkinter import ttk

class Page1View(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        relyNum = 0.1
        lb1 = ttk.Label(self, text="해당 작업사항이 맞습니까?")
        lb1.place(relx=0.05, rely=0.05)

        self.values_label = ttk.Label(self, text="선택 없음")
        self.values_label.place(relx=0.05, rely=0.1)
        
        self.button2 = ttk.Button(self, text="예", command=lambda: controller.show_frame("Page2View"))
        self.button2.place(relx=0.05, rely=0.3)
        self.button1 = ttk.Button(self, text="아니오 (이전으로)", command=lambda: controller.show_frame("HomeView"))
        self.button1.place(relx=0.25, rely=0.3)

        
        # self.button.pack()

    def update_data(self, selected_texts):
        # 선택된 값을 업데이트하고 라벨에 표시
        self.values_label.config(text=f"{', '.join(selected_texts)}")
