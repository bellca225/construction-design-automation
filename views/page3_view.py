import tkinter as tk
from tkinter import ttk

class Page3View(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="노무임")
        lb1.place(relx=0.05, rely=0.05)
        
        relx = 0.05
        lbs = ['작업', '수량', '단가', '금액']
        for lb in lbs :
            lb = ttk.Label(self, text = lb)
            lb.place(relx = relx, rely = 0.1)
            relx += 0.08
            
            
        self.selected_texts_label = ttk.Label(self, text="선택된 텍스트: ")
        self.selected_texts_label.pack()

        self.decision_label = ttk.Label(self, text="공사 30일 이상 여부: ")
        self.decision_label.pack()
        
        # self.button_yes = ttk.Button(self, text="예", command=self.on_yes)
        # self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="다음", command=self.on_no)
        self.button_no.place(relx=0.25, rely=0.3)

    # def on_yes(self):
    #     self.controller.selection = "예"
    #     self.controller.show_frame("Page3View")

    def on_no(self):
        self.controller.show_frame("Page4View")

    def update_data(self, selected_texts, decision):
        self.selected_texts_label.config(text=f"선택된 텍스트: {', '.join(selected_texts)}")
        self.decision_label.config(text=f"공사 30일 이상 여부: {decision}")
        
        
        
        

        