import tkinter as tk
from tkinter import ttk

class Page2View(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="3. 해당 공사가 30일 이상입니까?")
        lb1.place(relx=0.05, rely=0.05)

        self.button_yes = ttk.Button(self, text="예", command=self.on_yes)
        self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="아니오", command=self.on_no)
        self.button_no.place(relx=0.24, rely=0.3)   
        

    def on_yes(self):
        self.controller.set_construct_date_decision("예")
        self.controller.show_frame("Page3View")

    def on_no(self):
        self.controller.set_construct_date_decision("아니오")
        self.controller.show_frame("Page3View")
        