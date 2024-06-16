# import tkinter as tk

# class Page2View(tk.Frame):
#     def __init__(self, parent, controller):
#         tk.Frame.__init__(self, parent)
#         self.controller = controller
        
#         label = tk.Label(self, text="Page 1")
#         label.pack(pady=10)

#         button = tk.Button(self, text="Go to Home", command=lambda: controller.show_frame("HomeView"))
#         button.pack()


import tkinter as tk
from tkinter import ttk

class Page2View(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="해당 공사가 30일 이상입니까?.")
        lb1.place(relx=0.05, rely=0.05)

        self.button_yes = ttk.Button(self, text="예", command=self.on_yes)
        self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="아니오", command=self.on_no)
        self.button_no.place(relx=0.25, rely=0.3)
        
        # button_home = ttk.Button(self, text="처음으로", command=lambda: controller.show_frame("homeView"))
        # button_home.place(relx=0.5, rely=0.3)

    def on_yes(self):
        self.controller.selection = "예"
        self.controller.show_frame("Page3View")

    def on_no(self):
        self.controller.selection = "아니오"
        self.controller.show_frame("Page3View")