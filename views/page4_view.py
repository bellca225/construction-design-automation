from tkinter import ttk

class Page4View(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="6. 공사금액이 맞습니까?")
        lb1.place(relx=0.05, rely=0.05)
        
        self.label_total_cost = ttk.Label(self, text="")
        self.label_total_cost.place(relx=0.05, rely=0.1)

        self.button_yes = ttk.Button(self, text="예", command=self.on_yes)
        
        self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="아니오", command=self.on_no)
        self.button_no.place(relx=0.24, rely=0.3)

    def on_yes(self):
        self.controller.show_frame("Page5View")

    def on_no(self):
        self.controller.show_frame("Page3View")
        
    def update_data(self, total_cost_value):
        # Page3View에서 전달받은 값을 Label에 표시
        self.label_total_cost.config(text=str(total_cost_value))