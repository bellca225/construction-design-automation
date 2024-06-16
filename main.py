import tkinter as tk
from tkinter import ttk
from views.home_view import HomeView
from views.page1_view import Page1View
from views.page2_view import Page2View
from views.page3_view import Page3View

class MainApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.geometry('500x600+100+100')
        self.title('자동 통신공사 설계 프로그램')

        # 스타일 정의
        style = ttk.Style()
        # style.configure('TFrame', background='#1d2a4c')
        # style.configure('TLabel', background='#1d2a4c', foreground='#ffffff', font=('맑은고딕', 12))
        # style.configure('TButton', background='#070e1d', foreground='#ffffff', font=('맑은고딕', 10))
        self.container = ttk.Frame(self)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.chkbtnNm = ['꼬임 케이블 포설', '커넥터 및 Jack 접속', '제어 케이블', '통신용 구내 전력케이블', 'FR 케이블']

        self.frames = {}
        for F in (HomeView, Page1View, Page2View, Page3View):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        self.show_frame("HomeView")

    def show_frame(self, page_name, *args):
        frame = self.frames[page_name]
        if hasattr(frame, 'update_data'):
            frame.update_data(*args)  # 새 데이터로 프레임 업데이트
        frame.tkraise()

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
