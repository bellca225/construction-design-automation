import tkinter as tk
from tkinter import ttk
import tkinter.font
from views.home_view import HomeView
from views.page1_view import Page1View
from views.page2_view import Page2View
from views.page3_view import Page3View
from views.page4_view import Page4View
from views.page5_view import Page5View
from views.page6_view import Page6View
from dataexporter import getData
from ttkthemes import ThemedTk
class MainApp(ThemedTk):
    def __init__(self, *args, **kwargs):
        ThemedTk.__init__(self, *args, **kwargs)
        self.set_theme("arc") 
        self.geometry('600x600+100+100')
        self.title('자동 통신공사 설계 프로그램')
        self.data = getData()
        print(self.data['twisted_cable']['4p'])
        self.selected_texts = []  # 선택된 텍스트를 저장할 리스트
        self.decision1 = None  # 공사 일수 결정을 저장할 변수
        self.decision2 = None  # 코멘트 작성 여부를 저장할 변수 (시트 하나 추가)
        # 스타일 정의
        font1=tkinter.font.Font(family="맑은 고딕", size=10)
        fontBtn=tkinter.font.Font(family="맑은 고딕", size=10, weight = 'bold')
        style = ttk.Style()
        style.configure('TFrame', background='#f5f6f7')
        style.configure('TLabel', background='#f5f6f7',foreground='#1d2a4c', font=font1)
        style.configure('TButton', background='#f5f6f7', foreground='#0673bd', font=fontBtn)
        style.configure('TCheckbutton', foreground='#1d2a4c', font=font1)
        
        self.container = ttk.Frame(self)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        
        self.chkbtnNm = ['꼬임 케이블 포설', '커넥터 및 Jack 접속', '제어 케이블', '통신용 구내 전력케이블', 'FR 케이블']

        self.frames = {}
        for F in (HomeView, Page1View, Page2View, Page3View, Page4View, Page5View, Page6View):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        self.show_frame("HomeView")
        
    def show_frame(self, page_name, *args):
        frame = self.frames[page_name]
        if page_name == "Page3View":
            frame.update_data(self.selected_texts, self.decision1)
        elif page_name == "Page6View":
            frame.update_data(self.decision2)
        elif hasattr(frame, 'update_data'):
            frame.update_data(*args)
        frame.tkraise()

    def set_selected_texts(self, selected_texts):
        self.selected_texts = selected_texts

    def set_construct_date_decision(self, decision):
        self.decision1 = decision
        
    def set_comment_decision(self, decision):
        self.decision2 = decision

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
