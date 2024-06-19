from tkinter import ttk
import tkinter.font
from views.home_view import HomeView
from views.page1_view import Page1View
from views.page2_view import Page2View
from views.page3_view import Page3View
from views.page4_view import Page4View
from views.page5_view import Page5View
from views.page6_view import Page6View
from ttkthemes import ThemedTk
from dataexporter import getData
from loguru import logger
import warnings 
# 경고 메시지 숨기기
warnings.filterwarnings("ignore")
class MainApp(ThemedTk):
    def __init__(self, *args, **kwargs):
        ThemedTk.__init__(self, *args, **kwargs)
        logger.info('MainApp init')
        self.set_theme("arc") 
        self.geometry('600x600+100+50')
        self.title('자동 통신공사 설계 프로그램')
        self.selected_texts = []  # 선택된 텍스트를 저장할 리스트
        self.decision1 = None  # 공사 일수 결정을 저장할 변수
        self.decision2 = None  # 코멘트 작성 여부를 저장할 변수 (시트 하나 추가)
        logger.info('Getting data from pdf, xlsx, hwp files..')
        self.raw_data = getData()
        logger.info('Getting data from files done!')
        self.excel_data = {
            '재료비' : 0,
            '노무비' : 0,
            '직접노무비' : 0,
            '간접노무비' : 0,
            '경비' : 0,
            '일반관리비' : 0,
            '이윤' : 0,
            '소계' : 0,
            '부가세' : 0,
            '합계' : 0,
            '재료' : [], # {품명, 수량, 단가}
            '노무' : [], # {품명, 수량, 단가}
            '공구손료' : 0,
            '산재보험료' : 0,
            '산재보험료계수' : 0,
            '고용보험료' : 0,
            '고용보험료계수' : 0,
            '기타경비' : 0,
            '기타경비계수' : 0,
            '일반관리비' : 0,
            '일반관리비계수' : 0,
            '이윤일반공사' : 0,
            '이윤일반공사계수' : 0.15,
            '이윤간이공사' : 0,
            '이윤간이공사계수' : 0.1,
        }  # 엑셀 데이터를 저장할 딕셔너리
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
        elif page_name == "Page4View" and hasattr(self.frames["Page3View"], "var_total_cost_value"):
            total_cost_value = self.frames["Page3View"].var_total_cost_value
            frame.update_data(total_cost_value)
        elif hasattr(frame, 'update_data'):
            frame.update_data(*args)
        frame.tkraise()
        

    def set_selected_texts(self, selected_texts):
        print('selected_texts:', selected_texts)
        self.selected_texts = selected_texts

    def set_construct_date_decision(self, decision):
        self.decision1 = decision
        
    def set_comment_decision(self, decision):
        self.decision2 = decision

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
