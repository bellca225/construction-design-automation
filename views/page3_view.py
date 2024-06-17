import tkinter as tk
from tkinter import ttk

class Page3View(ttk.Frame):
    ynConstructDate = ''
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="4. 노무임 확인")
        lb1.place(relx=0.05, rely=0.05)
        
        relx = 0.05
        lbs = ['작업', '수량', '단가', '금액']
        for lb in lbs :
            print(lb)
            lb = ttk.Label(self, text = lb)
            lb.place(relx = relx, rely = 0.1)
            relx += 0.25

        button_yes = ttk.Button(self, text="수량 적용", command=self.on_cnt_confirm)
        button_yes.place(relx=0.05, rely=0.45)
    
        button_yes = ttk.Button(self, text="경비 확인", command=self.on_confirm)
        button_yes.place(relx=0.24, rely=0.45)

        self.button_no = ttk.Button(self, text="다음", command=self.on_next)
        self.button_no.place(relx=0.05, rely=0.88)
        
        self.labels = []
        self.entries = []

    def on_confirm(self):
        print('엑셀 다운로드 함수 실행!')
        print('공사 30일 이상 여부 : ' + self.ynConstructDate)
        print('확인, 밑의 영역 채우기')
        
        lbs1 = ['기타경비', '일반관리비', '산재보험료', '고용보험료', '이윤']
        lbs2 = ['건강보험', '노인장기요양', '연금보험']
        
        if self.ynConstructDate == '예':
            combined_labels = lbs1 + lbs2
        else:
            combined_labels = lbs1
        
        title = ttk.Label(self, text="5. 경비 확인")
        title.place(relx=0.05, rely=0.52)
        self.labels.append(title)
        total_cost = 0
        for idx in range(len(combined_labels)):
            label_text = combined_labels[idx]
            label = ttk.Label(self, text=label_text)
            label.place(relx = 0.05, rely = 0.57 + idx * 0.03)
            self.labels.append(label)
            # 예시 값, 실제 숫자 값에 맞게 수정해야 함
            # lbs1 혹은 lbs1 + lbs2의 순서대로 가격 입력 후 idx로 접근
            number_value = 50  
            label_number = ttk.Label(self, text=number_value)
            label_number.place(relx=0.50, rely=0.57 + idx * 0.03)
            self.labels.append(label_number)
            total_cost += number_value
            
        label_total_cost = ttk.Label(self, text=f"총 합계")
        label_total_cost.place(relx=0.05, rely=0.81)
        self.labels.append(label_total_cost)
        label_total_cost_value = ttk.Label(self, text=f"{total_cost}")
        label_total_cost_value.place(relx=0.5, rely=0.81)
        self.labels.append(label_total_cost_value)
    
    def on_cnt_confirm(self):
        print(1)

    def on_next(self):
        self.controller.show_frame("Page4View")

    def update_data(self, selected_texts, decision):
        self.ynConstructDate = decision
        
        # 기존에 생성된 위젯이 있다면 모두 제거
        for label in self.labels:
            label.destroy()
        for entry in self.entries:
            entry.destroy()

        self.labels.clear()
        self.entries.clear()
        
        total_cost = 0  # 노무임 합계를 저장할 변수

        # 동적으로 라벨 3개와 인풋 1개를 배열의 각 원소마다 생성
        for idx, text in enumerate(selected_texts):
            # 라벨 이름
            label1 = ttk.Label(self, text=f"{text}")
            label1.place(relx=0.05, rely=0.15 + idx * 0.03)
            self.labels.append(label1)

            # 수량
            entry = ttk.Entry(self)
            entry.place(relx=0.3, rely=0.15 + idx * 0.035, width=50, height=20)
            self.entries.append(entry)
            entry.insert(0, "1")
            
            # 노무임.hwp * 품샘.pdf 넣어줘야 함!
            단가 = 100
            label2 = ttk.Label(self, text = 단가)
            label2.place(relx=0.55, rely=0.15 + idx * 0.035)
            self.labels.append(label2)

            # 금액 = 수량 * 단가 넣어줘야 함! 
            cost_value = 200
            label3 = ttk.Label(self, text = cost_value)
            label3.place(relx=0.80, rely=0.15 + idx * 0.035)
            self.labels.append(label3)
            
            # 노무임 합계 계산
            total_cost += cost_value
            
            # 노무임 합계 라벨 생성 및 표시
            label_total_cost = ttk.Label(self, text=f"노무비 합계: {total_cost}")
            label_total_cost.place(relx=0.05, rely=0.13 + len(selected_texts) * 0.05)
            self.labels.append(label_total_cost)
            
            #  간접 노무비 * 0.13
            label_total_cost = ttk.Label(self, text=f"간접 노무비: {total_cost * 0.13}")
            label_total_cost.place(relx=0.05, rely=0.16 + len(selected_texts) * 0.05)
            self.labels.append(label_total_cost)