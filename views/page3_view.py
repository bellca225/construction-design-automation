import tkinter as tk
from tkinter import ttk

class Page3View(ttk.Frame):
    # 주호 변경 예정
    # data['twisted_cable']['4p']
    # data['connector_jack']['RS-232C(10Pin)']
    # data['control_cable']['2C']['2.5mm']
    # data['Communication_premises_power_cable']['25mm']
    # 마지막 FR 케이블은 현영이한테 확인받을 거임
    
    품셈 = {
        '꼬임 케이블 포설' : 1, 
        '커넥터 및 Jack 접속' : 2, 
        '제어 케이블' : 3,
        '통신용 구내 전력케이블' : 4, 
        'FR 케이블' : 5
        }
    노무임 = {
        '꼬임 케이블 포설' : '통신케이블공', 
        '커넥터 및 Jack 접속' : '통신내선공', 
        '제어 케이블' : '통신케이블공',
        '통신용 구내 전력케이블' : '통신케이블공', 
        'FR 케이블' : '통신케이블공'
        }
    raw_data = {
        '간접노무비': 12.2, 
        '기타경비': 5.8, 
        '일반관리비': 6.0, 
        '이윤': 15.0, 
        '고용보험료': 1.57, 
        '건강보험료': 3.545, 
        '노인장기요양보험료': 12.81, 
        '연금보험료': 4.5, 
        '산재보험료' : 3.7, 
        'twisted_cable': {'4p': 0.15}, 
        'connector_jack': {'RS-232C(10Pin)': 0.49, 
        'Modular(RJ45-8PinPlug)': 0.13, 
        'Modular(Outlet)': 0.28, 
        'TELCO(50Pin)': 1.19, 
        'Data Line': 0.84}, 
        'control_cable': {'2C': {'2.5mm': 0.1, }}, #여기 } 안 닫혀서 추가함
        '통신내선공': 267277, 
        '통신설비공': 296882, 
        '통신외선공': 387376, 
        '통신케이블공': 414944, 
        '무선안테나공': 339642, 
        '광케이블설치사': 444142, 
        'H/W시험사': 375020, 
        'S/W시험사': 433747, 
        '통신관련기사': 305806, 
        '통신관련산업기사': 294019, 
        '통신관련기능사': 242587
        }
    
    var_selected_texts = []
    
    ynConstructDate = ''
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="4. 노무임 확인")
        lb1.place(relx=0.05, rely=0.05)
        
        relx = 0.05
        lbs = ['작업', '수량', '단가', '금액']
        for lb in lbs :
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
        print('공사 30일 이상 여부 : ' + self.ynConstructDate)
        print('확인, 밑의 영역 채우기')
        
        lbs1 = ['기타경비', '일반관리비', '산재보험료', '고용보험료', '이윤']
        lbs2 = ['건강보험료', '노인장기요양보험료', '연금보험료']
        
        if self.ynConstructDate == '예':
            combined_labels = lbs1 + lbs2
        else:
            combined_labels = lbs1
        
        title = ttk.Label(self, text="5. 경비 확인")
        title.place(relx=0.05, rely=0.52)
        self.labels.append(title)
        total_cost = 0
        경비_list = [self.raw_data[key] for key in combined_labels]
        for idx in range(len(combined_labels)):
            label_text = combined_labels[idx]
            label = ttk.Label(self, text=label_text)
            label.place(relx = 0.05, rely = 0.57 + idx * 0.03)
            self.labels.append(label)
            number_value = 경비_list[idx]  
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
        품셈_list = [self.품셈[key] for key in self.var_selected_texts]
        노무임1_list = [self.노무임[key] for key in self.var_selected_texts]
        노무임2_list = [self.raw_data[key] for key in 노무임1_list]

        # 기존에 생성된 위젯이 entry 제외하고 모두 제거
        for label in self.labels:
            label.destroy()

        self.labels.clear()
        
        total_cost = 0  # 노무임 합계를 저장할 변수
        user_count = []
        # 동적으로 라벨 3개와 인풋 1개를 배열의 각 원소마다 생성
        for entry in self.entries:
                print(entry.get())
                user_count.append(entry.get())
                
        for idx, text in enumerate(self.var_selected_texts):
            # 라벨 이름
            label1 = ttk.Label(self, text=f"{text}")
            label1.place(relx=0.05, rely=0.15 + idx * 0.03)
            self.labels.append(label1)

            # 수량
            # entry = ttk.Entry(self)
            # entry.place(relx=0.3, rely=0.15 + idx * 0.035, width=50, height=20)
            # self.entries.append(entry)
            # entry.insert(0, "1")
            
        
            단가 = 품셈_list[idx] * 노무임2_list[idx]
            label2 = ttk.Label(self, text = 단가)
            label2.place(relx=0.55, rely=0.15 + idx * 0.035)
            self.labels.append(label2)

            # 금액 = 수량 * 단가 넣어줘야 함! 
            cost_value = 단가 * int(user_count[idx])
            label3 = ttk.Label(self, text = cost_value)
            label3.place(relx=0.80, rely=0.15 + idx * 0.035)
            self.labels.append(label3)
            
            # 노무임 합계 계산
            total_cost += cost_value
            
            # 노무임 합계 라벨 생성 및 표시
            label_total_cost = ttk.Label(self, text=f"노무비 합계: {total_cost}")
            label_total_cost.place(relx=0.05, rely=0.125 + len(self.var_selected_texts) * 0.05)
            self.labels.append(label_total_cost)
            
            #  간접 노무비 * 0.13
            label_total_cost = ttk.Label(self, text=f"간접 노무비: {total_cost * 0.13}")
            label_total_cost.place(relx=0.05, rely=0.155 + len(self.var_selected_texts) * 0.05)
            self.labels.append(label_total_cost)
            
        print(self.var_selected_texts)
        

    def on_next(self):
        self.controller.show_frame("Page4View")

    def update_data(self, selected_texts, decision = None):
        self.var_selected_texts = selected_texts
        품셈_list = [self.품셈[key] for key in selected_texts]
        노무임1_list = [self.노무임[key] for key in selected_texts]
        노무임2_list = [self.raw_data[key] for key in 노무임1_list]
        
        if decision != None :
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
            
            단가 = 품셈_list[idx] * 노무임2_list[idx]
            label2 = ttk.Label(self, text = 단가)
            label2.place(relx=0.55, rely=0.15 + idx * 0.035)
            self.labels.append(label2)

            # 금액 = 수량 * 단가 넣어줘야 함! 
            cost_value = 단가
            label3 = ttk.Label(self, text = cost_value)
            label3.place(relx=0.80, rely=0.15 + idx * 0.035)
            self.labels.append(label3)
            
            # 노무임 합계 계산
            total_cost += cost_value
            
            # 노무임 합계 라벨 생성 및 표시
            label_total_cost = ttk.Label(self, text=f"노무비 합계: {total_cost}")
            label_total_cost.place(relx=0.05, rely=0.125 + len(selected_texts) * 0.05)
            self.labels.append(label_total_cost)
            
            #  간접 노무비 * 0.13
            label_total_cost = ttk.Label(self, text=f"간접 노무비: {total_cost * 0.13}")
            label_total_cost.place(relx=0.05, rely=0.155 + len(selected_texts) * 0.05)
            self.labels.append(label_total_cost)