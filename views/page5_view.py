import tkinter as tk
from tkinter import ttk
import tkinter.font

class Page5View(ttk.Frame):
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        lb1 = ttk.Label(self, text="7. 엑셀 내 코멘트를 남겨주세요.")
        lb1.place(relx=0.05, rely=0.05)
        
        
        self.entry_text = tk.StringVar()  # 변수를 사용하여 entry의 텍스트를 관리
    
        font=tkinter.font.Font(family="맑은 고딕", size=10)
        self.entry = tk.Text(self, wrap='word', font=font)
        
        self.entry.place(relx=0.05, rely=0.1, width=220, height=80)
        self.entry.bind('<Return>', lambda event: self.entry.insert('end', '\n'))  #
        
        self.button_yes = ttk.Button(self, text="저장", command=self.on_yes)
        self.button_yes.place(relx=0.05, rely=0.3)

        self.button_no = ttk.Button(self, text="취소", command=self.on_no)
        self.button_no.place(relx=0.24, rely=0.3)
    

    def on_yes(self):
        comment = self.entry.get("1.0", 'end-1c')  # entry에 입력된 텍스트 가져오기
        self.controller.set_comment_decision(comment)
        self.entry.delete("1.0", tk.END)
        self.controller.show_frame("Page6View")

    def on_no(self):
        # self.controller.selection = "아니오"
        self.controller.set_comment_decision("")
        self.entry.delete("1.0", tk.END)
        self.controller.show_frame("Page6View")