import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import random
import json
import os

class StudentSeatArrangement:
    def __init__(self, root):
        self.root = root
        self.root.title("학생 자리 배정 프로그램")
        self.root.geometry("1200x800")
        self.root.configure(bg="#F5F5F5")
        
        # 파스텔 톤 색상 정의
        self.colors = {
            "bg": "#F5F5F5",  # 배경색 (연한 회색)
            "frame": "#FFFFFF",  # 프레임 배경색 (흰색)
            "button": "#B8E6B8",  # 버튼 배경색 (연한 민트)
            "button_hover": "#9DD89D",  # 버튼 호버 색상
            "seat": "#E8F4F8",  # 일반 자리 색상 (연한 하늘)
            "seat_selected": "#FFD6E8",  # 선택된 자리 색상 (연한 핑크)
            "teacher": "#FFE5CC",  # 교탁 색상 (연한 복숭아)
            "front_fixed": "#FFCCE5",  # 앞자리 고정석 색상
            "back_fixed": "#E5CCFF",  # 뒷자리 고정석 색상
            "text": "#333333",  # 텍스트 색상
            "border": "#D0D0D0"  # 테두리 색상
        }
        
        # 변수 초기화
        self.rows = 0
        self.cols = 0
        self.students = []  # 학생 이름 리스트
        self.seats = []  # 2D 좌석 배열
        self.seat_buttons = []  # 좌석 버튼 참조
        self.selected_seats = []  # 선택된 좌석 [(row, col), (row, col)]
        self.front_fixed_seats = set()  # 앞자리 고정석 좌표
        self.back_fixed_seats = set()  # 뒷자리 고정석 좌표
        
        # UI 생성
        self.create_ui()
        
        # 설정 파일 로드
        self.load_settings()
    
    def create_ui(self):
        """UI 생성"""
        # 메인 컨테이너
        main_container = tk.Frame(self.root, bg=self.colors["bg"])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 좌측 설정 패널
        left_panel = tk.Frame(main_container, bg=self.colors["frame"], relief=tk.RAISED, bd=2)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_panel.config(width=350)
        
        # 제목
        title_label = tk.Label(left_panel, text="학생 자리 배정", 
                              font=("맑은 고딕", 18, "bold"),
                              bg=self.colors["frame"], fg=self.colors["text"])
        title_label.pack(pady=20)
        
        # 자리 설정 프레임
        seat_config_frame = tk.LabelFrame(left_panel, text="자리 설정", 
                                         font=("맑은 고딕", 11, "bold"),
                                         bg=self.colors["frame"], fg=self.colors["text"],
                                         padx=15, pady=15)
        seat_config_frame.pack(fill=tk.X, padx=15, pady=10)
        
        # 행 설정
        row_frame = tk.Frame(seat_config_frame, bg=self.colors["frame"])
        row_frame.pack(fill=tk.X, pady=5)
        tk.Label(row_frame, text="행 수:", font=("맑은 고딕", 10),
                bg=self.colors["frame"], fg=self.colors["text"]).pack(side=tk.LEFT, padx=(0, 10))
        self.row_var = tk.StringVar()
        self.row_entry = tk.Entry(row_frame, textvariable=self.row_var, width=10,
                                 font=("맑은 고딕", 10), relief=tk.SUNKEN, bd=2)
        self.row_entry.pack(side=tk.LEFT)
        self.row_entry.bind("<Return>", lambda e: self.create_seat_layout())
        
        # 열 설정
        col_frame = tk.Frame(seat_config_frame, bg=self.colors["frame"])
        col_frame.pack(fill=tk.X, pady=5)
        tk.Label(col_frame, text="열 수:", font=("맑은 고딕", 10),
                bg=self.colors["frame"], fg=self.colors["text"]).pack(side=tk.LEFT, padx=(0, 10))
        self.col_var = tk.StringVar()
        self.col_entry = tk.Entry(col_frame, textvariable=self.col_var, width=10,
                                 font=("맑은 고딕", 10), relief=tk.SUNKEN, bd=2)
        self.col_entry.pack(side=tk.LEFT)
        self.col_entry.bind("<Return>", lambda e: self.create_seat_layout())
        
        # 학생 명단 프레임
        student_frame = tk.LabelFrame(left_panel, text="학생 명단", 
                                     font=("맑은 고딕", 11, "bold"),
                                     bg=self.colors["frame"], fg=self.colors["text"],
                                     padx=15, pady=15)
        student_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # 학생 추가 프레임
        add_frame = tk.Frame(student_frame, bg=self.colors["frame"])
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(add_frame, text="이름:", font=("맑은 고딕", 10),
                bg=self.colors["frame"], fg=self.colors["text"]).pack(side=tk.LEFT, padx=(0, 5))
        self.student_name_var = tk.StringVar()
        self.student_entry = tk.Entry(add_frame, textvariable=self.student_name_var, width=15,
                                     font=("맑은 고딕", 10), relief=tk.SUNKEN, bd=2)
        self.student_entry.pack(side=tk.LEFT, padx=(0, 5))
        self.student_entry.bind("<Return>", lambda e: self.add_student())
        
        add_btn = tk.Button(add_frame, text="추가", command=self.add_student,
                           bg=self.colors["button"], fg=self.colors["text"],
                           font=("맑은 고딕", 9, "bold"),
                           relief=tk.RAISED, bd=2, padx=10, pady=3,
                           cursor="hand2")
        add_btn.pack(side=tk.LEFT)
        add_btn.bind("<Enter>", lambda e: add_btn.config(bg=self.colors["button_hover"]))
        add_btn.bind("<Leave>", lambda e: add_btn.config(bg=self.colors["button"]))
        
        # 학생 목록 (스크롤 가능한 텍스트 영역)
        list_frame = tk.Frame(student_frame, bg=self.colors["frame"])
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.student_listbox = tk.Listbox(list_frame, font=("맑은 고딕", 10),
                                          selectmode=tk.SINGLE, relief=tk.SUNKEN, bd=2)
        self.student_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.student_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.student_listbox.config(yscrollcommand=scrollbar.set)
        
        # 학생 삭제 버튼
        delete_btn = tk.Button(student_frame, text="선택 삭제", command=self.delete_student,
                              bg=self.colors["button"], fg=self.colors["text"],
                              font=("맑은 고딕", 9, "bold"),
                              relief=tk.RAISED, bd=2, padx=10, pady=3,
                              cursor="hand2")
        delete_btn.pack(pady=(10, 0))
        delete_btn.bind("<Enter>", lambda e: delete_btn.config(bg=self.colors["button_hover"]))
        delete_btn.bind("<Leave>", lambda e: delete_btn.config(bg=self.colors["button"]))
        
        # 고정석 설정 프레임
        fixed_frame = tk.LabelFrame(left_panel, text="고정석 설정", 
                                   font=("맑은 고딕", 11, "bold"),
                                   bg=self.colors["frame"], fg=self.colors["text"],
                                   padx=15, pady=15)
        fixed_frame.pack(fill=tk.X, padx=15, pady=10)
        
        fixed_info = tk.Label(fixed_frame, 
                             text="자리 배정 후 자리를 클릭하여\n앞자리/뒷자리 고정석으로 설정",
                             font=("맑은 고딕", 9),
                             bg=self.colors["frame"], fg=self.colors["text"],
                             justify=tk.LEFT)
        fixed_info.pack(pady=5)
        
        # 앞자리 고정석 버튼
        front_fixed_btn = tk.Button(fixed_frame, text="앞자리 고정석 선택", 
                                   command=self.set_front_fixed_mode,
                                   bg="#FFCCE5", fg=self.colors["text"],
                                   font=("맑은 고딕", 9, "bold"),
                                   relief=tk.RAISED, bd=2, padx=10, pady=5,
                                   cursor="hand2")
        front_fixed_btn.pack(fill=tk.X, pady=5)
        front_fixed_btn.bind("<Enter>", lambda e: front_fixed_btn.config(bg="#FFB3D9"))
        front_fixed_btn.bind("<Leave>", lambda e: front_fixed_btn.config(bg="#FFCCE5"))
        
        # 뒷자리 고정석 버튼
        back_fixed_btn = tk.Button(fixed_frame, text="뒷자리 고정석 선택", 
                                  command=self.set_back_fixed_mode,
                                  bg="#E5CCFF", fg=self.colors["text"],
                                  font=("맑은 고딕", 9, "bold"),
                                  relief=tk.RAISED, bd=2, padx=10, pady=5,
                                  cursor="hand2")
        back_fixed_btn.pack(fill=tk.X, pady=5)
        back_fixed_btn.bind("<Enter>", lambda e: back_fixed_btn.config(bg="#D9B3FF"))
        back_fixed_btn.bind("<Leave>", lambda e: back_fixed_btn.config(bg="#E5CCFF"))
        
        # 고정석 해제 버튼
        clear_fixed_btn = tk.Button(fixed_frame, text="고정석 해제", 
                                   command=self.clear_fixed_mode,
                                   bg=self.colors["button"], fg=self.colors["text"],
                                   font=("맑은 고딕", 9, "bold"),
                                   relief=tk.RAISED, bd=2, padx=10, pady=5,
                                   cursor="hand2")
        clear_fixed_btn.pack(fill=tk.X, pady=5)
        clear_fixed_btn.bind("<Enter>", lambda e: clear_fixed_btn.config(bg=self.colors["button_hover"]))
        clear_fixed_btn.bind("<Leave>", lambda e: clear_fixed_btn.config(bg=self.colors["button"]))
        
        # 버튼 프레임
        button_frame = tk.Frame(left_panel, bg=self.colors["frame"])
        button_frame.pack(fill=tk.X, padx=15, pady=15)
        
        # 배치 버튼
        arrange_btn = tk.Button(button_frame, text="배치", command=self.arrange_seats,
                               bg="#B8E6B8", fg=self.colors["text"],
                               font=("맑은 고딕", 12, "bold"),
                               relief=tk.RAISED, bd=3, padx=20, pady=10,
                               cursor="hand2")
        arrange_btn.pack(fill=tk.X, pady=(0, 10))
        arrange_btn.bind("<Enter>", lambda e: arrange_btn.config(bg="#9DD89D"))
        arrange_btn.bind("<Leave>", lambda e: arrange_btn.config(bg="#B8E6B8"))
        
        # 저장 버튼
        save_btn = tk.Button(button_frame, text="설정 저장", command=self.save_settings,
                           bg=self.colors["button"], fg=self.colors["text"],
                           font=("맑은 고딕", 10, "bold"),
                           relief=tk.RAISED, bd=2, padx=10, pady=5,
                           cursor="hand2")
        save_btn.pack(fill=tk.X)
        save_btn.bind("<Enter>", lambda e: save_btn.config(bg=self.colors["button_hover"]))
        save_btn.bind("<Leave>", lambda e: save_btn.config(bg=self.colors["button"]))
        
        # 우측 자리 배치 패널
        right_panel = tk.Frame(main_container, bg=self.colors["frame"], relief=tk.RAISED, bd=2)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 자리 배치 컨테이너
        self.seat_container = tk.Frame(right_panel, bg=self.colors["frame"])
        self.seat_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 안내 레이블
        self.guide_label = tk.Label(self.seat_container, 
                                   text="행과 열을 입력하고 Enter를 누르면 자리 배치가 표시됩니다.",
                                   font=("맑은 고딕", 12),
                                   bg=self.colors["frame"], fg=self.colors["text"])
        self.guide_label.pack(pady=50)
        
        # 편집 모드 변수
        self.edit_mode = "swap"  # "swap", "front_fixed", "back_fixed", "clear_fixed"
    
    def add_student(self):
        """학생 추가"""
        name = self.student_name_var.get().strip()
        if not name:
            messagebox.showwarning("경고", "학생 이름을 입력해주세요.")
            return
        
        if name in self.students:
            messagebox.showwarning("경고", "이미 존재하는 학생입니다.")
            return
        
        self.students.append(name)
        self.student_listbox.insert(tk.END, name)
        self.student_name_var.set("")
        self.student_entry.focus()
    
    def delete_student(self):
        """학생 삭제"""
        selection = self.student_listbox.curselection()
        if not selection:
            messagebox.showwarning("경고", "삭제할 학생을 선택해주세요.")
            return
        
        index = selection[0]
        name = self.student_listbox.get(index)
        self.students.remove(name)
        self.student_listbox.delete(index)
        
        # 자리에서도 제거
        if self.seats:
            for r in range(len(self.seats)):
                for c in range(len(self.seats[r])):
                    if self.seats[r][c] == name:
                        self.seats[r][c] = ""
                        if self.seat_buttons and r < len(self.seat_buttons) and c < len(self.seat_buttons[r]):
                            self.seat_buttons[r][c].config(text="")
                            self.update_seat_color(r, c)
    
    def create_seat_layout(self):
        """자리 레이아웃 생성"""
        try:
            rows = int(self.row_var.get())
            cols = int(self.col_var.get())
            if rows <= 0 or cols <= 0:
                messagebox.showerror("오류", "행과 열은 1 이상이어야 합니다.")
                return
        except ValueError:
            messagebox.showerror("오류", "행과 열에 숫자를 입력해주세요.")
            return
        
        self.rows = rows
        self.cols = cols
        
        # 기존 자리 초기화
        self.seats = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        
        # 기존 위젯 제거
        for widget in self.seat_container.winfo_children():
            widget.destroy()
        
        # 자리 배치 프레임
        seat_frame = tk.Frame(self.seat_container, bg=self.colors["frame"])
        seat_frame.pack(expand=True)
        
        # 교탁 추가 (맨 위)
        teacher_desk = tk.Label(seat_frame, text="교탁", 
                               font=("맑은 고딕", 14, "bold"),
                               bg=self.colors["teacher"], fg=self.colors["text"],
                               relief=tk.RAISED, bd=3, width=15, height=2)
        teacher_desk.grid(row=0, column=0, columnspan=self.cols, pady=(0, 30), padx=5)
        
        # 좌석 버튼 생성
        self.seat_buttons = []
        for r in range(self.rows):
            row_buttons = []
            for c in range(self.cols):
                btn = tk.Button(seat_frame, text="", width=12, height=2,
                              font=("맑은 고딕", 9),
                              bg=self.colors["seat"], fg=self.colors["text"],
                              relief=tk.RAISED, bd=2,
                              command=lambda row=r, col=c: self.on_seat_click(row, col),
                              cursor="hand2")
                btn.grid(row=r+1, column=c, padx=3, pady=3)
                row_buttons.append(btn)
            self.seat_buttons.append(row_buttons)
    
    def set_front_fixed_mode(self):
        """앞자리 고정석 모드 설정"""
        self.edit_mode = "front_fixed"
        messagebox.showinfo("알림", "앞자리 고정석으로 설정할 자리를 클릭하세요.")
    
    def set_back_fixed_mode(self):
        """뒷자리 고정석 모드 설정"""
        self.edit_mode = "back_fixed"
        messagebox.showinfo("알림", "뒷자리 고정석으로 설정할 자리를 클릭하세요.")
    
    def clear_fixed_mode(self):
        """고정석 해제 모드 설정"""
        self.edit_mode = "clear_fixed"
        messagebox.showinfo("알림", "고정석을 해제할 자리를 클릭하세요.")
    
    def on_seat_click(self, row, col):
        """자리 클릭 이벤트 처리"""
        if self.edit_mode == "swap":
            # 자리 교환 모드
            if (row, col) in self.selected_seats:
                # 이미 선택된 자리면 선택 해제
                self.selected_seats.remove((row, col))
                self.update_seat_color(row, col)
            else:
                # 새 자리 선택
                self.selected_seats.append((row, col))
                self.seat_buttons[row][col].config(bg=self.colors["seat_selected"])
                
                # 두 자리가 선택되면 교환
                if len(self.selected_seats) == 2:
                    self.swap_seats()
                    self.selected_seats = []
        
        elif self.edit_mode == "front_fixed":
            # 앞자리 고정석 설정
            pos = (row, col)
            if pos in self.back_fixed_seats:
                self.back_fixed_seats.remove(pos)
            if pos in self.front_fixed_seats:
                self.front_fixed_seats.remove(pos)
                self.update_seat_color(row, col)
            else:
                self.front_fixed_seats.add(pos)
                self.update_seat_color(row, col)
        
        elif self.edit_mode == "back_fixed":
            # 뒷자리 고정석 설정
            pos = (row, col)
            if pos in self.front_fixed_seats:
                self.front_fixed_seats.remove(pos)
            if pos in self.back_fixed_seats:
                self.back_fixed_seats.remove(pos)
                self.update_seat_color(row, col)
            else:
                self.back_fixed_seats.add(pos)
                self.update_seat_color(row, col)
        
        elif self.edit_mode == "clear_fixed":
            # 고정석 해제
            pos = (row, col)
            if pos in self.front_fixed_seats:
                self.front_fixed_seats.remove(pos)
            if pos in self.back_fixed_seats:
                self.back_fixed_seats.remove(pos)
            self.update_seat_color(row, col)
            self.edit_mode = "swap"  # 해제 후 일반 모드로 복귀
    
    def swap_seats(self):
        """선택된 두 자리의 학생 교환"""
        if len(self.selected_seats) != 2:
            return
        
        r1, c1 = self.selected_seats[0]
        r2, c2 = self.selected_seats[1]
        
        # 학생 이름 교환
        self.seats[r1][c1], self.seats[r2][c2] = self.seats[r2][c2], self.seats[r1][c1]
        
        # 버튼 텍스트 업데이트
        self.seat_buttons[r1][c1].config(text=self.seats[r1][c1])
        self.seat_buttons[r2][c2].config(text=self.seats[r2][c2])
        
        # 색상 업데이트
        self.update_seat_color(r1, c1)
        self.update_seat_color(r2, c2)
    
    def update_seat_color(self, row, col):
        """자리 색상 업데이트"""
        if not self.seat_buttons or row >= len(self.seat_buttons) or col >= len(self.seat_buttons[row]):
            return
        
        pos = (row, col)
        
        # 고정석 색상 적용
        if pos in self.front_fixed_seats:
            self.seat_buttons[row][col].config(bg=self.colors["front_fixed"])
        elif pos in self.back_fixed_seats:
            self.seat_buttons[row][col].config(bg=self.colors["back_fixed"])
        else:
            self.seat_buttons[row][col].config(bg=self.colors["seat"])
    
    def arrange_seats(self):
        """학생 자리 배정"""
        # 입력 검증
        try:
            rows = int(self.row_var.get())
            cols = int(self.col_var.get())
            if rows <= 0 or cols <= 0:
                messagebox.showerror("오류", "행과 열은 1 이상이어야 합니다.")
                return
        except ValueError:
            messagebox.showerror("오류", "행과 열에 숫자를 입력해주세요.")
            return
        
        if not self.students:
            messagebox.showerror("오류", "학생 명단을 입력해주세요.")
            return
        
        self.rows = rows
        self.cols = cols
        
        # 자리 레이아웃이 없으면 생성
        if not self.seat_buttons:
            self.create_seat_layout()
        
        # 자리 초기화
        self.seats = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        
        # 고정석에 학생 배정 (이미 배정된 학생 제외)
        fixed_students = set()
        
        # 앞자리 고정석 배정
        front_positions = list(self.front_fixed_seats)
        random.shuffle(front_positions)
        front_students = [s for s in self.students if s not in fixed_students]
        random.shuffle(front_students)
        
        for i, pos in enumerate(front_positions):
            if i < len(front_students) and 0 <= pos[0] < self.rows and 0 <= pos[1] < self.cols:
                self.seats[pos[0]][pos[1]] = front_students[i]
                fixed_students.add(front_students[i])
        
        # 뒷자리 고정석 배정
        back_positions = list(self.back_fixed_seats)
        random.shuffle(back_positions)
        back_students = [s for s in self.students if s not in fixed_students]
        random.shuffle(back_students)
        
        for i, pos in enumerate(back_positions):
            if i < len(back_students) and 0 <= pos[0] < self.rows and 0 <= pos[1] < self.cols:
                self.seats[pos[0]][pos[1]] = back_students[i]
                fixed_students.add(back_students[i])
        
        # 남은 학생들
        remaining_students = [s for s in self.students if s not in fixed_students]
        random.shuffle(remaining_students)
        
        # 남은 자리에 배정
        remaining_positions = []
        for r in range(self.rows):
            for c in range(self.cols):
                pos = (r, c)
                if pos not in self.front_fixed_seats and pos not in self.back_fixed_seats:
                    remaining_positions.append(pos)
        
        random.shuffle(remaining_positions)
        
        for i, student in enumerate(remaining_students):
            if i < len(remaining_positions):
                r, c = remaining_positions[i]
                self.seats[r][c] = student
        
        # UI 업데이트
        for r in range(self.rows):
            for c in range(self.cols):
                if self.seat_buttons and r < len(self.seat_buttons) and c < len(self.seat_buttons[r]):
                    self.seat_buttons[r][c].config(text=self.seats[r][c])
                    self.update_seat_color(r, c)
        
        messagebox.showinfo("완료", "자리 배정이 완료되었습니다.")
        self.edit_mode = "swap"  # 배정 후 일반 모드로 복귀
    
    def save_settings(self):
        """설정 저장"""
        data = {
            "rows": self.rows,
            "cols": self.cols,
            "students": self.students,
            "seats": self.seats,
            "front_fixed_seats": list(self.front_fixed_seats),
            "back_fixed_seats": list(self.back_fixed_seats)
        }
        
        try:
            with open("seat_settings.json", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"설정 저장 중 오류가 발생했습니다.\n{str(e)}")
    
    def load_settings(self):
        """설정 불러오기"""
        try:
            if os.path.exists("seat_settings.json"):
                with open("seat_settings.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # 데이터 로드
                self.rows = data.get("rows", 0)
                self.cols = data.get("cols", 0)
                self.students = data.get("students", [])
                self.seats = data.get("seats", [])
                self.front_fixed_seats = set(tuple(pos) for pos in data.get("front_fixed_seats", []))
                self.back_fixed_seats = set(tuple(pos) for pos in data.get("back_fixed_seats", []))
                
                # UI 업데이트
                if self.rows > 0 and self.cols > 0:
                    self.row_var.set(str(self.rows))
                    self.col_var.set(str(self.cols))
                    self.create_seat_layout()
                
                # 학생 목록 업데이트
                for student in self.students:
                    self.student_listbox.insert(tk.END, student)
                
                # 자리 배정 표시
                if self.seats and self.seat_buttons:
                    for r in range(len(self.seats)):
                        for c in range(len(self.seats[r])):
                            if r < len(self.seat_buttons) and c < len(self.seat_buttons[r]):
                                self.seat_buttons[r][c].config(text=self.seats[r][c])
                                self.update_seat_color(r, c)
        except Exception as e:
            print(f"설정 로드 오류: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentSeatArrangement(root)
    root.mainloop()
