import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import random
from typing import List, Tuple, Dict, Set
import json
import os

class StudentSeatArrangement:
    def __init__(self, root):
        self.root = root
        self.root.title("학생 자리 배정 프로그램2")
        self.root.geometry("1000x700")
        self.root.configure(bg="#F8F6F2")
        
        # 창 크기 조절 설정
        self.root.resizable(True, True)
        self.root.update_idletasks()  # 모든 대기 중인 UI 작업을 처리
        
        # 색상 정의 (파스텔 톤)
        self.colors = {
            "bg": "#F8F6F2",  # 배경색 (연한 베이지)
            "button": "#E8D0D0",  # 버튼 배경색 (연한 핑크)
            "button_active": "#D8B0B0",  # 버튼 활성화 색상
            "frame": "#E8E0D8",  # 프레임 배경색 (연한 베이지)
            "seat": "#D0E8E0",  # 일반 자리 색상 (연한 민트)
            "seat_selected": "#FFD0D0",  # 선택된 자리 색상 (연한 코랄)
            "front_fixed": "#FFD0E0",  # 앞쪽 고정석 색상 (연한 핑크)
            "back_fixed": "#FFD0E0",  # 뒤쪽 고정석 색상 (연한 핑크)
            "normal_fixed": "#FFD0E0",  # 일반석 고정석 색상 (연한 핑크)
            "teacher": "#C0E0FF",  # 교탁 색상 (연한 하늘)
            "front_area": "#FFEFD5",  # 앞쪽 영역 표시 색상 (연한 복숭아)
            "back_area": "#E6E6FA",   # 뒤쪽 영역 표시 색상 (연한 라벤더)
            "normal_area": "#E0F5E9",  # 일반석 영역 표시 색상 (연한 민트)
            "front_area_border": "#FFC090",  # 앞쪽 영역 테두리 색상
            "back_area_border": "#B090FF",   # 뒤쪽 영역 테두리 색상
            "normal_area_border": "#90D0B0",  # 일반석 영역 테두리 색상
            "disabled_seat": "#CCCCCC",  # 비활성화된 자리 색상 (회색)
        }
        
        # 스타일 설정
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TButton', 
                            background=self.colors["button"], 
                            foreground='#333333', 
                            font=('맑은 고딕', 10, 'bold'),
                            borderwidth=0,
                            focuscolor=self.colors["button_active"])
        self.style.map('TButton',
                      background=[('active', self.colors["button_active"])])
        
        # 변수 초기화
        self.rows = 0
        self.cols = 0
        self.students = []  # [{name: 이름, position: None/front/back}]
        self.seats = []  # 2D 좌석 배열
        self.seat_buttons = []  # 좌석 버튼 참조 저장
        self.selected_seats = []  # 선택된 좌석 위치 [(row, col), (row, col)]
        self.edit_mode = "swap"  # 편집 모드: "swap" 또는 "fixed" 또는 "front" 또는 "back" 또는 "normal"
        
        # 앞/뒤/일반 영역 정보
        self.front_area = set()  # 앞쪽 영역으로 지정된 좌표 (r, c)
        self.back_area = set()   # 뒤쪽 영역으로 지정된 좌표 (r, c)
        self.normal_area = set()  # 일반석 영역으로 지정된 좌표 (r, c)
        
        # 고정석 정보: {(row, col): student_name} - 특정 좌표에 고정된 학생
        self.fixed_seats = {}
        
        # 비활성화된 자리 정보: set of (row, col) - 자리 배정에서 제외할 자리
        self.disabled_seats = set()
        
        # 메인 프레임 생성
        self.create_main_frame()
        
        # 설정 파일 로드 (있는 경우)
        self.load_settings()
    
    def create_main_frame(self):
        """메인 UI 프레임 생성"""
        # 좌측 설정 프레임
        self.settings_frame = tk.Frame(self.root, bg=self.colors["bg"], padx=20, pady=20)
        self.settings_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # 제목
        tk.Label(self.settings_frame, text="학생 자리 배정", font=("맑은 고딕", 16, "bold"), 
                bg=self.colors["bg"], fg="#333333").pack(pady=(0, 20))
        
        # 행/열 설정 프레임
        seat_frame = tk.LabelFrame(self.settings_frame, text="자리 설정", font=("맑은 고딕", 10), 
                                  bg=self.colors["bg"], fg="#555555", padx=10, pady=10)
        seat_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 행 설정
        tk.Label(seat_frame, text="행 수:", bg=self.colors["bg"], font=("맑은 고딕", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.row_var = tk.StringVar()
        self.row_entry = ttk.Entry(seat_frame, textvariable=self.row_var, width=10)
        self.row_entry.grid(row=0, column=1, padx=5, pady=5)
        self.row_entry.bind("<KeyRelease>", self.check_row_col_input)
        self.row_entry.bind("<Return>", lambda e: self.create_initial_layout())
        
        # 열 설정
        tk.Label(seat_frame, text="열 수:", bg=self.colors["bg"], font=("맑은 고딕", 10)).grid(row=1, column=0, sticky="w", pady=5)
        self.col_var = tk.StringVar()
        self.col_entry = ttk.Entry(seat_frame, textvariable=self.col_var, width=10)
        self.col_entry.grid(row=1, column=1, padx=5, pady=5)
        self.col_entry.bind("<KeyRelease>", self.check_row_col_input)
        self.col_entry.bind("<Return>", lambda e: self.create_initial_layout())
        
        # 학생 명단 프레임
        student_frame = tk.LabelFrame(self.settings_frame, text="학생 명단", font=("맑은 고딕", 10), 
                                     bg=self.colors["bg"], fg="#555555", padx=10, pady=10)
        student_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 학생 추가 프레임
        add_student_frame = tk.Frame(student_frame, bg=self.colors["bg"])
        add_student_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 학생 이름 입력
        tk.Label(add_student_frame, text="이름:", bg=self.colors["bg"], font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        self.student_name_var = tk.StringVar()
        self.student_name_entry = ttk.Entry(add_student_frame, textvariable=self.student_name_var, width=10)
        self.student_name_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 자리 위치 라디오 버튼
        self.position_var = tk.StringVar(value="normal")
        
        position_frame = tk.Frame(add_student_frame, bg=self.colors["bg"])
        position_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Radiobutton(position_frame, text="일반", variable=self.position_var, value="normal").pack(side=tk.LEFT)
        ttk.Radiobutton(position_frame, text="앞자리", variable=self.position_var, value="front").pack(side=tk.LEFT)
        ttk.Radiobutton(position_frame, text="뒷자리", variable=self.position_var, value="back").pack(side=tk.LEFT)
        
        # 학생 추가 버튼
        add_btn = ttk.Button(add_student_frame, text="추가", command=self.add_student)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # 엑셀 파일에서 불러오기 버튼
        import_btn = ttk.Button(add_student_frame, text="엑셀에서 불러오기", command=self.import_students_from_excel)
        import_btn.pack(side=tk.LEFT, padx=5)
        
        # 일괄 추가 프레임
        bulk_add_frame = tk.Frame(student_frame, bg=self.colors["bg"], pady=5)
        bulk_add_frame.pack(fill=tk.X, pady=(5, 0))
        
        tk.Label(bulk_add_frame, text="일괄 추가:", bg=self.colors["bg"], font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        
        # 일괄 추가 개수 입력
        tk.Label(bulk_add_frame, text="개수:", bg=self.colors["bg"], font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        self.bulk_count_var = tk.StringVar(value="24")
        bulk_count_entry = ttk.Entry(bulk_add_frame, textvariable=self.bulk_count_var, width=5)
        bulk_count_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 일괄 추가 버튼
        bulk_add_btn = ttk.Button(bulk_add_frame, text="일괄 추가", command=self.bulk_add_students)
        bulk_add_btn.pack(side=tk.LEFT, padx=5)
        
        # 학생 목록 트리뷰
        self.student_tree = ttk.Treeview(student_frame, columns=("position"), show="tree headings", height=10, selectmode="extended")
        self.student_tree.pack(fill=tk.BOTH, expand=True)
        
        # 스크롤바 추가
        student_scrollbar = ttk.Scrollbar(student_frame, orient="vertical", command=self.student_tree.yview)
        student_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.student_tree.configure(yscrollcommand=student_scrollbar.set)
        
        # 트리뷰 열 설정
        self.student_tree.heading("#0", text="이름")
        self.student_tree.heading("position", text="위치")
        
        self.student_tree.column("#0", width=150, anchor=tk.W)
        self.student_tree.column("position", width=80, anchor=tk.CENTER)
        
        # 학생 목록 편집 버튼 프레임
        student_edit_frame = tk.Frame(student_frame, bg=self.colors["bg"], pady=5)
        student_edit_frame.pack(fill=tk.X)
        
        # 학생 삭제 버튼
        delete_btn = ttk.Button(student_edit_frame, text="삭제", command=self.delete_student)
        delete_btn.pack(side=tk.LEFT, padx=2)
        
        # 학생 위치 변경 버튼
        change_position_frame = tk.Frame(student_edit_frame, bg=self.colors["bg"])
        change_position_frame.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(change_position_frame, text="일반", command=lambda: self.change_student_position("normal")).pack(side=tk.LEFT, padx=2)
        ttk.Button(change_position_frame, text="앞자리", command=lambda: self.change_student_position("front")).pack(side=tk.LEFT, padx=2)
        ttk.Button(change_position_frame, text="뒷자리", command=lambda: self.change_student_position("back")).pack(side=tk.LEFT, padx=2)
        
        # 버튼 프레임
        button_frame = tk.Frame(self.settings_frame, bg=self.colors["bg"], pady=10)
        button_frame.pack(fill=tk.X)
        
        # 배치 버튼
        self.arrange_btn = ttk.Button(button_frame, text="학생 배치", command=self.arrange_seats)
        self.arrange_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 저장 버튼
        self.save_btn = ttk.Button(button_frame, text="설정 저장", command=self.save_settings)
        self.save_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 편집 모드 프레임
        mode_frame = tk.LabelFrame(self.settings_frame, text="편집 모드", font=("맑은 고딕", 10), 
                                 bg=self.colors["bg"], fg="#555555", padx=10, pady=10)
        mode_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 편집 모드 라디오 버튼
        self.mode_var = tk.StringVar(value="swap")
        
        # 자리 교환 모드
        self.swap_radio = ttk.Radiobutton(mode_frame, text="자리 교환", variable=self.mode_var, 
                                        value="swap", command=self.update_edit_mode)
        self.swap_radio.pack(anchor="w", pady=(0, 5))
        
        # 앞자리 영역 설정 모드
        self.front_radio = ttk.Radiobutton(mode_frame, text="앞자리 영역 설정", variable=self.mode_var, 
                                         value="front_area", command=self.update_edit_mode)
        self.front_radio.pack(anchor="w", pady=(0, 5))
        
        # 뒷자리 영역 설정 모드
        self.back_radio = ttk.Radiobutton(mode_frame, text="뒷자리 영역 설정", variable=self.mode_var, 
                                        value="back_area", command=self.update_edit_mode)
        self.back_radio.pack(anchor="w", pady=(0, 5))
        
        # 일반석 영역 설정 모드
        self.normal_radio = ttk.Radiobutton(mode_frame, text="일반석 영역 설정", variable=self.mode_var,
                                          value="normal_area", command=self.update_edit_mode)
        self.normal_radio.pack(anchor="w", pady=(0, 5))
        
        # 고정석 설정 모드
        self.fixed_radio = ttk.Radiobutton(mode_frame, text="고정석 설정", variable=self.mode_var, 
                                         value="fixed", command=self.update_edit_mode)
        self.fixed_radio.pack(anchor="w", pady=(0, 5))
        
        # 자리 비활성화 모드
        self.disabled_radio = ttk.Radiobutton(mode_frame, text="자리 비활성화", variable=self.mode_var, 
                                            value="disabled", command=self.update_edit_mode)
        self.disabled_radio.pack(anchor="w")
        
        # 우측 자리 배치 프레임
        self.seat_container = tk.Frame(self.root, bg=self.colors["bg"], padx=20, pady=20)
        self.seat_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 자리 배치 안내 레이블
        self.guide_label = tk.Label(self.seat_container, 
                                  text="행과 열을 입력하면 자리 배치가 표시됩니다.",
                                  font=("맑은 고딕", 12), 
                                  bg=self.colors["bg"], 
                                  fg="#555555")
        self.guide_label.pack(pady=50)
        
        # 고정석 설정 안내 프레임
        fixed_info_frame = tk.Frame(self.seat_container, bg=self.colors["bg"], pady=10)
        fixed_info_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 고정석 설명
        info_frame = tk.Frame(fixed_info_frame, bg=self.colors["bg"])
        info_frame.pack(side=tk.LEFT, fill=tk.Y, anchor="w")
        
        # 앞자리 영역 설명
        area_frame = tk.Frame(info_frame, bg=self.colors["bg"], pady=3)
        area_frame.pack(anchor="w", fill=tk.X)
        
        area_color = tk.Frame(area_frame, width=15, height=15, bg=self.colors["front_area"],
                            highlightthickness=1, highlightbackground=self.colors["front_area_border"])
        area_color.pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(area_frame, text="앞자리 영역", font=("맑은 고딕", 9),
                bg=self.colors["bg"]).pack(side=tk.LEFT, padx=(0, 15))
        
        # 뒷자리 영역 설명
        area_frame = tk.Frame(info_frame, bg=self.colors["bg"], pady=3)
        area_frame.pack(anchor="w", fill=tk.X)
        
        area_color = tk.Frame(area_frame, width=15, height=15, bg=self.colors["back_area"],
                            highlightthickness=1, highlightbackground=self.colors["back_area_border"])
        area_color.pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(area_frame, text="뒷자리 영역", font=("맑은 고딕", 9),
                bg=self.colors["bg"]).pack(side=tk.LEFT, padx=(0, 15))
        
        # 일반석 영역 설명
        area_frame = tk.Frame(info_frame, bg=self.colors["bg"], pady=3)
        area_frame.pack(anchor="w", fill=tk.X)
        
        area_color = tk.Frame(area_frame, width=15, height=15, bg=self.colors["normal_area"],
                            highlightthickness=1, highlightbackground=self.colors["normal_area_border"])
        area_color.pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(area_frame, text="일반석 영역", font=("맑은 고딕", 9),
                bg=self.colors["bg"]).pack(side=tk.LEFT, padx=(0, 15))
        
        # 현재 모드에 맞는 안내 텍스트 표시
        self.mode_help_label = tk.Label(fixed_info_frame, 
                                      text="* 자리 교환 모드: 두 자리를 차례로 클릭하여 교환", 
                                      font=("맑은 고딕", 8), 
                                      bg=self.colors["bg"], 
                                      fg="#777777")
        self.mode_help_label.pack(side=tk.RIGHT)

        # GUI 크기 자동 조절
        self.root.update_idletasks()
        required_width = self.settings_frame.winfo_reqwidth() + self.seat_container.winfo_reqwidth() + 40
        required_height = max(self.settings_frame.winfo_reqheight(), self.seat_container.winfo_reqheight()) + 40
        # 너무 작지 않도록 최소값 유지
        required_width = max(required_width, 1000)
        required_height = max(required_height, 700)
        self.root.geometry(f"{required_width}x{required_height}")
    
    def add_student(self):
        """학생 추가"""
        name = self.student_name_var.get().strip()
        position = self.position_var.get()
        
        if not name:
            messagebox.showerror("오류", "학생 이름을 입력해주세요.")
            return
        
        # 학생 목록에 추가
        self.students.append({"name": name, "position": position if position != "normal" else None})
        
        # 트리뷰에 추가
        position_text = "일반" if position == "normal" else ("앞자리" if position == "front" else "뒷자리")
        self.student_tree.insert("", "end", text=name, values=(position_text,))
        
        # 입력 필드 초기화
        self.student_name_var.set("")
        self.position_var.set("normal")
    
    def import_students_from_excel(self):
        """엑셀 파일에서 1열의 학생 이름을 읽어서 추가"""
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[
                ("엑셀 파일", "*.xlsx *.xls"),
                ("모든 파일", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        names = []
        try:
            # openpyxl 라이브러리 사용 시도
            try:
                import openpyxl
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                ws = wb.active
                
                # 1열(A열)의 모든 셀 읽기
                for row in ws.iter_rows(min_row=1, min_col=1, max_col=1, values_only=True):
                    value = row[0]
                    if value is not None:
                        name = str(value).strip()
                        if name:  # 빈 문자열이 아닌 경우만
                            names.append(name)
                
                wb.close()
            except ImportError:
                # openpyxl이 없으면 xlrd 시도 (구버전 엑셀 파일용)
                try:
                    import xlrd
                    wb = xlrd.open_workbook(file_path)
                    ws = wb.sheet_by_index(0)
                    
                    # 1열의 모든 행 읽기
                    for row_idx in range(ws.nrows):
                        value = ws.cell_value(row_idx, 0)
                        if value:
                            name = str(value).strip()
                            if name:
                                names.append(name)
                except ImportError:
                    messagebox.showerror(
                        "오류", 
                        "엑셀 파일을 읽으려면 openpyxl 라이브러리가 필요합니다.\n"
                        "명령프롬프트에서 다음 명령어로 설치해주세요:\n"
                        "pip install openpyxl"
                    )
                    return
        except Exception as e:
            messagebox.showerror("오류", f"파일 읽기 중 오류가 발생했습니다.\n{str(e)}")
            return
        
        if not names:
            messagebox.showinfo("알림", "엑셀 파일에서 이름을 찾을 수 없습니다.")
            return
        
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        # 기존 학생 이름 목록
        existing_names = {s["name"] for s in self.students}
        
        # 현재 선택된 위치 설정 가져오기
        position = self.position_var.get()
        
        # 학생 추가
        added_count = 0
        for name in names:
            if name not in existing_names:
                self.students.append({"name": name, "position": position if position != "normal" else None})
                existing_names.add(name)
                
                # 트리뷰에 추가
                position_text = "일반" if position == "normal" else ("앞자리" if position == "front" else "뒷자리")
                self.student_tree.insert("", "end", text=name, values=(position_text,))
                added_count += 1
        
        if added_count > 0:
            messagebox.showinfo("완료", f"{added_count}명의 학생이 추가되었습니다.")
        else:
            messagebox.showinfo("알림", "추가할 학생이 없습니다. (모든 이름이 이미 목록에 있음)")
    
    def bulk_add_students(self):
        """학생 일괄 추가 (1번부터 지정된 숫자까지)"""
        try:
            count = int(self.bulk_count_var.get())
            if count <= 0:
                messagebox.showerror("오류", "개수는 1 이상의 숫자여야 합니다.")
                return
            if count > 100:
                messagebox.showerror("오류", "한 번에 추가할 수 있는 최대 개수는 100명입니다.")
                return
        except ValueError:
            messagebox.showerror("오류", "개수에 숫자를 입력해주세요.")
            return
        
        # 현재 선택된 위치 설정 가져오기
        position = self.position_var.get()
        
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        # 기존 학생 이름 목록 가져오기
        existing_names = {s["name"] for s in self.students}
        
        # 1번부터 count번까지 추가
        added_count = 0
        for i in range(1, count + 1):
            student_name = f"{i}번"
            
            # 이미 존재하는 이름이면 건너뛰기
            if student_name in existing_names:
                continue
            
            # 학생 목록에 추가
            self.students.append({"name": student_name, "position": position if position != "normal" else None})
            existing_names.add(student_name)
            
            # 트리뷰에 추가
            position_text = "일반" if position == "normal" else ("앞자리" if position == "front" else "뒷자리")
            self.student_tree.insert("", "end", text=student_name, values=(position_text,))
            added_count += 1
        
        if added_count > 0:
            messagebox.showinfo("완료", f"{added_count}명의 학생이 추가되었습니다.")
        else:
            messagebox.showinfo("알림", "추가할 학생이 없습니다. (이미 모두 추가되어 있음)")
    
    def delete_student(self):
        """선택한 학생 삭제"""
        selected_items = self.student_tree.selection()
        if not selected_items:
            messagebox.showerror("오류", "삭제할 학생을 선택해주세요.")
            return
        
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        # 선택된 모든 학생 이름 수집
        selected_names = [self.student_tree.item(item_id, "text") for item_id in selected_items]
        name_set = set(selected_names)
        
        # 학생 목록에서 제거
        self.students = [s for s in self.students if s["name"] not in name_set]
        
        # 좌석에서 해당 학생 제거 및 버튼 텍스트 갱신
        if self.seats:
            for r in range(len(self.seats)):
                for c in range(len(self.seats[r])):
                    if self.seats[r][c] in name_set:
                        self.seats[r][c] = ""
                        if self.seat_buttons and r < len(self.seat_buttons) and c < len(self.seat_buttons[r]):
                            self.seat_buttons[r][c].config(text="")
                            # 비활성화 자리가 아니라면 기본색으로 복원
                            if (r, c) not in self.disabled_seats:
                                self.update_seat_color(r, c)
        
        # 트리뷰에서 제거
        for item_id in selected_items:
            self.student_tree.delete(item_id)
    
    def change_student_position(self, position):
        """선택한 학생의 위치 변경"""
        selected_items = self.student_tree.selection()
        if not selected_items:
            messagebox.showerror("오류", "위치를 변경할 학생을 선택해주세요.")
            return
        
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        # 선택된 모든 항목 처리
        position_text = "일반" if position == "normal" else ("앞자리" if position == "front" else "뒷자리")
        selected_names = []
        for item_id in selected_items:
            name = self.student_tree.item(item_id, "text")
            selected_names.append(name)
            # 트리뷰 업데이트
            self.student_tree.item(item_id, values=(position_text,))
        
        # 학생 목록에서 위치 변경
        new_pos = None if position == "normal" else position
        name_set = set(selected_names)
        for student in self.students:
            if student["name"] in name_set:
                student["position"] = new_pos
        
        # 좌석 색상 갱신 (이미 배치된 경우 시각적 반영)
        if self.seats and self.seat_buttons:
            for r in range(len(self.seats)):
                for c in range(len(self.seats[r])):
                    if self.seats[r][c] in name_set:
                        self.update_seat_color(r, c)
    
    def update_student_list(self, event=None):
        """학생 명단 업데이트 (이전 메소드, 이제 사용하지 않음)"""
        pass
    
    def check_row_col_input(self, event=None):
        """행/열 입력 확인 및 자리 레이아웃 생성"""
        # 입력값 검증
        try:
            rows_str = self.row_var.get().strip()
            cols_str = self.col_var.get().strip()
            
            if not rows_str or not cols_str:
                return
            
            rows = int(rows_str)
            cols = int(cols_str)
            
            if rows <= 0 or cols <= 0:
                return
        except ValueError:
            return
        
        # 행/열이 모두 입력되었고 변경된 경우에만 업데이트
        if self.rows != rows or self.cols != cols:
            self.create_initial_layout()
    
    def create_initial_layout(self, event=None):
        """행/열 입력 시 빈 자리 레이아웃 생성"""
        # 입력값 검증
        try:
            rows = int(self.row_var.get())
            cols = int(self.col_var.get())
            if rows <= 0 or cols <= 0:
                return
        except ValueError:
            return
        
        # 행/열이 변경된 경우에만 업데이트
        if self.rows != rows or self.cols != cols:
            self.rows = rows
            self.cols = cols
            
            # 영역 초기화 (행/열이 변경되면)
            self.front_area = set()
            self.back_area = set()
            self.normal_area = set()
            self.fixed_seats = {}
            self.disabled_seats = set()
            
            # 자리 초기화 (빈 자리)
            self.seats = [["" for _ in range(self.cols)] for _ in range(self.rows)]
            
            # 자리 레이아웃 생성
            self.create_seat_layout()
    
    def ensure_student_dict(self, student):
        """학생 데이터가 딕셔너리 형태인지 확인하고 변환"""
        if isinstance(student, str):
            return {"name": student, "position": None}
        return student
    
    def create_seat_layout(self):
        """자리 레이아웃 생성"""
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        # 기존 자리 제거
        for widget in self.seat_container.winfo_children():
            widget.destroy()
        
        # 안내 레이블 다시 추가
        self.guide_label = tk.Label(self.seat_container, 
                                  text=f"{self.rows}행 {self.cols}열 자리 배치",
                                  font=("맑은 고딕", 12), 
                                  bg=self.colors["bg"], 
                                  fg="#555555")
        self.guide_label.pack(pady=(0, 20))
        
        # 좌석 프레임
        seat_frame = tk.Frame(self.seat_container, bg=self.colors["bg"], padx=10, pady=10)
        seat_frame.pack(expand=True)
        
        # 교탁 추가
        teacher_desk = tk.Label(seat_frame, text="교탁", width=10, height=2,
                              bg=self.colors["teacher"], fg="#333333",
                              font=("맑은 고딕", 10, "bold"),
                              relief=tk.RAISED)
        teacher_desk.grid(row=0, column=0, columnspan=self.cols, padx=2, pady=(0, 20))
        
        # 좌석 버튼 생성
        self.seat_buttons = []
        self.seat_frames = []  # 자리 프레임 저장
        for r in range(self.rows):
            row_buttons = []
            row_frames = []
            for c in range(self.cols):
                # 좌석에 표시할 학생 이름
                student_name = self.seats[r][c] if r < len(self.seats) and c < len(self.seats[r]) else ""
                
                # 좌석 상태에 따라 배경색 결정
                bg_color = self.colors["seat"]
                
                # 영역 배경색 및 테두리 결정
                frame_bg = self.colors["normal_area"]  # 기본 배경색은 일반석 영역
                frame_border = self.colors["normal_area_border"]  # 기본 테두리는 일반석 테두리
                
                if (r, c) in self.front_area:
                    frame_bg = self.colors["front_area"]
                    frame_border = self.colors["front_area_border"]
                elif (r, c) in self.back_area:
                    frame_bg = self.colors["back_area"]
                    frame_border = self.colors["back_area_border"]
                
                # 비활성화된 자리 확인
                if (r, c) in self.disabled_seats:
                    bg_color = self.colors["disabled_seat"]
                # 고정석만 색상 적용
                elif student_name and (r, c) in self.fixed_seats:
                    if (r, c) in self.front_area:
                        bg_color = self.colors["front_fixed"]
                    elif (r, c) in self.back_area:
                        bg_color = self.colors["back_fixed"]
                    else:
                        # 일반석 영역 고정석
                        bg_color = self.colors["normal_fixed"]
                
                # 자리 프레임 생성 (영역 표시용)
                seat_frame_cell = tk.Frame(seat_frame, bg=frame_bg, padx=2, pady=2,
                                        highlightthickness=1, highlightbackground=frame_border)
                seat_frame_cell.grid(row=r+1, column=c, padx=3, pady=3)
                row_frames.append(seat_frame_cell)
                
                # 좌석 버튼 생성
                seat_btn = tk.Button(seat_frame_cell, text=student_name, width=10, height=2,
                                   bg=bg_color, fg="#333333",
                                   font=("맑은 고딕", 9),
                                   relief=tk.RAISED,
                                   command=lambda r=r, c=c: self.on_seat_click(r, c))
                seat_btn.pack(padx=0, pady=0)
                row_buttons.append(seat_btn)
            self.seat_buttons.append(row_buttons)
            self.seat_frames.append(row_frames)
        
        # 고정석 설정 안내 프레임
        fixed_info_frame = tk.Frame(self.seat_container, bg=self.colors["bg"], pady=10)
        fixed_info_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 고정석 설명
        info_frame = tk.Frame(fixed_info_frame, bg=self.colors["bg"])
        info_frame.pack(side=tk.LEFT, fill=tk.Y, anchor="w")
        
        # 앞자리 영역 설명
        area_frame = tk.Frame(info_frame, bg=self.colors["bg"], pady=3)
        area_frame.pack(anchor="w", fill=tk.X)
        
        area_color = tk.Frame(area_frame, width=15, height=15, bg=self.colors["front_area"],
                            highlightthickness=1, highlightbackground=self.colors["front_area_border"])
        area_color.pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(area_frame, text="앞자리 영역", font=("맑은 고딕", 9),
                bg=self.colors["bg"]).pack(side=tk.LEFT, padx=(0, 15))
        
        # 뒷자리 영역 설명
        area_frame = tk.Frame(info_frame, bg=self.colors["bg"], pady=3)
        area_frame.pack(anchor="w", fill=tk.X)
        
        area_color = tk.Frame(area_frame, width=15, height=15, bg=self.colors["back_area"],
                            highlightthickness=1, highlightbackground=self.colors["back_area_border"])
        area_color.pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(area_frame, text="뒷자리 영역", font=("맑은 고딕", 9),
                bg=self.colors["bg"]).pack(side=tk.LEFT, padx=(0, 15))
        
        # 일반석 영역 설명
        area_frame = tk.Frame(info_frame, bg=self.colors["bg"], pady=3)
        area_frame.pack(anchor="w", fill=tk.X)
        
        area_color = tk.Frame(area_frame, width=15, height=15, bg=self.colors["normal_area"],
                            highlightthickness=1, highlightbackground=self.colors["normal_area_border"])
        area_color.pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(area_frame, text="일반석 영역", font=("맑은 고딕", 9),
                bg=self.colors["bg"]).pack(side=tk.LEFT, padx=(0, 15))
        
        # 현재 모드에 맞는 안내 텍스트 표시
        self.mode_help_label = tk.Label(fixed_info_frame, 
                                      text="* 자리 교환 모드: 두 자리를 차례로 클릭하여 교환", 
                                      font=("맑은 고딕", 8), 
                                      bg=self.colors["bg"], 
                                      fg="#777777")
        self.mode_help_label.pack(side=tk.RIGHT)

        # GUI 크기 자동 조절
        self.root.update_idletasks()
        required_width = self.settings_frame.winfo_reqwidth() + self.seat_container.winfo_reqwidth() + 40
        required_height = max(self.settings_frame.winfo_reqheight(), self.seat_container.winfo_reqheight()) + 40
        # 너무 작지 않도록 최소값 유지
        required_width = max(required_width, 1000)
        required_height = max(required_height, 700)
        self.root.geometry(f"{required_width}x{required_height}")
    
    def arrange_seats(self):
        """학생 자리 배정"""
        # 입력값 검증
        try:
            self.rows = int(self.row_var.get())
            self.cols = int(self.col_var.get())
            if self.rows <= 0 or self.cols <= 0:
                messagebox.showerror("오류", "행과 열은 양수여야 합니다.")
                return
        except ValueError:
            messagebox.showerror("오류", "행과 열에 숫자를 입력해주세요.")
            return
        
        if not self.students:
            messagebox.showerror("오류", "학생 명단을 입력해주세요.")
            return
        
        # 영역이 설정되지 않은 좌표를 모두 일반석으로 설정
        all_coordinates = {(r, c) for r in range(self.rows) for c in range(self.cols)}
        
        # 앞/뒤 영역이 설정되지 않은 경우 기본값 설정
        if not self.front_area and not self.back_area and not self.normal_area:
            # 전체 좌석 수의 1/3을 앞자리로, 1/3을 뒷자리로 설정
            total_seats = self.rows * self.cols
            front_part = self.rows // 3
            back_part = self.rows - front_part
            
            # 기본 앞자리 영역: 앞쪽 1/3
            self.front_area = {(r, c) for r in range(front_part) for c in range(self.cols)}
            # 기본 뒷자리 영역: 뒤쪽 1/3
            self.back_area = {(r, c) for r in range(back_part, self.rows) for c in range(self.cols)}
            # 기본 일반석 영역: 중간 1/3
            self.normal_area = {(r, c) for r in range(front_part, back_part) for c in range(self.cols)}
            
            messagebox.showinfo("알림", "자리 영역이 설정되지 않아 기본값으로 설정되었습니다.\n앞자리: 앞쪽 1/3, 뒷자리: 뒤쪽 1/3, 일반석: 중간 1/3")
        else:
            # 영역이 하나도 설정되지 않은 좌표는, 일반석 영역으로 자동 설정
            unassigned = all_coordinates - (self.front_area | self.back_area | self.normal_area)
            if unassigned:
                self.normal_area |= unassigned
        
        # 앞/뒤/일반석 자리 수 계산
        front_seats_count = len(self.front_area)
        back_seats_count = len(self.back_area)
        normal_seats_count = len(self.normal_area)
        
        # 앞자리/뒷자리 학생 수 계산
        front_students = [s for s in self.students if s["position"] == "front"]
        back_students = [s for s in self.students if s["position"] == "back"]
        normal_students = [s for s in self.students if s["position"] is None or s["position"] == "normal"]
        
        # 자리 수와 학생 수 비교
        if len(front_students) > front_seats_count:
            messagebox.showwarning("경고", f"앞자리로 지정된 학생({len(front_students)}명)이 앞쪽 자리 수({front_seats_count}개)보다 많습니다.\n일부 학생은 다른 자리에 배정될 수 있습니다.")
        
        if len(back_students) > back_seats_count:
            messagebox.showwarning("경고", f"뒷자리로 지정된 학생({len(back_students)}명)이 뒤쪽 자리 수({back_seats_count}개)보다 많습니다.\n일부 학생은 다른 자리에 배정될 수 있습니다.")
        
        # 고정석 정보 백업 (기존 자리 배치에서)
        old_fixed_seats = self.fixed_seats.copy()
        
        # 자리 초기화
        self.seats = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        
        # 고정석 정보 복원 (유효한 좌표만)
        self.fixed_seats = {}
        for (r, c), student_name in old_fixed_seats.items():
            if 0 <= r < self.rows and 0 <= c < self.cols:
                self.fixed_seats[(r, c)] = student_name
                self.seats[r][c] = student_name
        
        # 학생 수와 자리 수 비교
        total_seats = self.rows * self.cols
        if len(self.students) > total_seats:
            messagebox.showwarning("경고", f"학생 수({len(self.students)}명)가 자리 수({total_seats}개)보다 많습니다.\n앞에서부터 {total_seats}명만 배정됩니다.")
            self.students = self.students[:total_seats]
        
        # 학생 배치
        self.assign_students()
        
        # 자리 레이아웃 생성
        self.create_seat_layout()
    
    def assign_students(self):
        """학생들을 자리에 배정"""
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        # 고정석이 이미 배정되어 있으므로, 고정석에 배정된 학생 목록 가져오기
        fixed_student_names = set(self.fixed_seats.values())
        
        # 앞쪽/뒤쪽/일반석 학생 분류 (고정석 제외)
        front_students = [s for s in self.students if s["position"] == "front" and s["name"] not in fixed_student_names]
        back_students = [s for s in self.students if s["position"] == "back" and s["name"] not in fixed_student_names]
        normal_students = [s for s in self.students if (s["position"] is None or s["position"] == "normal") and s["name"] not in fixed_student_names]
        
        # 각 그룹 내에서 섞기
        random.shuffle(front_students)
        random.shuffle(back_students)
        random.shuffle(normal_students)
        
        # 앞쪽 자리 좌표 리스트 (고정석 및 비활성화 자리 제외)
        front_positions = [pos for pos in self.front_area if pos not in self.fixed_seats and pos not in self.disabled_seats]
        random.shuffle(front_positions)
        
        # 뒤쪽 자리 좌표 리스트 (고정석 및 비활성화 자리 제외)
        back_positions = [pos for pos in self.back_area if pos not in self.fixed_seats and pos not in self.disabled_seats]
        random.shuffle(back_positions)
        
        # 일반석 자리 좌표 리스트 (고정석 및 비활성화 자리 제외)
        normal_positions = [pos for pos in self.normal_area if pos not in self.fixed_seats and pos not in self.disabled_seats]
        random.shuffle(normal_positions)
        
        # 고정석은 이미 배정되어 있으므로 assigned_positions에 추가
        assigned_positions = list(self.fixed_seats.keys())
        
        # 앞쪽 고정석 학생 배정 (고정석이 아닌 학생만)
        for i, student in enumerate(front_students):
            if i < len(front_positions):
                r, c = front_positions[i]
                self.seats[r][c] = student["name"]
                assigned_positions.append((r, c))
        
        # 뒤쪽 고정석 학생 배정 (고정석이 아닌 학생만)
        for i, student in enumerate(back_students):
            if i < len(back_positions):
                r, c = back_positions[i]
                self.seats[r][c] = student["name"]
                assigned_positions.append((r, c))
        
        # 남은 앞쪽 좌표 계산
        remaining_front = [pos for pos in front_positions if pos not in assigned_positions]
        
        # 남은 뒤쪽 좌표 계산
        remaining_back = [pos for pos in back_positions if pos not in assigned_positions]
        
        # 남은 일반석 좌표 계산
        remaining_normal = [pos for pos in normal_positions if pos not in assigned_positions]
        
        # 앞쪽 고정석 학생이 앞쪽 자리보다 많은 경우, 일반석으로 배정
        overflow_front = front_students[len(front_positions):] if len(front_students) > len(front_positions) else []
        
        # 뒤쪽 고정석 학생이 뒤쪽 자리보다 많은 경우, 일반석으로 배정
        overflow_back = back_students[len(back_positions):] if len(back_students) > len(back_positions) else []
        
        # 남은 모든 좌표 합치기 - 일반석을 우선 사용, 그 다음에 앞/뒤 자리 (비활성화 자리 제외)
        remaining_all = [pos for pos in remaining_normal + remaining_front + remaining_back if pos not in self.disabled_seats]
        random.shuffle(remaining_all)
        
        # 남은 좌표가 없으면 배정 불가
        if not remaining_all and (normal_students or overflow_front or overflow_back):
            messagebox.showwarning("경고", "모든 자리가 고정석으로 지정되어 일반 학생을 배정할 수 없습니다.")
            return
            
        # 고정석에서 넘친 학생들을 일반석에 배정
        overflow_students = overflow_front + overflow_back
        random.shuffle(overflow_students)
        
        for i, student in enumerate(overflow_students):
            if i < len(remaining_all):
                r, c = remaining_all[i]
                self.seats[r][c] = student["name"]
                remaining_all.remove((r, c))  # 배정된 좌표 제거
        
        # 일반 학생 배정
        for i, student in enumerate(normal_students):
            if i < len(remaining_all):
                r, c = remaining_all[i]
                self.seats[r][c] = student["name"]
                
    def update_edit_mode(self):
        """편집 모드 업데이트"""
        self.edit_mode = self.mode_var.get()
        # 선택된 자리 초기화
        self.selected_seats = []
        
        # 모드에 따른 도움말 업데이트
        if self.edit_mode == "swap":
            help_text = "* 자리 교환 모드: 두 자리를 차례로 클릭하여 교환"
        elif self.edit_mode == "front_area":
            help_text = "* 앞자리 영역 설정 모드: 앞자리 영역으로 지정할 자리를 클릭"
        elif self.edit_mode == "back_area":
            help_text = "* 뒷자리 영역 설정 모드: 뒷자리 영역으로 지정할 자리를 클릭"
        elif self.edit_mode == "normal_area":
            help_text = "* 일반석 영역 설정 모드: 일반석 영역으로 지정할 자리를 클릭"
        elif self.edit_mode == "fixed":
            help_text = "* 고정석 설정 모드: 자리를 클릭하여 고정석 설정/해제"
        else:  # disabled
            help_text = "* 자리 비활성화 모드: 자리를 클릭하여 배정에서 제외/포함"
        
        self.mode_help_label.config(text=help_text)
        
        # 자리 레이아웃이 있는 경우 모든 자리 색상 복원
        if hasattr(self, 'seat_buttons') and self.seat_buttons:
            for r in range(len(self.seat_buttons)):
                for c in range(len(self.seat_buttons[r])):
                    self.update_seat_color(r, c)
    
    def on_seat_click(self, row, col):
        """좌석 클릭 이벤트 처리"""
        # 자리 교환 모드
        if self.edit_mode == "swap":
            self.handle_swap_mode(row, col)
        # 앞자리 영역 설정 모드
        elif self.edit_mode == "front_area":
            self.handle_area_mode(row, col, "front")
        # 뒷자리 영역 설정 모드
        elif self.edit_mode == "back_area":
            self.handle_area_mode(row, col, "back")
        # 일반석 영역 설정 모드
        elif self.edit_mode == "normal_area":
            self.handle_area_mode(row, col, "normal")
        # 고정석 설정 모드
        elif self.edit_mode == "fixed":
            self.handle_fixed_mode(row, col)
        # 자리 비활성화 모드
        else:
            self.handle_disabled_mode(row, col)
    
    def handle_area_mode(self, row, col, area_type):
        """영역 설정 모드 처리"""
        if not (self.seat_buttons and row < len(self.seat_buttons) and col < len(self.seat_buttons[row])):
            return
        
        # 앞/뒤/일반 영역 참조 설정
        if area_type == "front":
            area_set = self.front_area
            other_area_sets = [self.back_area, self.normal_area]
            area_color = self.colors["front_area"]
            border_color = self.colors["front_area_border"]
        elif area_type == "back":
            area_set = self.back_area
            other_area_sets = [self.front_area, self.normal_area]
            area_color = self.colors["back_area"]
            border_color = self.colors["back_area_border"]
        else:  # normal
            area_set = self.normal_area
            other_area_sets = [self.front_area, self.back_area]
            area_color = self.colors["normal_area"]
            border_color = self.colors["normal_area_border"]
        
        # 좌표
        pos = (row, col)
        
        # 다른 영역에 이미 포함되어 있는지 확인
        for other_set in other_area_sets:
            if pos in other_set:
                other_set.remove(pos)
        
        # 토글 처리
        if pos in area_set:
            area_set.remove(pos)
            # 배경색 초기화 - 일반석 영역으로 전환
            if area_type != "normal":  # 일반석 모드가 아닌 경우에만
                self.normal_area.add(pos)
                self.seat_frames[row][col].config(bg=self.colors["normal_area"], 
                                               highlightbackground=self.colors["normal_area_border"])
            else:
                # 일반석이 제거되면 배경색을 기본 배경색으로
                self.seat_frames[row][col].config(bg=self.colors["bg"], 
                                               highlightbackground=self.colors["bg"])
        else:
            area_set.add(pos)
            # 배경색 변경 (영역 색상)
            self.seat_frames[row][col].config(bg=area_color, highlightbackground=border_color)
    
    def handle_swap_mode(self, row, col):
        """자리 교환 모드 처리"""
        # 선택된 자리가 없으면 첫 번째 선택으로 추가
        if not self.selected_seats:
            self.selected_seats.append((row, col))
            self.seat_buttons[row][col].config(bg=self.colors["seat_selected"])
            
        # 이미 선택된 자리가 있고, 다른 자리를 선택한 경우 자리 교환
        elif (row, col) != self.selected_seats[0]:
            self.selected_seats.append((row, col))
            self.swap_seats()
            
            # 선택 초기화
            for r, c in self.selected_seats:
                # 좌석 색상 복원
                self.update_seat_color(r, c)
            
            self.selected_seats = []
            
        # 같은 자리를 다시 클릭한 경우, 선택 취소
        else:
            # 좌석 색상 복원
            self.update_seat_color(row, col)
            self.selected_seats = []
    
    def handle_fixed_mode(self, row, col):
        """고정석 설정 모드 처리"""
        # 자리에 학생이 있는 경우에만 처리
        if not self.seats or row >= len(self.seats) or col >= len(self.seats[0]):
            return
            
        student_name = self.seats[row][col] if self.seats[row][col] else ""
        if not student_name:
            return
        
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        pos = (row, col)
        
        # 이미 고정석인지 확인
        if pos in self.fixed_seats and self.fixed_seats[pos] == student_name:
            # 고정석 해제
            del self.fixed_seats[pos]
            self.update_seat_color(row, col)
        else:
            # 기존에 이 자리에 다른 학생이 고정되어 있으면 해제
            if pos in self.fixed_seats:
                del self.fixed_seats[pos]
            
            # 해당 학생이 다른 자리에 고정되어 있으면 해제
            fixed_positions = [pos for pos, name in self.fixed_seats.items() if name == student_name]
            for old_pos in fixed_positions:
                del self.fixed_seats[old_pos]
            
            # 고정석 설정
            self.fixed_seats[pos] = student_name
            
            # 좌표가 어느 영역에 속하는지 확인하여 색상 설정
            is_front_area = (row, col) in self.front_area
            is_back_area = (row, col) in self.back_area
            
            if is_front_area:
                self.seat_buttons[row][col].config(bg=self.colors["front_fixed"])
            elif is_back_area:
                self.seat_buttons[row][col].config(bg=self.colors["back_fixed"])
            else:
                # 일반석 영역 고정석
                self.seat_buttons[row][col].config(bg=self.colors["normal_fixed"])
        
        # 트리뷰 업데이트
        self.update_student_tree()
    
    def handle_disabled_mode(self, row, col):
        """자리 비활성화 모드 처리"""
        if not (self.seat_buttons and row < len(self.seat_buttons) and col < len(self.seat_buttons[row])):
            return
        
        pos = (row, col)
        
        # 비활성화 토글
        if pos in self.disabled_seats:
            # 비활성화 해제
            self.disabled_seats.remove(pos)
            self.update_seat_color(row, col)
        else:
            # 비활성화 설정
            self.disabled_seats.add(pos)
            # 비활성화된 자리는 회색으로 표시
            self.seat_buttons[row][col].config(bg=self.colors["disabled_seat"])
            # 자리에 학생이 있으면 제거
            if self.seats and row < len(self.seats) and col < len(self.seats[row]):
                self.seats[row][col] = ""
                self.seat_buttons[row][col].config(text="")
    
    def update_student_tree(self):
        """학생 트리뷰 업데이트"""
        # 학생 데이터 구조 검증
        self.students = [self.ensure_student_dict(s) for s in self.students]
        
        # 트리뷰 업데이트
        for item in self.student_tree.get_children():
            item_text = self.student_tree.item(item, "text")
            # 해당 학생 찾기
            for student in self.students:
                if student["name"] == item_text:
                    position_text = "일반"
                    if student["position"] == "front":
                        position_text = "앞자리"
                    elif student["position"] == "back":
                        position_text = "뒷자리"
                    # 트리뷰 아이템 업데이트
                    self.student_tree.item(item, values=(position_text,))
                    break
    
    def swap_seats(self):
        """선택된 두 자리의 학생 교환"""
        if len(self.selected_seats) != 2:
            return
        
        r1, c1 = self.selected_seats[0]
        r2, c2 = self.selected_seats[1]
        
        # 학생 이름 교환
        self.seats[r1][c1], self.seats[r2][c2] = self.seats[r2][c2], self.seats[r1][c1]
        
        # 학생 위치 속성 업데이트
        student1_name = self.seats[r1][c1]
        student2_name = self.seats[r2][c2]
        
        # 앞쪽/뒤쪽 자리 계산
        front_half = self.rows // 2
        
        # 학생1 업데이트
        if student1_name:
            for student in self.students:
                if student["name"] == student1_name:
                    # 앞쪽/뒤쪽 영역에 따라 위치 속성 업데이트
                    if student["position"] == "front" and r1 >= front_half:
                        student["position"] = None  # 앞자리 학생이 뒷영역으로 갔을 때
                    elif student["position"] == "back" and r1 < front_half:
                        student["position"] = None  # 뒷자리 학생이 앞영역으로 갔을 때
                    break
        
        # 학생2 업데이트
        if student2_name:
            for student in self.students:
                if student["name"] == student2_name:
                    # 앞쪽/뒤쪽 영역에 따라 위치 속성 업데이트
                    if student["position"] == "front" and r2 >= front_half:
                        student["position"] = None  # 앞자리 학생이 뒷영역으로 갔을 때
                    elif student["position"] == "back" and r2 < front_half:
                        student["position"] = None  # 뒷자리 학생이 앞영역으로 갔을 때
                    break
        
        # 트리뷰 업데이트
        self.update_student_tree()
        
        # 버튼 텍스트 업데이트
        self.seat_buttons[r1][c1].config(text=self.seats[r1][c1])
        self.seat_buttons[r2][c2].config(text=self.seats[r2][c2])
    
    def save_settings(self):
        """현재 설정 저장"""
        # 저장할 데이터
        data = {
            "rows": self.rows,
            "cols": self.cols,
            "students": self.students,
            "seats": self.seats,
            "front_area": list(self.front_area),
            "back_area": list(self.back_area),
            "normal_area": list(self.normal_area),
            "fixed_seats": {f"{r},{c}": name for (r, c), name in self.fixed_seats.items()},
            "disabled_seats": list(self.disabled_seats)
        }
        
        try:
            with open("seat_settings.json", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"설정 저장 중 오류가 발생했습니다.\n{str(e)}")
    
    def load_settings(self):
        """저장된 설정 불러오기"""
        try:
            if os.path.exists("seat_settings.json"):
                with open("seat_settings.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # 데이터 로드
                self.rows = data.get("rows", 0)
                self.cols = data.get("cols", 0)
                self.students = data.get("students", [])
                self.seats = data.get("seats", [])
                
                # 앞/뒤/일반석 영역 설정 로드
                self.front_area = set(tuple(pos) for pos in data.get("front_area", []))
                self.back_area = set(tuple(pos) for pos in data.get("back_area", []))
                self.normal_area = set(tuple(pos) for pos in data.get("normal_area", []))
                
                # 고정석 정보 로드
                fixed_seats_data = data.get("fixed_seats", {})
                self.fixed_seats = {}
                for pos_str, student_name in fixed_seats_data.items():
                    try:
                        r, c = map(int, pos_str.split(","))
                        self.fixed_seats[(r, c)] = student_name
                    except:
                        pass
                
                # 비활성화된 자리 정보 로드
                self.disabled_seats = set(tuple(pos) for pos in data.get("disabled_seats", []))
                
                # 학생 데이터 구조 검증 및 변환 (문자열인 경우 딕셔너리로 변환)
                for i in range(len(self.students)):
                    if isinstance(self.students[i], str):
                        self.students[i] = {"name": self.students[i], "position": None}
                
                # UI 업데이트
                self.row_var.set(str(self.rows))
                self.col_var.set(str(self.cols))
                
                # 트리뷰 초기화
                for item in self.student_tree.get_children():
                    self.student_tree.delete(item)
                
                # 트리뷰에 학생 추가
                for student in self.students:
                    name = student["name"]
                    position = student["position"]
                    position_text = "일반"
                    if position == "front":
                        position_text = "앞자리"
                    elif position == "back":
                        position_text = "뒷자리"
                    
                    self.student_tree.insert("", "end", text=name, values=(position_text,))
                
                # 자리가 있으면 레이아웃 생성
                if self.rows > 0 and self.cols > 0:
                    self.create_seat_layout()
        except Exception as e:
            print(f"설정 로드 오류: {str(e)}")

    def update_seat_color(self, row, col):
        """자리 색상 업데이트"""
        if not (self.seats and row < len(self.seats) and col < len(self.seats[row]) and 
                self.seat_buttons and row < len(self.seat_buttons) and col < len(self.seat_buttons[row])):
            return
        
        # 기본 색상
        bg_color = self.colors["seat"]
        
        # 고정석 여부 확인
        pos = (row, col)
        is_fixed = pos in self.fixed_seats
        
        # 비활성화된 자리 확인
        if pos in self.disabled_seats:
            bg_color = self.colors["disabled_seat"]
        # 고정석만 색상 적용
        elif is_fixed:
            student_name = self.seats[row][col]
            if student_name:
                if (row, col) in self.front_area:
                    bg_color = self.colors["front_fixed"]
                elif (row, col) in self.back_area:
                    bg_color = self.colors["back_fixed"]
                else:
                    bg_color = self.colors["normal_fixed"]
        
        # 버튼 색상 업데이트
        self.seat_buttons[row][col].config(bg=bg_color)
        
        # 프레임 색상 업데이트 (영역 표시)
        frame_bg = self.colors["normal_area"]  # 기본 배경색은 일반석
        frame_border = self.colors["normal_area_border"]  # 기본 테두리 색상은 일반석 테두리
        
        if (row, col) in self.front_area:
            frame_bg = self.colors["front_area"]
            frame_border = self.colors["front_area_border"]
        elif (row, col) in self.back_area:
            frame_bg = self.colors["back_area"]
            frame_border = self.colors["back_area_border"]
        elif (row, col) not in self.normal_area:
            # 어느 영역에도 속하지 않는 경우
            frame_bg = self.colors["bg"]
            frame_border = self.colors["bg"]
        
        self.seat_frames[row][col].config(bg=frame_bg, highlightbackground=frame_border)

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentSeatArrangement(root)
    root.mainloop()
