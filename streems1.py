import tkinter as tk
import random

# 스트림스 카드 뽑기 프로그램 (단일 파일 버전)

class StreamsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("스트림스 카드 뽑기")
        self.root.geometry("900x650")
        self.root.minsize(800, 600)
        
        # 파스텔 톤 색상 팔레트
        self.colors = {
            "bg": "#F3F4F6",            # 전체 배경 (연한 그레이)
            "card_bg": "#FFFFFF",       # 카드 배경 (화이트)
            "card_border": "#CBD5F5",   # 카드 테두리 (연보라)
            "primary": "#A5B4FC",       # 버튼 기본색 (라벤더)
            "primary_dark": "#818CF8",  # 버튼 호버/클릭 색
            "secondary": "#FDE68A",     # 보조 버튼 (파스텔 옐로우)
            "secondary_dark": "#FACC15",
            "accent": "#F9A8D4",        # 포인트 색 (핑크)
            "table_bg": "#E5E7EB",      # 테이블 배경
            "table_cell": "#F9FAFB",    # 테이블 셀
            "text_main": "#111827",     # 메인 텍스트
            "text_sub": "#6B7280",      # 서브 텍스트
        }
        
        self.root.configure(bg=self.colors["bg"])
        
        # 폰트 설정
        self.font_title = ("맑은 고딕", 24, "bold")
        self.font_subtitle = ("맑은 고딕", 12)
        self.font_card = ("맑은 고딕", 56, "bold")
        self.font_button = ("맑은 고딕", 14, "bold")
        self.font_table_header = ("맑은 고딕", 11, "bold")
        self.font_table_cell = ("맑은 고딕", 12)
        
        # 프레임 구성
        self.start_frame = tk.Frame(self.root, bg=self.colors["bg"])
        self.game_frame = tk.Frame(self.root, bg=self.colors["bg"])
        
        self.create_start_screen()
        self.create_game_screen()
        
        # 시작 화면 보이기
        self.show_start_screen()
    
    # ---------- 카드 덱 관련 ----------
    def build_deck(self):
        """스트림스 카드 덱 생성: 1~10 한 장, 11~20 두 장, 21~30 세 장"""
        deck = []
        deck.extend(range(1, 11))             # 1~10: 1장
        deck.extend(range(11, 21))            # 11~20: 1세트
        deck.extend(range(11, 21))            # 11~20: 2세트 (총 2장)
        deck.extend(range(21, 31))            # 21~30: 1세트
        random.shuffle(deck)
        return deck
    
    def reset_game(self):
        """게임 상태 초기화"""
        self.deck = self.build_deck()
        self.drawn_cards = []
        self.current_index = 0
        self.update_card_display(None)
        self.update_counter()
        
        # 표 초기화
        for label in self.table_cells:
            label.config(text=" ", bg=self.colors["table_cell"])
        
        # 버튼 활성화
        self.draw_button.config(state="normal", bg=self.colors["primary"])
    
    # ---------- 화면 전환 ----------
    def show_start_screen(self):
        self.game_frame.pack_forget()
        self.start_frame.pack(fill="both", expand=True)
    
    def show_game_screen(self):
        self.start_frame.pack_forget()
        self.game_frame.pack(fill="both", expand=True)
        # 게임 시작 시 덱 초기화
        self.reset_game()
    
    # ---------- 시작 화면 ----------
    def create_start_screen(self):
        # 가운데 정렬 컨테이너
        container = tk.Frame(self.start_frame, bg=self.colors["bg"])
        container.place(relx=0.5, rely=0.5, anchor="center")
        
        title = tk.Label(
            container,
            text="스트림스 카드 뽑기",
            font=self.font_title,
            fg=self.colors["text_main"],
            bg=self.colors["bg"]
        )
        title.pack(pady=(0, 10))
        
        subtitle = tk.Label(
            container,
            text="1~30까지의 숫자 카드로 즐기는 스트림스 보드게임\n"
                 "시작하기 버튼을 눌러 카드를 한 장씩 뽑아보세요!",
            font=self.font_subtitle,
            fg=self.colors["text_sub"],
            bg=self.colors["bg"],
            justify="center"
        )
        subtitle.pack(pady=(0, 25))
        
        start_button = tk.Button(
            container,
            text="시작하기",
            font=self.font_button,
            fg="white",
            bg=self.colors["primary"],
            activebackground=self.colors["primary_dark"],
            activeforeground="white",
            bd=0,
            padx=40,
            pady=10,
            relief="flat",
            cursor="hand2",
            command=self.show_game_screen
        )
        start_button.pack(pady=(0, 10))
        
        tip = tk.Label(
            container,
            text="※ 총 20장의 카드를 뽑을 수 있습니다.",
            font=("맑은 고딕", 10),
            fg=self.colors["text_sub"],
            bg=self.colors["bg"]
        )
        tip.pack(pady=(10, 0))
    
    # ---------- 게임 화면 ----------
    def create_game_screen(self):
        # 상단 헤더
        header = tk.Frame(self.game_frame, bg=self.colors["bg"])
        header.pack(fill="x", padx=30, pady=(20, 10))
        
        # 그리드로 가운데 정렬: 좌/중앙/우 3칸
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=1)
        
        title_label = tk.Label(
            header,
            text="스트림스 카드 뽑기",
            font=self.font_title,
            fg=self.colors["text_main"],
            bg=self.colors["bg"]
        )
        # 중앙(열 1)에 배치
        title_label.grid(row=0, column=1)
        
        self.counter_label = tk.Label(
            header,
            text="0 / 20 장",
            font=("맑은 고딕", 12, "bold"),
            fg=self.colors["accent"],
            bg=self.colors["bg"]
        )
        # 오른쪽(열 2) 정렬
        self.counter_label.grid(row=0, column=2, sticky="e")
        
        # 중앙 메인 영역
        main_area = tk.Frame(self.game_frame, bg=self.colors["bg"])
        main_area.pack(fill="both", expand=True, padx=30, pady=10)
        
        # 중앙 카드 영역
        card_frame = tk.Frame(
            main_area,
            bg=self.colors["card_bg"],
            bd=0,
            highlightthickness=2,
            highlightbackground=self.colors["card_border"]
        )
        card_frame.place(relx=0.5, rely=0.32, anchor="center", width=260, height=180)
        
        self.card_label = tk.Label(
            card_frame,
            text="Ready",
            font=self.font_card,
            fg=self.colors["accent"],
            bg=self.colors["card_bg"]
        )
        self.card_label.pack(expand=True)
        
        # 카드 설명 텍스트
        info_label = tk.Label(
            main_area,
            text="카드 뽑기 버튼을 눌러 다음 카드를 확인하세요.",
            font=self.font_subtitle,
            fg=self.colors["text_sub"],
            bg=self.colors["bg"]
        )
        info_label.place(relx=0.5, rely=0.55, anchor="center")
        
        # 버튼 영역
        button_frame = tk.Frame(main_area, bg=self.colors["bg"])
        button_frame.place(relx=0.5, rely=0.70, anchor="center")
        
        self.draw_button = tk.Button(
            button_frame,
            text="카드 뽑기",
            font=self.font_button,
            fg="white",
            bg=self.colors["primary"],
            activebackground=self.colors["primary_dark"],
            activeforeground="white",
            bd=0,
            padx=30,
            pady=10,
            relief="flat",
            cursor="hand2",
            command=self.draw_card
        )
        self.draw_button.grid(row=0, column=0, padx=8)
        
        reset_button = tk.Button(
            button_frame,
            text="다시하기",
            font=self.font_button,
            fg=self.colors["text_main"],
            bg=self.colors["secondary"],
            activebackground=self.colors["secondary_dark"],
            activeforeground=self.colors["text_main"],
            bd=0,
            padx=24,
            pady=10,
            relief="flat",
            cursor="hand2",
            command=self.reset_game
        )
        reset_button.grid(row=0, column=1, padx=8)
        
        back_button = tk.Button(
            button_frame,
            text="처음으로",
            font=self.font_button,
            fg=self.colors["text_main"],
            bg="#E5E7EB",
            activebackground="#D1D5DB",
            activeforeground=self.colors["text_main"],
            bd=0,
            padx=24,
            pady=10,
            relief="flat",
            cursor="hand2",
            command=self.show_start_screen
        )
        back_button.grid(row=0, column=2, padx=8)
        
        # 하단 테이블 영역
        table_outer = tk.Frame(
            self.game_frame,
            bg=self.colors["table_bg"],
            bd=0,
            highlightthickness=0
        )
        table_outer.pack(fill="x", padx=30, pady=(0, 25))
        
        # 제목을 가운데 정렬
        header_label = tk.Label(
            table_outer,
            text="뽑힌 카드 기록 (왼쪽부터 순서대로)",
            font=self.font_table_header,
            fg=self.colors["text_main"],
            bg=self.colors["table_bg"]
        )
        header_label.pack(anchor="center", padx=12, pady=(8, 4))
        
        table_frame = tk.Frame(table_outer, bg=self.colors["table_bg"])
        # 테이블 전체를 가운데에 배치
        table_frame.pack(anchor="center", padx=12, pady=(0, 10))
        
        # 20개 셀 (2행 x 10열) 생성
        self.table_cells = []
        total_cells = 20
        cols = 10
        
        for i in range(total_cells):
            r = i // cols
            c = i % cols
            cell = tk.Label(
                table_frame,
                text=" ",
                width=4,
                height=2,
                font=self.font_table_cell,
                fg=self.colors["text_main"],
                bg=self.colors["table_cell"],
                bd=0,
                relief="flat"
            )
            cell.grid(row=r, column=c, padx=3, pady=3, ipadx=4, ipady=2)
            self.table_cells.append(cell)
    
    # ---------- 카드 뽑기/화면 업데이트 ----------
    def update_card_display(self, value):
        if value is None:
            self.card_label.config(text="Ready", fg=self.colors["accent"])
        else:
            self.card_label.config(text=str(value), fg=self.colors["text_main"])
    
    def update_counter(self):
        count = len(self.drawn_cards)
        self.counter_label.config(text=f"{count} / 20 장")
    
    def draw_card(self):
        # 20장까지 뽑기
        if len(self.drawn_cards) >= 20:
            self.draw_button.config(state="disabled", bg="#9CA3AF")
            return
        
        if not self.deck:
            # 덱이 비어 있을 경우 (이론상 거의 없음)
            self.draw_button.config(state="disabled", bg="#9CA3AF")
            return
        
        card = self.deck.pop()  # 덱에서 한 장 뽑기 (중복 없이)
        self.drawn_cards.append(card)
        
        # 중앙 카드 표시
        self.update_card_display(card)
        # 카운터 갱신
        self.update_counter()
        
        # 테이블에 기록
        idx = len(self.drawn_cards) - 1
        if idx < len(self.table_cells):
            cell = self.table_cells[idx]
            cell.config(text=str(card), bg="#FEF3C7")  # 뽑힌 칸은 살짝 다른 색
        
        # 20번째 카드를 뽑으면 버튼 비활성화
        if len(self.drawn_cards) >= 20:
            self.draw_button.config(state="disabled", bg="#9CA3AF")
            self.card_label.config(fg=self.colors["accent"])
    

if __name__ == "__main__":
    root = tk.Tk()
    app = StreamsApp(root)
    root.mainloop()
