import tkinter as tk
from tkinter import ttk, messagebox
import requests
import json
import logging
from datetime import datetime
import re
import threading

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# API 설정
BASE_URL = "https://open.neis.go.kr/hub/"
SCHOOL_INFO_URL = BASE_URL + "schoolInfo"
MEAL_INFO_URL = BASE_URL + "mealServiceDietInfo"

# 색상 설정 (파스텔톤)
COLORS = {
    "background": "#F8F9FA",
    "primary": "#E8F5E9",
    "secondary": "#F3E5F5",
    "accent": "#FFEBEE",
    "text": "#212121",
    "button": "#D1C4E9",
    "button_hover": "#B39DDB",
    "frame": "#E0F7FA"
}

class SchoolMealApp:
    def __init__(self, root):
        self.root = root
        self.root.title("학교 급식 조회 프로그램")
        self.root.configure(bg=COLORS["background"])
        
        # 변수 초기화
        self.school_level_var = tk.StringVar()
        self.school_name_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        
        # 학교 레벨 목록
        self.school_levels = {
            "유치원": "1",
            "초등학교": "2",
            "중학교": "3",
            "고등학교": "4",
            "특수학교": "5"
        }
        
        # 지역코드 (교육청)
        self.region_codes = {
            "서울특별시": "B10",
            "부산광역시": "C10",
            "대구광역시": "D10",
            "인천광역시": "E10",
            "광주광역시": "F10",
            "대전광역시": "G10",
            "울산광역시": "H10",
            "세종특별자치시": "I10",
            "경기도": "J10",
            "강원도": "K10",
            "충청북도": "M10",
            "충청남도": "N10",
            "전라북도": "P10",
            "전라남도": "Q10",
            "경상북도": "R10",
            "경상남도": "S10",
            "제주특별자치도": "T10"
        }
        
        # 메인 프레임
        self.main_frame = tk.Frame(self.root, bg=COLORS["background"], padx=20, pady=20)
        self.main_frame.pack(fill="both", expand=True)
        
        # 검색 프레임
        self.create_search_frame()
        
        # 결과 프레임
        self.create_result_frame()
        
        # 학교 데이터 캐시
        self.school_cache = {}
        
    def create_search_frame(self):
        # 검색 프레임
        search_frame = tk.LabelFrame(self.main_frame, text="학교 검색", bg=COLORS["frame"], fg=COLORS["text"], 
                                     padx=15, pady=15, font=("나눔고딕", 10, "bold"))
        search_frame.pack(fill="x", padx=10, pady=10)
        
        # 지역 선택
        region_label = tk.Label(search_frame, text="지역:", bg=COLORS["frame"], fg=COLORS["text"], font=("나눔고딕", 10))
        region_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.region_var = tk.StringVar()
        region_combo = ttk.Combobox(search_frame, textvariable=self.region_var, width=15, 
                                    values=list(self.region_codes.keys()), state="readonly")
        region_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        region_combo.bind("<<ComboboxSelected>>", self.update_school_list)
        
        # 학교급 선택
        level_label = tk.Label(search_frame, text="학교급:", bg=COLORS["frame"], fg=COLORS["text"], font=("나눔고딕", 10))
        level_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        level_combo = ttk.Combobox(search_frame, textvariable=self.school_level_var, width=12, 
                                  values=list(self.school_levels.keys()), state="readonly")
        level_combo.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        level_combo.bind("<<ComboboxSelected>>", self.update_school_list)
        
        # 학교명 검색
        school_label = tk.Label(search_frame, text="학교명:", bg=COLORS["frame"], fg=COLORS["text"], font=("나눔고딕", 10))
        school_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        self.school_entry = tk.Entry(search_frame, width=20, font=("나눔고딕", 10))
        self.school_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.school_entry.bind('<Return>', lambda event: self.search_schools())
        
        search_button = tk.Button(search_frame, text="검색", bg=COLORS["button"], fg=COLORS["text"],
                                 activebackground=COLORS["button_hover"], font=("나눔고딕", 9, "bold"),
                                 command=self.search_schools)
        search_button.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        
        # 로딩 라벨
        self.loading_label = tk.Label(search_frame, text="", bg=COLORS["frame"], fg="red", font=("나눔고딕", 9))
        self.loading_label.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        
        # 학교 목록
        school_list_label = tk.Label(search_frame, text="학교 선택:", bg=COLORS["frame"], fg=COLORS["text"], font=("나눔고딕", 10))
        school_list_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        
        # 스크롤바가 있는 리스트박스
        listbox_frame = tk.Frame(search_frame, bg=COLORS["frame"])
        listbox_frame.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="we")
        
        self.school_listbox = tk.Listbox(listbox_frame, width=50, height=8, font=("나눔고딕", 10))
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical")
        self.school_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.school_listbox.yview)
        
        self.school_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 날짜 선택
        date_label = tk.Label(search_frame, text="날짜:", bg=COLORS["frame"], fg=COLORS["text"], font=("나눔고딕", 10))
        date_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        
        self.date_entry = tk.Entry(search_frame, textvariable=self.date_var, width=10, font=("나눔고딕", 10))
        self.date_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        date_format_label = tk.Label(search_frame, text="(YYYYMMDD 형식)", bg=COLORS["frame"], fg=COLORS["text"], font=("나눔고딕", 8))
        date_format_label.grid(row=3, column=2, padx=5, pady=5, sticky="w")
        
        # 조회 버튼
        search_meal_button = tk.Button(search_frame, text="급식 조회", bg=COLORS["button"], fg=COLORS["text"],
                                     activebackground=COLORS["button_hover"], font=("나눔고딕", 10, "bold"),
                                     command=self.search_meal)
        search_meal_button.grid(row=4, column=1, padx=5, pady=10, sticky="we")
        
    def create_result_frame(self):
        # 결과 프레임
        self.result_frame = tk.LabelFrame(self.main_frame, text="급식 정보", bg=COLORS["frame"], fg=COLORS["text"], 
                                        padx=15, pady=15, font=("나눔고딕", 10, "bold"))
        self.result_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 학교 정보
        self.school_info_label = tk.Label(self.result_frame, text="", bg=COLORS["frame"], fg=COLORS["text"], font=("나눔고딕", 11, "bold"))
        self.school_info_label.pack(fill="x", padx=5, pady=5)
        
        # 탭 컨트롤 생성
        self.tab_control = ttk.Notebook(self.result_frame)
        
        # 아침, 점심, 저녁 탭
        self.breakfast_tab = tk.Frame(self.tab_control, bg=COLORS["primary"])
        self.lunch_tab = tk.Frame(self.tab_control, bg=COLORS["secondary"])
        self.dinner_tab = tk.Frame(self.tab_control, bg=COLORS["accent"])
        
        self.tab_control.add(self.breakfast_tab, text="아침")
        self.tab_control.add(self.lunch_tab, text="점심")
        self.tab_control.add(self.dinner_tab, text="저녁")
        
        self.tab_control.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 각 탭에 텍스트 위젯 추가
        self.breakfast_text = tk.Text(self.breakfast_tab, bg=COLORS["primary"], fg=COLORS["text"], 
                                    font=("나눔고딕", 11), wrap=tk.WORD, height=10, width=50)
        self.breakfast_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.lunch_text = tk.Text(self.lunch_tab, bg=COLORS["secondary"], fg=COLORS["text"], 
                                font=("나눔고딕", 11), wrap=tk.WORD, height=10, width=50)
        self.lunch_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.dinner_text = tk.Text(self.dinner_tab, bg=COLORS["accent"], fg=COLORS["text"], 
                                 font=("나눔고딕", 11), wrap=tk.WORD, height=10, width=50)
        self.dinner_text.pack(fill="both", expand=True, padx=10, pady=10)
        
    def fetch_schools_from_api(self, region_code=None, school_level=None, school_name=None):
        """나이스 API에서 학교 정보를 가져오는 함수"""
        try:
            params = {
                'Type': 'json',
                'pIndex': 1,
                'pSize': 1000  # 한 번에 가져올 최대 개수
            }
            
            if region_code:
                params['ATPT_OFCDC_SC_CODE'] = region_code
            if school_level:
                params['SCHUL_KND_SC_NM'] = school_level
            if school_name:
                params['SCHUL_NM'] = school_name
                
            response = requests.get(SCHOOL_INFO_URL, params=params, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            
            if 'schoolInfo' in data and len(data['schoolInfo']) > 1:
                schools = data['schoolInfo'][1]['row']
                return schools
            else:
                logger.info("검색 결과가 없습니다.")
                return []
                
        except requests.exceptions.RequestException as e:
            logger.error(f"API 요청 중 오류 발생: {str(e)}")
            return []
        except Exception as e:
            logger.error(f"학교 정보 조회 중 오류 발생: {str(e)}")
            return []
    
    def update_school_list(self, event=None):
        """지역과 학교급 선택 시 학교 목록 업데이트"""
        region = self.region_var.get()
        school_level = self.school_level_var.get()
        
        if not region or not school_level:
            return
            
        # 백그라운드에서 학교 목록 가져오기
        def fetch_and_update():
            self.loading_label.config(text="로딩 중...")
            self.school_listbox.delete(0, tk.END)
            
            region_code = self.region_codes[region]
            schools = self.fetch_schools_from_api(region_code=region_code, school_level=school_level)
            
            # UI 업데이트는 메인 스레드에서
            self.root.after(0, lambda: self.update_listbox_with_schools(schools))
            
        threading.Thread(target=fetch_and_update, daemon=True).start()
    
    def search_schools(self):
        """학교명으로 검색"""
        search_text = self.school_entry.get().strip()
        if not search_text:
            messagebox.showwarning("경고", "학교명을 입력해주세요.")
            return
            
        # 백그라운드에서 검색
        def fetch_and_search():
            self.loading_label.config(text="검색 중...")
            self.school_listbox.delete(0, tk.END)
            
            # 선택된 지역과 학교급이 있으면 해당 조건으로 검색
            region = self.region_var.get()
            school_level = self.school_level_var.get()
            
            region_code = self.region_codes[region] if region else None
            
            schools = self.fetch_schools_from_api(
                region_code=region_code,
                school_level=school_level if school_level else None,
                school_name=search_text
            )
            
            # UI 업데이트는 메인 스레드에서
            self.root.after(0, lambda: self.update_listbox_with_schools(schools, is_search=True))
            
        threading.Thread(target=fetch_and_search, daemon=True).start()
    
    def update_listbox_with_schools(self, schools, is_search=False):
        """학교 목록을 리스트박스에 업데이트"""
        self.loading_label.config(text="")
        self.school_listbox.delete(0, tk.END)
        
        if not schools:
            self.school_listbox.insert(tk.END, "검색 결과가 없습니다.")
            return
            
        for school in schools:
            try:
                school_name = school.get('SCHUL_NM', '')
                region_name = school.get('LCTN_SC_NM', '')
                school_kind = school.get('SCHUL_KND_SC_NM', '')
                
                if is_search:
                    display_text = f"{school_name} ({region_name} {school_kind})"
                else:
                    display_text = school_name
                    
                self.school_listbox.insert(tk.END, display_text)
                
            except Exception as e:
                logger.error(f"학교 정보 처리 중 오류: {str(e)}")
                continue
                
        logger.info(f"총 {len(schools)}개의 학교를 찾았습니다.")
    
    def get_school_info(self):
        """선택된 학교의 정보 반환"""
        selected_indices = self.school_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("경고", "학교를 선택해주세요.")
            return None
            
        selected_text = self.school_listbox.get(selected_indices[0])
        
        if "검색 결과가 없습니다." in selected_text:
            messagebox.showwarning("경고", "유효한 학교를 선택해주세요.")
            return None
        
        # 선택된 학교의 상세 정보를 API에서 다시 가져오기
        if " (" in selected_text and ")" in selected_text:
            # 검색 결과 형태: "학교명 (지역 학교급)"
            school_name = selected_text.split(" (")[0]
        else:
            # 일반 목록에서 선택한 경우
            school_name = selected_text
            
        # API에서 해당 학교의 상세 정보 가져오기
        schools = self.fetch_schools_from_api(school_name=school_name)
        
        if schools:
            # 첫 번째 결과 반환 (정확한 이름 매칭)
            for school in schools:
                if school.get('SCHUL_NM') == school_name:
                    return {
                        'name': school.get('SCHUL_NM', ''),
                        'code': school.get('SD_SCHUL_CODE', ''),
                        'region_code': school.get('ATPT_OFCDC_SC_CODE', ''),
                        'region_name': school.get('LCTN_SC_NM', ''),
                        'school_kind': school.get('SCHUL_KND_SC_NM', '')
                    }
        
        return None
    
    def search_meal(self):
        school_info = self.get_school_info()
        if not school_info:
            return
            
        # 날짜 검증
        date_str = self.date_var.get()
        if not re.match(r'^\d{8}$', date_str):
            messagebox.showwarning("경고", "날짜는 YYYYMMDD 형식으로 입력해주세요.")
            return
            
        try:
            # 백그라운드에서 급식 정보 가져오기
            def fetch_meal_info():
                self.loading_label.config(text="급식 조회 중...")
                
                meals = self.fetch_meal_from_api(school_info, date_str)
                
                # UI 업데이트는 메인 스레드에서
                self.root.after(0, lambda: self.display_meal_info(school_info, date_str, meals))
                
            threading.Thread(target=fetch_meal_info, daemon=True).start()
            
        except Exception as e:
            logger.error(f"급식 정보 조회 중 오류 발생: {str(e)}")
            messagebox.showerror("오류", f"급식 정보 조회 중 오류가 발생했습니다.\n{str(e)}")
    
    def fetch_meal_from_api(self, school_info, date_str):
        """나이스 API에서 급식 정보를 가져오는 함수"""
        try:
            params = {
                'Type': 'json',
                'pIndex': 1,
                'pSize': 100,
                'ATPT_OFCDC_SC_CODE': school_info['region_code'],
                'SD_SCHUL_CODE': school_info['code'],
                'MLSV_YMD': date_str
            }
            
            response = requests.get(MEAL_INFO_URL, params=params, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            
            if 'mealServiceDietInfo' in data and len(data['mealServiceDietInfo']) > 1:
                meals = data['mealServiceDietInfo'][1]['row']
                return meals
            else:
                logger.info(f"{date_str}에 대한 급식 정보가 없습니다.")
                return []
                
        except requests.exceptions.RequestException as e:
            logger.error(f"급식 API 요청 중 오류 발생: {str(e)}")
            return []
        except Exception as e:
            logger.error(f"급식 정보 조회 중 오류 발생: {str(e)}")
            return []
    
    def display_meal_info(self, school_info, date_str, meals):
        """급식 정보를 화면에 표시"""
        self.loading_label.config(text="")
        
        # 학교 정보 표시
        self.school_info_label.config(
            text=f"{school_info['name']} ({date_str[:4]}년 {date_str[4:6]}월 {date_str[6:8]}일)"
        )
        
        # 텍스트 위젯 초기화
        self.breakfast_text.delete(1.0, tk.END)
        self.lunch_text.delete(1.0, tk.END)
        self.dinner_text.delete(1.0, tk.END)
        
        if not meals:
            no_meal_msg = "해당 날짜에 급식 정보가 없습니다."
            self.breakfast_text.insert(tk.END, no_meal_msg)
            self.lunch_text.insert(tk.END, no_meal_msg)
            self.dinner_text.insert(tk.END, no_meal_msg)
            return
        
        # 급식 정보를 시간대별로 분류
        meal_by_time = {
            '조식': '',
            '중식': '',
            '석식': ''
        }
        
        for meal in meals:
            try:
                meal_type = meal.get('MMEAL_SC_NM', '')  # 급식구분명 (조식, 중식, 석식)
                dish_names = meal.get('DDISH_NM', '')    # 요리명
                
                if dish_names:
                    # HTML 태그 제거 및 정리
                    clean_dishes = self.clean_meal_text(dish_names)
                    meal_by_time[meal_type] = clean_dishes
                    
            except Exception as e:
                logger.error(f"급식 정보 처리 중 오류: {str(e)}")
                continue
        
        # 각 탭에 급식 정보 표시
        self.breakfast_text.insert(tk.END, meal_by_time.get('조식', '조식 정보가 없습니다.'))
        self.lunch_text.insert(tk.END, meal_by_time.get('중식', '중식 정보가 없습니다.'))
        self.dinner_text.insert(tk.END, meal_by_time.get('석식', '석식 정보가 없습니다.'))
        
        logger.info(f"{school_info['name']}의 {date_str} 급식 정보를 조회했습니다.")
    
    def clean_meal_text(self, text):
        """급식 텍스트에서 HTML 태그 및 알레르기 정보 정리"""
        if not text:
            return ""
            
        # HTML 태그 제거
        import re
        text = re.sub(r'<[^>]+>', '', text)
        
        # 알레르기 정보 (숫자.숫자 형태) 제거
        text = re.sub(r'\d+\.?\d*\.?', '', text)
        
        # 특수문자 정리
        text = text.replace('*', '').replace('&', '').replace('#', '')
        
        # 연속된 공백 제거 및 줄바꿈 정리
        text = re.sub(r'\s+', ' ', text)
        text = text.replace('<br/>', '\n').replace('<br>', '\n')
        
        # 각 메뉴를 줄바꿈으로 구분
        dishes = [dish.strip() for dish in text.split('<br/>') if dish.strip()]
        if not dishes:
            dishes = [dish.strip() for dish in text.split('\n') if dish.strip()]
        if not dishes:
            dishes = [dish.strip() for dish in text.split(',') if dish.strip()]
            
        return '\n'.join(dishes) if dishes else text.strip()
    
    def display_sample_meal(self, school_info, date_str):
        # 이 함수는 더 이상 사용하지 않음 (실제 API 사용)
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = SchoolMealApp(root)
    
    # 창 크기 자동 조정
    root.update()
    root.minsize(root.winfo_width(), root.winfo_height())
    
    root.mainloop()
