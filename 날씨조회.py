import tkinter as tk
from tkinter import ttk, messagebox, font
import requests
import json
from datetime import datetime
import logging

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WeatherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("대한민국 지역별 날씨 조회")
        
        # 파스텔 톤 색상 정의
        self.colors = {
            "bg": "#f8f9fa",
            "highlight": "#d0e8f2",
            "accent1": "#a2d2ff",
            "accent2": "#bde0fe",
            "text": "#495057",
            "border": "#dee2e6"
        }
        
        # 지역 목록 (대한민국 주요 도시)
        self.cities = [
            "서울", "부산", "인천", "대구", "대전", "광주", "울산", "세종",
            "경기도", "강원도", "충청북도", "충청남도", "전라북도", "전라남도", "경상북도", "경상남도", "제주도"
        ]
        
        # 영어 지역명 매핑 (WeatherAPI용)
        self.city_to_eng = {
            "서울": "Seoul",
            "부산": "Busan",
            "인천": "Incheon",
            "대구": "Daegu",
            "대전": "Daejeon",
            "광주": "Gwangju",
            "울산": "Ulsan",
            "세종": "Sejong",
            "경기도": "Gyeonggi-do",
            "강원도": "Gangwon-do",
            "충청북도": "Chungcheongbuk-do",
            "충청남도": "Chungcheongnam-do",
            "전라북도": "Jeollabuk-do",
            "전라남도": "Jeollanam-do",
            "경상북도": "Gyeongsangbuk-do",
            "경상남도": "Gyeongsangnam-do",
            "제주도": "Jeju-do"
        }
        
        # 폰트 설정
        self.custom_font = font.Font(family="맑은 고딕", size=10)
        self.title_font = font.Font(family="맑은 고딕", size=14, weight="bold")
        
        self.setup_ui()
    
    def setup_ui(self):
        # 메인 프레임 설정
        self.root.configure(bg=self.colors["bg"])
        self.root.resizable(False, False)
        
        # 패딩 및 스타일 설정
        padx, pady = 15, 10
        
        # 제목 레이블
        title_frame = tk.Frame(self.root, bg=self.colors["bg"])
        title_frame.pack(fill="x", padx=padx, pady=pady)
        
        title_label = tk.Label(
            title_frame, 
            text="대한민국 지역별 날씨 조회", 
            font=self.title_font,
            bg=self.colors["bg"],
            fg=self.colors["text"]
        )
        title_label.pack(pady=10)
        
        # 지역 선택 프레임
        selection_frame = tk.Frame(self.root, bg=self.colors["bg"], bd=1, relief=tk.SOLID, highlightbackground=self.colors["border"])
        selection_frame.pack(fill="x", padx=padx, pady=pady)
        
        city_label = tk.Label(
            selection_frame, 
            text="지역 선택:", 
            font=self.custom_font,
            bg=self.colors["bg"],
            fg=self.colors["text"]
        )
        city_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        # 콤보박스 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TCombobox', 
                        fieldbackground=self.colors["highlight"],
                        background=self.colors["accent1"],
                        foreground=self.colors["text"])
        
        self.city_combobox = ttk.Combobox(
            selection_frame, 
            values=self.cities,
            font=self.custom_font,
            width=15,
            state="readonly"
        )
        self.city_combobox.grid(row=0, column=1, padx=10, pady=10)
        self.city_combobox.current(0)  # 기본값: 서울
        
        search_button = tk.Button(
            selection_frame, 
            text="날씨 조회",
            font=self.custom_font,
            bg=self.colors["accent1"],
            fg=self.colors["text"],
            activebackground=self.colors["accent2"],
            relief=tk.RAISED,
            command=self.get_weather
        )
        search_button.grid(row=0, column=2, padx=10, pady=10)
        
        # 결과 표시 프레임
        self.result_frame = tk.Frame(self.root, bg=self.colors["bg"], bd=1, relief=tk.SOLID, highlightbackground=self.colors["border"])
        self.result_frame.pack(fill="both", expand=True, padx=padx, pady=pady)
        
        # 날씨 정보 레이블
        self.weather_info = tk.Label(
            self.result_frame,
            text="지역을 선택하고 날씨 조회 버튼을 눌러주세요.",
            font=self.custom_font,
            bg=self.colors["bg"],
            fg=self.colors["text"],
            justify=tk.LEFT,
            padx=20,
            pady=20
        )
        self.weather_info.pack(fill="both", expand=True)
        
        # 하단 상태바
        status_frame = tk.Frame(self.root, bg=self.colors["border"], height=25)
        status_frame.pack(fill="x", side=tk.BOTTOM)
        
        self.status_label = tk.Label(
            status_frame,
            text="준비됨",
            font=("맑은 고딕", 8),
            bg=self.colors["border"],
            fg=self.colors["text"],
            anchor="w",
            padx=10
        )
        self.status_label.pack(side=tk.LEFT)
        
        # 현재 시간 표시
        self.time_label = tk.Label(
            status_frame,
            text=self.get_current_time(),
            font=("맑은 고딕", 8),
            bg=self.colors["border"],
            fg=self.colors["text"],
            anchor="e",
            padx=10
        )
        self.time_label.pack(side=tk.RIGHT)
        
        # 1초마다 시간 업데이트
        self.update_time()
    
    def get_current_time(self):
        now = datetime.now()
        return now.strftime("%Y-%m-%d %H:%M:%S")
    
    def update_time(self):
        self.time_label.config(text=self.get_current_time())
        self.root.after(1000, self.update_time)
    
    def get_weather(self):
        city = self.city_combobox.get()
        eng_city = self.city_to_eng.get(city, city)  # 영어 이름으로 변환
        self.status_label.config(text=f"{city} 날씨 정보를 가져오는 중...")
        
        try:
            # WeatherAPI 사용 (무료 API)
            url = f"https://wttr.in/{eng_city}?format=j1"
            
            logger.info(f"{city} 날씨 정보 요청: {url}")
            
            response = requests.get(url)
            
            if response.status_code == 200:
                try:
                    weather_data = response.json()
                    
                    # 날씨 정보 파싱
                    current = weather_data["current_condition"][0]
                    temp_c = current["temp_C"]
                    feels_like = current["FeelsLikeC"]
                    humidity = current["humidity"]
                    description = current["weatherDesc"][0]["value"]
                    wind_speed = current["windspeedKmph"]
                    
                    # 일출, 일몰 시간
                    astronomy = weather_data["weather"][0]["astronomy"][0]
                    sunrise = astronomy["sunrise"]
                    sunset = astronomy["sunset"]
                    
                    # 결과 메시지 구성
                    result_message = f"""
지역: {city}
현재 기온: {temp_c}°C (체감 온도: {feels_like}°C)
날씨 상태: {description}
습도: {humidity}%
풍속: {wind_speed} km/h
일출: {sunrise}
일몰: {sunset}
                    """
                    
                    # 결과 업데이트
                    self.weather_info.config(text=result_message)
                    self.status_label.config(text=f"{city} 날씨 정보 조회 완료")
                    logger.info(f"{city} 날씨 정보 조회 성공")
                except json.JSONDecodeError as e:
                    error_msg = f"JSON 파싱 오류: {str(e)}"
                    self.weather_info.config(text=error_msg)
                    self.status_label.config(text="JSON 파싱 오류")
                    logger.error(error_msg, exc_info=True)
                
            else:
                error_msg = f"날씨 정보를 가져오는데 실패했습니다. 상태 코드: {response.status_code}"
                self.weather_info.config(text=error_msg)
                self.status_label.config(text="날씨 정보 조회 실패")
                logger.error(error_msg)
                
        except Exception as e:
            error_msg = f"오류 발생: {str(e)}"
            self.weather_info.config(text=error_msg)
            self.status_label.config(text="오류 발생")
            logger.error(error_msg, exc_info=True)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = WeatherApp(root)
        
        # 창 크기 조정
        root.update()
        root.geometry("")  # 자동으로 크기 조정
        
        root.mainloop()
    except Exception as e:
        logger.critical(f"애플리케이션 실행 중 심각한 오류 발생: {str(e)}", exc_info=True)
