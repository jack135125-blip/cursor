import tkinter as tk
from tkinter import messagebox, scrolledtext
from youtube_transcript_api import YouTubeTranscriptApi, TranscriptsDisabled

class YoutubeTranscriptApp:
    def __init__(self, root):
        self.root = root
        self.root.title("유튜브 자막 추출기")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")

        # 프레임 생성
        self.top_frame = tk.Frame(root, bg="#f0f0f0")
        self.top_frame.pack(pady=20, fill=tk.X, padx=20)

        self.middle_frame = tk.Frame(root, bg="#f0f0f0")
        self.middle_frame.pack(pady=10, fill=tk.X, padx=20)

        self.bottom_frame = tk.Frame(root, bg="#f0f0f0")
        self.bottom_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)

        # 입력 필드 및 버튼
        self.url_label = tk.Label(self.top_frame, text="유튜브 URL 또는 영상 ID:", bg="#f0f0f0", font=("맑은 고딕", 12))
        self.url_label.pack(side=tk.LEFT, padx=(0, 10))

        self.url_entry = tk.Entry(self.top_frame, width=50, font=("맑은 고딕", 12))
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 언어 선택
        self.lang_label = tk.Label(self.middle_frame, text="언어 선택:", bg="#f0f0f0", font=("맑은 고딕", 12))
        self.lang_label.pack(side=tk.LEFT, padx=(0, 10))

        self.lang_var = tk.StringVar()
        self.lang_var.set("ko")  # 기본값 한국어
        
        langs = [("한국어", "ko"), ("영어", "en"), ("자동 감지", "")]
        
        self.lang_frame = tk.Frame(self.middle_frame, bg="#f0f0f0")
        self.lang_frame.pack(side=tk.LEFT)
        
        for text, value in langs:
            tk.Radiobutton(
                self.lang_frame, 
                text=text, 
                variable=self.lang_var, 
                value=value, 
                bg="#f0f0f0",
                font=("맑은 고딕", 10)
            ).pack(side=tk.LEFT, padx=5)

        # 실행 버튼
        self.fetch_button = tk.Button(
            self.middle_frame, 
            text="자막 가져오기", 
            command=self.fetch_transcript,
            bg="#4CAF50",
            fg="white",
            font=("맑은 고딕", 12, "bold"),
            relief=tk.RAISED,
            padx=15
        )
        self.fetch_button.pack(side=tk.RIGHT, padx=10)
        
        # 결과 표시 영역
        self.result_label = tk.Label(self.bottom_frame, text="자막 결과:", bg="#f0f0f0", font=("맑은 고딕", 12))
        self.result_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.result_text = scrolledtext.ScrolledText(
            self.bottom_frame, 
            wrap=tk.WORD, 
            width=80, 
            height=20,
            font=("맑은 고딕", 10)
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 복사, 저장 버튼 프레임
        self.button_frame = tk.Frame(self.bottom_frame, bg="#f0f0f0")
        self.button_frame.pack(pady=10, fill=tk.X)
        
        self.copy_button = tk.Button(
            self.button_frame, 
            text="클립보드에 복사", 
            command=self.copy_to_clipboard,
            bg="#2196F3",
            fg="white",
            font=("맑은 고딕", 10)
        )
        self.copy_button.pack(side=tk.LEFT, padx=5)
        
        self.save_button = tk.Button(
            self.button_frame, 
            text="텍스트 파일로 저장", 
            command=self.save_to_file,
            bg="#2196F3",
            fg="white",
            font=("맑은 고딕", 10)
        )
        self.save_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = tk.Button(
            self.button_frame, 
            text="내용 지우기", 
            command=self.clear_result,
            bg="#f44336",
            fg="white",
            font=("맑은 고딕", 10)
        )
        self.clear_button.pack(side=tk.RIGHT, padx=5)

    def extract_video_id(self, url):
        """URL에서 유튜브 영상 ID를 추출합니다."""
        if 'youtube.com' in url:
            if 'v=' in url:
                video_id = url.split('v=')[1].split('&')[0]
                return video_id
            return None
        elif 'youtu.be' in url:
            video_id = url.split('/')[-1].split('?')[0]
            return video_id
        else:
            # URL이 아니라 직접 ID를 입력한 경우
            return url

    def fetch_transcript(self):
        """유튜브 자막을 가져오는 함수"""
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showerror("오류", "유튜브 URL 또는 영상 ID를 입력해주세요.")
            return
        
        video_id = self.extract_video_id(url)
        if not video_id:
            messagebox.showerror("오류", "올바른 유튜브 URL이 아닙니다.")
            return
        
        lang = self.lang_var.get()
        
        try:
            if lang:
                transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=[lang])
            else:
                transcript = YouTubeTranscriptApi.get_transcript(video_id)
                
            # 자막 텍스트 형식 가공
            transcript_text = ""
            for item in transcript:
                start_time = self.format_time(item['start'])
                text = item['text']
                transcript_text += f"[{start_time}] {text}\n\n"
                
            # 결과 표시
            self.result_text.delete(1.0, tk.END)  # 기존 내용 삭제
            self.result_text.insert(tk.END, transcript_text)
            
        except TranscriptsDisabled:
            messagebox.showerror("오류", "이 영상에는 자막이 비활성화되어 있습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"자막을 가져오는 중 오류가 발생했습니다: {str(e)}")

    def format_time(self, seconds):
        """초를 시:분:초 형식으로 변환"""
        minutes, seconds = divmod(int(seconds), 60)
        hours, minutes = divmod(minutes, 60)
        
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        else:
            return f"{minutes:02d}:{seconds:02d}"

    def copy_to_clipboard(self):
        """결과를 클립보드에 복사"""
        text = self.result_text.get(1.0, tk.END)
        if text.strip():
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("알림", "클립보드에 복사되었습니다.")
        else:
            messagebox.showinfo("알림", "복사할 내용이 없습니다.")

    def save_to_file(self):
        """결과를 텍스트 파일로 저장"""
        from tkinter import filedialog
        import os
        
        text = self.result_text.get(1.0, tk.END)
        if not text.strip():
            messagebox.showinfo("알림", "저장할 내용이 없습니다.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("텍스트 파일", "*.txt"), ("모든 파일", "*.*")],
            initialdir=os.path.expanduser("~\\Documents"),
            title="자막 저장"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(text)
                messagebox.showinfo("저장 완료", f"파일이 저장되었습니다:\n{file_path}")
            except Exception as e:
                messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다: {str(e)}")

    def clear_result(self):
        """결과 창 내용 지우기"""
        self.result_text.delete(1.0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = YoutubeTranscriptApp(root)
    root.mainloop()
