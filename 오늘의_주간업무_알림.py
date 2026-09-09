# -*- coding: utf-8 -*-
"""
창원여고 주간 업무 계획 - 오늘의 알림 팝업
Google 스프레드시트에서 오늘 날짜 열만 읽어 둥근 팝업으로 표시합니다.
창을 닫으면 프로그램이 종료됩니다.
"""

from __future__ import annotations

import csv
import io
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from pathlib import Path

import customtkinter as ctk

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------
SPREADSHEET_ID = "1dOzQRmT3eIKW3OuIQlDAw7WkHRqn8j-kdLbv18x2zPM"
CACHE_PATH = Path(__file__).with_name("today_schedule_cache.json")

# 아기자기한 파스텔 톤 (민트 + 코랄)
COLORS = {
    "bg": "#E8F6F3",
    "card": "#FFFFFF",
    "header": "#5EB7A8",
    "header_text": "#FFFFFF",
    "chip": "#FFE8E1",
    "chip_text": "#C45C48",
    "dept": "#2F5D56",
    "body": "#3A4A47",
    "muted": "#7A8F8A",
    "button": "#5EB7A8",
    "button_hover": "#4AA496",
    "empty": "#F7FBFA",
    "shadow": "#D5EBE6",
}

WEEKDAY_KR = ("월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일")
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) TodaySchedulePopup/1.0"

# 둥글고 부드러운 글씨체 우선 사용
_FONT_CANDIDATES = (
    "나눔스퀘어라운드 Regular",
    "나눔스퀘어라운드",
    "한컴 말랑말랑 Regular",
    "한컴 말랑말랑",
    "나눔스퀘어",
    "Noto Sans KR",
    "맑은 고딕",
)
_FONT_BOLD_CANDIDATES = (
    "나눔스퀘어라운드 Bold",
    "나눔스퀘어라운드 ExtraBold",
    "한컴 말랑말랑 Bold",
    "나눔스퀘어 Bold",
    "Noto Sans KR Medium",
    "맑은 고딕",
)

FONT_UI = "맑은 고딕"
FONT_UI_BOLD = "맑은 고딕"


def _resolve_fonts() -> None:
    global FONT_UI, FONT_UI_BOLD
    try:
        from tkinter import font as tkfont

        available = set(tkfont.families())
        for name in _FONT_CANDIDATES:
            if name in available:
                FONT_UI = name
                break
        for name in _FONT_BOLD_CANDIDATES:
            if name in available:
                FONT_UI_BOLD = name
                break
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 스프레드시트 연동
# ---------------------------------------------------------------------------
def _candidate_sheet_names() -> list[str]:
    names = [f"2학기 {i}주" for i in range(1, 23)]
    names += [f"{i}주" for i in range(1, 22)]
    return names


def _load_cache() -> dict:
    if not CACHE_PATH.exists():
        return {}
    try:
        return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _save_cache(sheet_name: str) -> None:
    try:
        CACHE_PATH.write_text(
            json.dumps({"last_sheet": sheet_name}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass


def fetch_sheet_csv(sheet_name: str | None = None, gid: str | None = None) -> list[list[str]]:
    params: dict[str, str] = {"tqx": "out:csv"}
    if sheet_name:
        params["sheet"] = sheet_name
    if gid:
        params["gid"] = gid
    query = urllib.parse.urlencode(params)
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?{query}"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8-sig", errors="replace")
    return list(csv.reader(io.StringIO(raw)))


def clean_cell(value: str) -> str:
    text = (value or "").replace("\r", "\n").strip()
    text = re.sub(r"\n+", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip(" -")


def split_tasks(text: str) -> list[str]:
    text = clean_cell(text)
    if not text:
        return []
    # "-항목 -항목" 형태 우선 분리
    parts = [p.strip() for p in re.split(r"\s*-\s*", text) if p.strip()]
    if len(parts) > 1:
        return parts
    return [text]


def find_date_column(header_row: list[str], target: date) -> int | None:
    needle = f"{target.year}년 {target.month}월 {target.day}일"
    for idx, cell in enumerate(header_row):
        if needle in clean_cell(cell):
            return idx
    return None


def find_header_row(rows: list[list[str]], target: date) -> tuple[int, int] | None:
    needle = f"{target.year}년 {target.month}월 {target.day}일"
    for r_idx, row in enumerate(rows[:8]):
        for c_idx, cell in enumerate(row):
            if needle in clean_cell(cell):
                return r_idx, c_idx
    return None


def extract_today_items(rows: list[list[str]], target: date) -> tuple[str, list[tuple[str, list[str]]]] | None:
    found = find_header_row(rows, target)
    if not found:
        return None
    header_idx, col_idx = found
    header_label = clean_cell(rows[header_idx][col_idx])

    items: list[tuple[str, list[str]]] = []
    for row in rows[header_idx + 1 :]:
        if not row:
            continue
        dept = clean_cell(row[0] if len(row) > 0 else "")
        if not dept or "부서" in dept:
            continue
        if dept in {"토의 안건"}:
            # 안건도 보여 주되, 내용 없으면 생략
            pass
        cell = row[col_idx] if col_idx < len(row) else ""
        tasks = split_tasks(cell)
        if tasks:
            items.append((dept, tasks))
    return header_label, items


def ordered_sheet_candidates() -> list[str]:
    cache = _load_cache()
    last = cache.get("last_sheet")
    base = _candidate_sheet_names()
    ordered: list[str] = []
    if last:
        ordered.append(last)
        # 이웃 주도 먼저 시도
        if last in base:
            i = base.index(last)
            for j in (i - 1, i + 1, i - 2, i + 2):
                if 0 <= j < len(base):
                    ordered.append(base[j])
    for name in base:
        if name not in ordered:
            ordered.append(name)
    return ordered


def load_today_schedule(target: date | None = None) -> dict:
    target = target or date.today()
    errors: list[str] = []

    for sheet_name in ordered_sheet_candidates():
        try:
            rows = fetch_sheet_csv(sheet_name=sheet_name)
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            errors.append(f"{sheet_name}: {exc}")
            continue

        parsed = extract_today_items(rows, target)
        if not parsed:
            continue

        header_label, items = parsed
        _save_cache(sheet_name)
        return {
            "ok": True,
            "sheet_name": sheet_name,
            "date_label": header_label,
            "items": items,
            "target": target,
        }

    return {
        "ok": False,
        "sheet_name": None,
        "date_label": None,
        "items": [],
        "target": target,
        "error": "오늘 날짜가 들어 있는 주간 시트를 찾지 못했습니다."
        + (("\n" + "\n".join(errors[:3])) if errors else ""),
    }


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
class TodaySchedulePopup(ctk.CTk):
    def __init__(self, payload: dict):
        super().__init__()
        self.payload = payload
        self._drag_x = 0
        self._drag_y = 0

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")
        _resolve_fonts()

        self.title("오늘의 주간 업무")
        self.geometry("540x720")
        self.minsize(500, 560)
        self.configure(fg_color=COLORS["bg"])
        self.attributes("-topmost", True)

        # 둥근 카드 느낌의 외곽
        self.outer = ctk.CTkFrame(
            self,
            fg_color=COLORS["shadow"],
            corner_radius=28,
            border_width=0,
        )
        self.outer.pack(fill="both", expand=True, padx=14, pady=14)

        self.card = ctk.CTkFrame(
            self.outer,
            fg_color=COLORS["card"],
            corner_radius=24,
            border_width=0,
        )
        self.card.pack(fill="both", expand=True, padx=6, pady=6)

        self._build_header()
        self._build_body()
        self._build_footer()

        self.bind("<Escape>", lambda _e: self.close_app())
        self.protocol("WM_DELETE_WINDOW", self.close_app)

        # 화면 중앙
        self.after(30, self._center_window)

    def _center_window(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = max((sw - w) // 2, 0), max((sh - h) // 5, 40)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self.card, fg_color=COLORS["header"], corner_radius=20)
        header.pack(fill="x", padx=18, pady=(18, 12))
        header.bind("<ButtonPress-1>", self._start_drag)
        header.bind("<B1-Motion>", self._on_drag)

        target: date = self.payload.get("target") or date.today()
        subtitle = f"{target.year}년 {target.month}월 {target.day}일 {WEEKDAY_KR[target.weekday()]}"
        if self.payload.get("date_label"):
            subtitle = self.payload["date_label"]

        top = ctk.CTkFrame(header, fg_color="transparent")
        top.pack(fill="x", padx=18, pady=(14, 2))

        ctk.CTkLabel(
            top,
            text="오늘의 주간 업무 알림",
            font=ctk.CTkFont(family=FONT_UI_BOLD, size=16, weight="bold"),
            text_color=COLORS["header_text"],
            anchor="w",
        ).pack(side="left")

        close_btn = ctk.CTkButton(
            top,
            text="✕",
            width=34,
            height=34,
            corner_radius=17,
            fg_color="#FFFFFF",
            hover_color="#FFE8E1",
            text_color=COLORS["chip_text"],
            font=ctk.CTkFont(family=FONT_UI_BOLD, size=14, weight="bold"),
            command=self.close_app,
        )
        close_btn.pack(side="right")

        ctk.CTkLabel(
            header,
            text=subtitle,
            font=ctk.CTkFont(family=FONT_UI, size=15),
            text_color="#F3FFFC",
            anchor="w",
        ).pack(fill="x", padx=18, pady=(2, 14))

    def _build_body(self) -> None:
        body = ctk.CTkScrollableFrame(
            self.card,
            fg_color=COLORS["empty"],
            corner_radius=18,
            scrollbar_button_color=COLORS["header"],
            scrollbar_button_hover_color=COLORS["button_hover"],
        )
        body.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        if not self.payload.get("ok"):
            self._add_message(
                body,
                "일정을 불러오지 못했어요",
                self.payload.get("error") or "잠시 후 다시 시도해 주세요.",
            )
            return

        items: list[tuple[str, list[str]]] = self.payload.get("items") or []
        if not items:
            self._add_message(
                body,
                "오늘은 등록된 업무가 없어요",
                "주간 업무 계획표의 오늘 날짜 열이 비어 있습니다.",
            )
            return

        count_chip = ctk.CTkLabel(
            body,
            text=f"  부서 {len(items)}곳의 일정  ",
            font=ctk.CTkFont(family=FONT_UI_BOLD, size=13, weight="bold"),
            text_color=COLORS["chip_text"],
            fg_color=COLORS["chip"],
            corner_radius=12,
        )
        count_chip.pack(anchor="w", padx=8, pady=(10, 6))

        for dept, tasks in items:
            self._add_dept_card(body, dept, tasks)

    def _add_message(self, parent, title: str, detail: str) -> None:
        box = ctk.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=18)
        box.pack(fill="x", padx=8, pady=16)
        ctk.CTkLabel(
            box,
            text=title,
            font=ctk.CTkFont(family=FONT_UI_BOLD, size=17, weight="bold"),
            text_color=COLORS["dept"],
        ).pack(padx=16, pady=(18, 6))
        ctk.CTkLabel(
            box,
            text=detail,
            font=ctk.CTkFont(family=FONT_UI, size=14),
            text_color=COLORS["muted"],
            wraplength=420,
            justify="left",
        ).pack(padx=16, pady=(0, 18))

    def _add_dept_card(self, parent, dept: str, tasks: list[str]) -> None:
        card = ctk.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=18)
        card.pack(fill="x", padx=8, pady=6)

        badge = ctk.CTkLabel(
            card,
            text=f"  {dept}  ",
            font=ctk.CTkFont(family=FONT_UI_BOLD, size=14, weight="bold"),
            text_color=COLORS["chip_text"],
            fg_color=COLORS["chip"],
            corner_radius=12,
            anchor="w",
        )
        badge.pack(anchor="w", padx=14, pady=(12, 6))

        for task in tasks:
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=2)
            ctk.CTkLabel(
                row,
                text="●",
                font=ctk.CTkFont(size=10),
                text_color=COLORS["header"],
                width=18,
            ).pack(side="left", anchor="n", pady=3)
            ctk.CTkLabel(
                row,
                text=task,
                font=ctk.CTkFont(family=FONT_UI, size=14),
                text_color=COLORS["body"],
                wraplength=420,
                justify="left",
                anchor="w",
            ).pack(side="left", fill="x", expand=True)

        ctk.CTkFrame(card, fg_color="transparent", height=10).pack()

    def _build_footer(self) -> None:
        footer = ctk.CTkFrame(self.card, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(4, 16))

        sheet = self.payload.get("sheet_name")
        hint = f"시트: {sheet}" if sheet else "Google 스프레드시트 연동"
        ctk.CTkLabel(
            footer,
            text=hint,
            font=ctk.CTkFont(family=FONT_UI, size=12),
            text_color=COLORS["muted"],
            anchor="w",
        ).pack(side="left")

        ctk.CTkButton(
            footer,
            text="확인했어요",
            width=132,
            height=44,
            corner_radius=22,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color="white",
            font=ctk.CTkFont(family=FONT_UI_BOLD, size=15, weight="bold"),
            command=self.close_app,
        ).pack(side="right")

    def _start_drag(self, event) -> None:
        self._drag_x = event.x_root - self.winfo_x()
        self._drag_y = event.y_root - self.winfo_y()

    def _on_drag(self, event) -> None:
        x = event.x_root - self._drag_x
        y = event.y_root - self._drag_y
        self.geometry(f"+{x}+{y}")

    def close_app(self) -> None:
        self.destroy()


def show_loading_then_popup() -> None:
    # 짧은 로딩 창
    boot = ctk.CTk()
    ctk.set_appearance_mode("light")
    _resolve_fonts()
    boot.title("불러오는 중")
    boot.geometry("360x170")
    boot.configure(fg_color=COLORS["bg"])
    boot.attributes("-topmost", True)
    frame = ctk.CTkFrame(boot, fg_color=COLORS["card"], corner_radius=22)
    frame.pack(fill="both", expand=True, padx=16, pady=16)
    ctk.CTkLabel(
        frame,
        text="오늘의 일정을 불러오는 중…",
        font=ctk.CTkFont(family=FONT_UI_BOLD, size=16, weight="bold"),
        text_color=COLORS["dept"],
    ).pack(expand=True)
    boot.update()

    payload = load_today_schedule()
    boot.destroy()

    app = TodaySchedulePopup(payload)
    app.mainloop()


if __name__ == "__main__":
    # 시작 프로그램으로 켜질 때는 네트워크/데스크톱이 준비될 때까지 잠시 기다립니다.
    if "--startup" in sys.argv:
        time.sleep(12)
    show_loading_then_popup()
