import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 색상 (파스텔톤, 심플)
COLORS = {
    "bg": "#F5F6FA",
    "card": "#FFFFFF",
    "primary": "#6C5CE7",
    "primary_hover": "#5A4BD1",
    "accent": "#00B894",
    "accent_hover": "#00A383",
    "text": "#2D3436",
    "text_muted": "#636E72",
    "border": "#DFE6E9",
    "list_bg": "#FAFBFC",
    "list_select": "#E8E4FF",
    "danger": "#E17055",
}

FONT = ("맑은 고딕", 10)
FONT_BOLD = ("맑은 고딕", 10, "bold")
FONT_TITLE = ("맑은 고딕", 14, "bold")


class FileRenameApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("파일명 일괄 변경")
        self.root.configure(bg=COLORS["bg"])
        self.root.minsize(720, 560)
        self.root.geometry("820x620")

        self.folder_path = tk.StringVar()
        self.find_text = tk.StringVar()
        self.replace_text = tk.StringVar()
        self.match_case = tk.BooleanVar(value=False)

        self.files: list[str] = []
        self._editing_item: str | None = None
        self._edit_entry: tk.Entry | None = None

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["card"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=FONT)
        style.configure("Card.TLabel", background=COLORS["card"], foreground=COLORS["text"], font=FONT)
        style.configure("Muted.TLabel", background=COLORS["card"], foreground=COLORS["text_muted"], font=FONT)
        style.configure("Title.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=FONT_TITLE)
        style.configure("TCheckbutton", background=COLORS["card"], foreground=COLORS["text"], font=FONT)

        style.configure(
            "Primary.TButton",
            background=COLORS["primary"],
            foreground="white",
            font=FONT_BOLD,
            padding=(14, 8),
            borderwidth=0,
        )
        style.map("Primary.TButton", background=[("active", COLORS["primary_hover"])])

        style.configure(
            "Accent.TButton",
            background=COLORS["accent"],
            foreground="white",
            font=FONT,
            padding=(12, 7),
            borderwidth=0,
        )
        style.map("Accent.TButton", background=[("active", COLORS["accent_hover"])])

        style.configure(
            "Ghost.TButton",
            background=COLORS["card"],
            foreground=COLORS["text"],
            font=FONT,
            padding=(12, 7),
            borderwidth=1,
        )
        style.map("Ghost.TButton", background=[("active", COLORS["list_bg"])])

        style.configure(
            "Treeview",
            background=COLORS["list_bg"],
            foreground=COLORS["text"],
            fieldbackground=COLORS["list_bg"],
            font=FONT,
            rowheight=30,
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background=COLORS["border"],
            foreground=COLORS["text"],
            font=FONT_BOLD,
            borderwidth=0,
        )
        style.map(
            "Treeview",
            background=[("selected", COLORS["list_select"])],
            foreground=[("selected", COLORS["text"])],
        )

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=24)
        container.pack(fill="both", expand=True)

        # 헤더
        header = ttk.Frame(container)
        header.pack(fill="x", pady=(0, 16))
        ttk.Label(header, text="📁 파일명 일괄 변경", style="Title.TLabel").pack(side="left")
        ttk.Label(
            header,
            text="폴더를 선택하고 파일명을 수정하세요",
            style="TLabel",
            foreground=COLORS["text_muted"],
        ).pack(side="left", padx=(12, 0))

        # 폴더 선택 카드
        folder_card = ttk.Frame(container, style="Card.TFrame", padding=16)
        folder_card.pack(fill="x", pady=(0, 12))
        folder_card.configure(relief="flat")

        folder_inner = tk.Frame(folder_card, bg=COLORS["card"])
        folder_inner.pack(fill="x")

        tk.Label(folder_inner, text="폴더 경로", bg=COLORS["card"], fg=COLORS["text_muted"], font=FONT).pack(
            anchor="w"
        )
        path_row = tk.Frame(folder_inner, bg=COLORS["card"])
        path_row.pack(fill="x", pady=(6, 0))

        self.path_entry = tk.Entry(
            path_row,
            textvariable=self.folder_path,
            font=FONT,
            bg=COLORS["list_bg"],
            fg=COLORS["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["primary"],
            state="readonly",
        )
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))

        ttk.Button(path_row, text="폴더 불러오기", style="Primary.TButton", command=self._load_folder).pack(
            side="right"
        )

        # 파일 목록 카드
        list_card = ttk.Frame(container, style="Card.TFrame", padding=16)
        list_card.pack(fill="both", expand=True, pady=(0, 12))

        list_header = tk.Frame(list_card, bg=COLORS["card"])
        list_header.pack(fill="x", pady=(0, 8))
        tk.Label(list_header, text="파일 목록", bg=COLORS["card"], fg=COLORS["text"], font=FONT_BOLD).pack(
            side="left"
        )
        self.count_label = tk.Label(
            list_header, text="0개", bg=COLORS["card"], fg=COLORS["text_muted"], font=FONT
        )
        self.count_label.pack(side="right")

        tk.Label(
            list_card,
            text="더블클릭 또는 F2로 파일명을 직접 수정할 수 있습니다",
            bg=COLORS["card"],
            fg=COLORS["text_muted"],
            font=("맑은 고딕", 9),
        ).pack(anchor="w", pady=(0, 8))

        tree_frame = tk.Frame(list_card, bg=COLORS["border"], padx=1, pady=1)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=("name",), show="headings", selectmode="extended")
        self.tree.heading("name", text="파일명")
        self.tree.column("name", width=600, anchor="w")

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self._start_inline_edit)
        self.tree.bind("<F2>", self._start_inline_edit)
        self.root.bind("<Escape>", self._cancel_inline_edit)

        # 찾아 바꾸기 카드 (엑셀 스타일)
        replace_card = ttk.Frame(container, style="Card.TFrame", padding=16)
        replace_card.pack(fill="x")

        tk.Label(
            replace_card, text="찾아 바꾸기", bg=COLORS["card"], fg=COLORS["text"], font=FONT_BOLD
        ).pack(anchor="w", pady=(0, 12))

        form = tk.Frame(replace_card, bg=COLORS["card"])
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        def make_row(row: int, label: str, var: tk.Variable, show: str = ""):
            tk.Label(form, text=label, bg=COLORS["card"], fg=COLORS["text_muted"], font=FONT, width=10, anchor="w").grid(
                row=row, column=0, sticky="w", pady=5
            )
            entry = tk.Entry(
                form,
                textvariable=var,
                font=FONT,
                bg=COLORS["list_bg"],
                fg=COLORS["text"],
                relief="flat",
                highlightthickness=1,
                highlightbackground=COLORS["border"],
                highlightcolor=COLORS["primary"],
                show=show,
            )
            entry.grid(row=row, column=1, sticky="ew", ipady=7, padx=(8, 0), pady=5)
            return entry

        make_row(0, "찾을 내용", self.find_text)
        make_row(1, "바꿀 내용", self.replace_text)

        option_row = tk.Frame(replace_card, bg=COLORS["card"])
        option_row.pack(fill="x", pady=(4, 12))
        tk.Checkbutton(
            option_row,
            text="대/소문자 구분",
            variable=self.match_case,
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=FONT,
            activebackground=COLORS["card"],
            selectcolor=COLORS["list_bg"],
        ).pack(side="left")

        btn_row = tk.Frame(replace_card, bg=COLORS["card"])
        btn_row.pack(fill="x")

        ttk.Button(btn_row, text="바꾸기", style="Ghost.TButton", command=self._replace_one).pack(side="left", padx=(0, 8))
        ttk.Button(btn_row, text="모두 바꾸기", style="Accent.TButton", command=self._replace_all).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="미리보기", style="Ghost.TButton", command=self._preview_replace).pack(side="left")

        tk.Label(
            replace_card,
            text="※ 엑셀과 같이 파일명 안의 텍스트를 찾아 바꿉니다. '바꾸기'는 선택한 파일 1개, '모두 바꾸기'는 전체 파일에 적용됩니다.",
            bg=COLORS["card"],
            fg=COLORS["text_muted"],
            font=("맑은 고딕", 9),
            wraplength=760,
            justify="left",
        ).pack(anchor="w", pady=(10, 0))

    def _load_folder(self):
        path = filedialog.askdirectory(title="폴더 선택")
        if not path:
            return
        self.folder_path.set(path)
        self._refresh_file_list()

    def _refresh_file_list(self):
        self._cancel_inline_edit()
        self.tree.delete(*self.tree.get_children())
        self.files.clear()

        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            self.count_label.config(text="0개")
            return

        try:
            entries = sorted(os.listdir(folder))
        except OSError as e:
            messagebox.showerror("오류", f"폴더를 읽을 수 없습니다.\n{e}")
            return

        for name in entries:
            full = os.path.join(folder, name)
            if os.path.isfile(full):
                self.files.append(name)
                self.tree.insert("", "end", iid=name, values=(name,))

        self.count_label.config(text=f"{len(self.files)}개")

    def _start_inline_edit(self, event=None):
        self._cancel_inline_edit()
        selection = self.tree.selection()
        if not selection:
            return

        item_id = selection[0]
        bbox = self.tree.bbox(item_id, "name")
        if not bbox:
            return

        x, y, w, h = bbox
        current_name = self.tree.item(item_id, "values")[0]

        self._editing_item = item_id
        self._edit_entry = tk.Entry(
            self.tree,
            font=FONT,
            bg="white",
            fg=COLORS["text"],
            relief="flat",
            highlightthickness=2,
            highlightbackground=COLORS["primary"],
            highlightcolor=COLORS["primary"],
        )
        self._edit_entry.place(x=x, y=y, width=w, height=h)
        self._edit_entry.insert(0, current_name)
        self._edit_entry.select_range(0, "end")
        self._edit_entry.focus_set()
        self._edit_entry.bind("<Return>", self._commit_inline_edit)
        self._edit_entry.bind("<FocusOut>", self._commit_inline_edit)

    def _cancel_inline_edit(self, event=None):
        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None
        self._editing_item = None

    def _commit_inline_edit(self, event=None):
        if not self._edit_entry or not self._editing_item:
            return

        old_name = self._editing_item
        new_name = self._edit_entry.get().strip()
        self._cancel_inline_edit()

        if not new_name or new_name == old_name:
            return

        self._rename_file(old_name, new_name)

    def _rename_file(self, old_name: str, new_name: str) -> bool:
        folder = self.folder_path.get()
        old_path = os.path.join(folder, old_name)
        new_path = os.path.join(folder, new_name)

        if os.path.exists(new_path):
            messagebox.showwarning("이름 중복", f"'{new_name}' 파일이 이미 존재합니다.")
            return False

        invalid = '<>:"/\\|?*'
        if any(ch in new_name for ch in invalid):
            messagebox.showwarning("잘못된 이름", f"파일명에 사용할 수 없는 문자가 포함되어 있습니다.\n({invalid})")
            return False

        try:
            os.rename(old_path, new_path)
        except OSError as e:
            messagebox.showerror("변경 실패", f"'{old_name}' → '{new_name}'\n{e}")
            return False

        idx = self.files.index(old_name)
        self.files[idx] = new_name
        self.tree.item(old_name, iid=new_name, values=(new_name,))
        return True

    def _apply_replace(self, name: str, find: str, replace: str) -> str | None:
        if self.match_case.get():
            if find not in name:
                return None
            return name.replace(find, replace)
        lower_name = name.lower()
        lower_find = find.lower()
        if lower_find not in lower_name:
            return None

        result = []
        start = 0
        while True:
            idx = lower_name.find(lower_find, start)
            if idx == -1:
                result.append(name[start:])
                break
            result.append(name[start:idx])
            result.append(replace)
            start = idx + len(find)
        return "".join(result)

    def _replace_one(self):
        find = self.find_text.get()
        if not find:
            messagebox.showinfo("안내", "찾을 내용을 입력해 주세요.")
            return

        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("안내", "바꿀 파일을 목록에서 선택해 주세요.")
            return

        old_name = selection[0]
        new_name = self._apply_replace(old_name, find, self.replace_text.get())
        if new_name is None:
            messagebox.showinfo("안내", f"선택한 파일에서 '{find}'을(를) 찾을 수 없습니다.")
            return

        if new_name == old_name:
            return

        self._rename_file(old_name, new_name)

    def _replace_all(self):
        find = self.find_text.get()
        if not find:
            messagebox.showinfo("안내", "찾을 내용을 입력해 주세요.")
            return

        if not self.files:
            messagebox.showinfo("안내", "변경할 파일이 없습니다.")
            return

        replace = self.replace_text.get()
        changes: list[tuple[str, str]] = []

        for name in list(self.files):
            new_name = self._apply_replace(name, find, replace)
            if new_name is not None and new_name != name:
                changes.append((name, new_name))

        if not changes:
            messagebox.showinfo("안내", f"'{find}'과(와) 일치하는 파일명이 없습니다.")
            return

        preview = "\n".join(f"  {old}  →  {new}" for old, new in changes[:15])
        if len(changes) > 15:
            preview += f"\n  ... 외 {len(changes) - 15}개"

        if not messagebox.askyesno(
            "모두 바꾸기 확인",
            f"{len(changes)}개 파일명을 변경합니다.\n\n{preview}\n\n계속하시겠습니까?",
        ):
            return

        success = 0
        for old_name, new_name in changes:
            if self._rename_file(old_name, new_name):
                success += 1

        messagebox.showinfo("완료", f"{success}개 파일명이 변경되었습니다.")

    def _preview_replace(self):
        find = self.find_text.get()
        if not find:
            messagebox.showinfo("안내", "찾을 내용을 입력해 주세요.")
            return

        replace = self.replace_text.get()
        lines = []
        for name in self.files:
            new_name = self._apply_replace(name, find, replace)
            if new_name is not None and new_name != name:
                lines.append(f"{name}  →  {new_name}")

        if not lines:
            messagebox.showinfo("미리보기", f"'{find}'과(와) 일치하는 파일명이 없습니다.")
            return

        preview = "\n".join(lines[:30])
        if len(lines) > 30:
            preview += f"\n... 외 {len(lines) - 30}개"

        messagebox.showinfo("미리보기", f"총 {len(lines)}개 변경 예정\n\n{preview}")


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except tk.TclError:
        pass
    FileRenameApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
