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
    "preview": "#0984E3",
    "danger": "#E17055",
}

FONT = ("맑은 고딕", 10)
FONT_BOLD = ("맑은 고딕", 10, "bold")
FONT_TITLE = ("맑은 고딕", 14, "bold")
LEFT_PANEL_WIDTH = 340


class FileRenameApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("파일명 일괄 변경")
        self.root.configure(bg=COLORS["bg"])
        self.root.minsize(900, 600)
        self.root.geometry("1000x680")

        self.folder_path = tk.StringVar()
        self.find_text = tk.StringVar()
        self.replace_text = tk.StringVar()
        self.match_case = tk.BooleanVar(value=False)
        self.whole_name = tk.BooleanVar(value=False)

        self.files: list[str] = []
        self._find_matches: list[str] = []
        self._find_index = -1
        self._editing_item: str | None = None
        self._edit_entry: tk.Entry | None = None

        self.find_text.trace_add("write", lambda *_: self._update_preview_column())
        self.replace_text.trace_add("write", lambda *_: self._update_preview_column())
        self.match_case.trace_add("write", lambda *_: self._update_preview_column())
        self.whole_name.trace_add("write", lambda *_: self._update_preview_column())

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=FONT)

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
            padding=(10, 6),
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

    def _card(self, parent, **pack_kw) -> tk.Frame:
        outer = tk.Frame(parent, bg=COLORS["border"], padx=1, pady=1)
        outer.pack(**pack_kw)
        inner = tk.Frame(outer, bg=COLORS["card"], padx=16, pady=14)
        inner.pack(fill="both", expand=True)
        return inner

    def _entry(self, parent, textvariable: tk.Variable, width: int | None = None) -> tk.Entry:
        kw = dict(
            textvariable=textvariable,
            font=FONT,
            bg=COLORS["list_bg"],
            fg=COLORS["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["primary"],
        )
        if width:
            kw["width"] = width
        return tk.Entry(parent, **kw)

    def _build_ui(self):
        container = tk.Frame(self.root, bg=COLORS["bg"], padx=24, pady=20)
        container.pack(fill="both", expand=True)

        # 헤더
        header = tk.Frame(container, bg=COLORS["bg"])
        header.pack(fill="x", pady=(0, 16))
        tk.Label(header, text="📁 파일명 일괄 변경", bg=COLORS["bg"], fg=COLORS["text"], font=FONT_TITLE).pack(
            side="left"
        )
        tk.Label(
            header,
            text="오른쪽에서 파일을 선택한 뒤, 왼쪽에서 찾아 바꾸기를 실행하세요",
            bg=COLORS["bg"],
            fg=COLORS["text_muted"],
            font=FONT,
        ).pack(side="left", padx=(12, 0))

        # 좌우 분할 본문
        body = tk.Frame(container, bg=COLORS["bg"])
        body.pack(fill="both", expand=True)

        # ── 왼쪽: 찾아 바꾸기 ──
        left_panel = tk.Frame(body, bg=COLORS["bg"], width=LEFT_PANEL_WIDTH)
        left_panel.pack(side="left", fill="y", padx=(0, 16))
        left_panel.pack_propagate(False)

        replace_card = self._card(left_panel, fill="both", expand=True)

        tk.Label(replace_card, text="🔍 찾아 바꾸기", bg=COLORS["card"], fg=COLORS["text"], font=FONT_BOLD).pack(
            anchor="w"
        )
        tk.Label(
            replace_card,
            text="선택한 파일에만 적용됩니다",
            bg=COLORS["card"],
            fg=COLORS["text_muted"],
            font=("맑은 고딕", 9),
        ).pack(anchor="w", pady=(2, 12))

        form = tk.Frame(replace_card, bg=COLORS["card"])
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        tk.Label(form, text="찾을 내용", bg=COLORS["card"], fg=COLORS["text_muted"], font=FONT, anchor="w").grid(
            row=0, column=0, sticky="w", pady=4
        )
        self.find_entry = self._entry(form, self.find_text)
        self.find_entry.grid(row=0, column=1, sticky="ew", ipady=7, pady=4)

        tk.Label(form, text="바꿀 내용", bg=COLORS["card"], fg=COLORS["text_muted"], font=FONT, anchor="w").grid(
            row=1, column=0, sticky="w", pady=4
        )
        self.replace_entry = self._entry(form, self.replace_text)
        self.replace_entry.grid(row=1, column=1, sticky="ew", ipady=7, pady=4)

        option_row = tk.Frame(replace_card, bg=COLORS["card"])
        option_row.pack(fill="x", pady=(10, 12))
        tk.Checkbutton(
            option_row,
            text="대/소문자 구분",
            variable=self.match_case,
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=FONT,
            activebackground=COLORS["card"],
            selectcolor=COLORS["list_bg"],
        ).pack(anchor="w", pady=2)
        tk.Checkbutton(
            option_row,
            text="전체 파일명 일치",
            variable=self.whole_name,
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=FONT,
            activebackground=COLORS["card"],
            selectcolor=COLORS["list_bg"],
        ).pack(anchor="w", pady=2)

        btn_row = tk.Frame(replace_card, bg=COLORS["card"])
        btn_row.pack(fill="x")

        ttk.Button(btn_row, text="다음 찾기", style="Ghost.TButton", command=self._find_next).pack(
            fill="x", pady=(0, 6)
        )
        ttk.Button(btn_row, text="바꾸기", style="Ghost.TButton", command=self._replace_one).pack(
            fill="x", pady=(0, 6)
        )
        ttk.Button(btn_row, text="선택 항목 모두 바꾸기", style="Accent.TButton", command=self._replace_all).pack(
            fill="x", pady=(0, 6)
        )
        ttk.Button(btn_row, text="초기화", style="Ghost.TButton", command=self._reset_find).pack(fill="x")

        tk.Frame(replace_card, bg=COLORS["border"], height=1).pack(fill="x", pady=14)

        self.selection_label = tk.Label(
            replace_card,
            text="선택된 파일: 없음",
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=FONT_BOLD,
            anchor="w",
        )
        self.selection_label.pack(fill="x")

        self.find_status = tk.Label(
            replace_card,
            text="오른쪽 목록에서 파일을 클릭해 선택하세요.\n(Ctrl·Shift로 여러 파일 선택 가능)",
            bg=COLORS["card"],
            fg=COLORS["text_muted"],
            font=("맑은 고딕", 9),
            anchor="w",
            justify="left",
        )
        self.find_status.pack(fill="x", pady=(8, 0))

        # ── 오른쪽: 폴더 + 파일 목록 ──
        right_panel = tk.Frame(body, bg=COLORS["bg"])
        right_panel.pack(side="left", fill="both", expand=True)

        folder_card = self._card(right_panel, fill="x", pady=(0, 12))
        tk.Label(folder_card, text="폴더 경로", bg=COLORS["card"], fg=COLORS["text_muted"], font=FONT).pack(
            anchor="w"
        )
        path_row = tk.Frame(folder_card, bg=COLORS["card"])
        path_row.pack(fill="x", pady=(6, 0))
        path_row.columnconfigure(0, weight=1)

        self.path_entry = self._entry(path_row, self.folder_path)
        self.path_entry.config(state="readonly")
        self.path_entry.grid(row=0, column=0, sticky="ew", ipady=8, padx=(0, 10))

        ttk.Button(path_row, text="폴더 불러오기", style="Primary.TButton", command=self._load_folder).grid(
            row=0, column=1
        )

        list_card = self._card(right_panel, fill="both", expand=True)

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
            text="클릭으로 선택 · 더블클릭/F2로 직접 수정 · 선택한 파일만 찾아 바꾸기 적용",
            bg=COLORS["card"],
            fg=COLORS["text_muted"],
            font=("맑은 고딕", 9),
        ).pack(anchor="w", pady=(0, 8))

        tree_frame = tk.Frame(list_card, bg=COLORS["border"], padx=1, pady=1)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("name", "preview"),
            show="headings",
            selectmode="extended",
        )
        self.tree.heading("name", text="현재 파일명")
        self.tree.heading("preview", text="변경 후 (미리보기)")
        self.tree.column("name", width=280, anchor="w", stretch=True)
        self.tree.column("preview", width=280, anchor="w", stretch=True)

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.tag_configure("match", foreground=COLORS["preview"])
        self.tree.tag_configure("selected_match", foreground=COLORS["accent"])

        self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_selection_changed())
        self.tree.bind("<Double-1>", self._start_inline_edit)
        self.tree.bind("<F2>", self._start_inline_edit)
        self.root.bind("<Escape>", self._cancel_inline_edit)
        self.root.bind("<Control-h>", lambda e: self.find_entry.focus_set())
        self.root.bind("<Control-H>", lambda e: self.find_entry.focus_set())
        self.find_entry.bind("<Return>", lambda e: self._find_next())
        self.replace_entry.bind("<Return>", lambda e: self._replace_one())

    def _get_selected_files(self) -> list[str]:
        return [f for f in self.tree.selection() if f in self.files]

    def _require_selection(self) -> list[str] | None:
        selected = self._get_selected_files()
        if not selected:
            messagebox.showinfo("안내", "오른쪽 목록에서 파일을 먼저 선택해 주세요.")
            return None
        return selected

    def _on_selection_changed(self):
        self._reset_find()
        selected = self._get_selected_files()
        if not selected:
            self.selection_label.config(text="선택된 파일: 없음", fg=COLORS["text_muted"])
        elif len(selected) == 1:
            self.selection_label.config(text=f"선택된 파일: {selected[0]}", fg=COLORS["primary"])
        else:
            self.selection_label.config(text=f"선택된 파일: {len(selected)}개", fg=COLORS["primary"])
        self._update_preview_column()

    def _load_folder(self):
        path = filedialog.askdirectory(title="폴더 선택")
        if not path:
            return
        self.folder_path.set(path)
        self._reset_find()
        self._refresh_file_list()

    def _refresh_file_list(self):
        self._cancel_inline_edit()
        self.tree.delete(*self.tree.get_children())
        self.files.clear()

        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            self.count_label.config(text="0개")
            self._on_selection_changed()
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
                self.tree.insert("", "end", iid=name, values=(name, ""))

        self.count_label.config(text=f"{len(self.files)}개")
        self._on_selection_changed()

    def _apply_replace(self, name: str, find: str, replace: str) -> str | None:
        if not find:
            return None

        if self.whole_name.get():
            if self.match_case.get():
                if name != find:
                    return None
            elif name.lower() != find.lower():
                return None
            return replace

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

    def _get_matches(self, scope: list[str] | None = None) -> list[str]:
        find = self.find_text.get()
        if not find:
            return []

        targets = scope if scope is not None else self._get_selected_files()
        replace = self.replace_text.get()
        matches = []
        for name in targets:
            new_name = self._apply_replace(name, find, replace)
            if new_name is not None and new_name != name:
                matches.append(name)
        return matches

    def _update_preview_column(self):
        find = self.find_text.get()
        replace = self.replace_text.get()
        selected_set = set(self._get_selected_files())

        for name in self.files:
            if name not in self.tree.get_children():
                continue

            if name not in selected_set:
                self.tree.item(name, values=(name, ""), tags=())
                continue

            if not find:
                self.tree.item(name, values=(name, ""), tags=())
                continue

            new_name = self._apply_replace(name, find, replace)
            if new_name is not None and new_name != name:
                self.tree.item(name, values=(name, new_name), tags=("selected_match",))
            else:
                self.tree.item(name, values=(name, "—"), tags=())

        self._update_status_text()

    def _update_status_text(self):
        find = self.find_text.get()
        selected = self._get_selected_files()

        if not selected:
            self.find_status.config(
                text="오른쪽 목록에서 파일을 클릭해 선택하세요.\n(Ctrl·Shift로 여러 파일 선택 가능)",
                fg=COLORS["text_muted"],
            )
            return

        if not find:
            self.find_status.config(
                text=f"{len(selected)}개 파일 선택됨 · 찾을 내용을 입력하세요.",
                fg=COLORS["text_muted"],
            )
            return

        matches = self._get_matches(selected)
        self.find_status.config(
            text=f"선택 {len(selected)}개 중 {len(matches)}개 일치 · 다음 찾기/바꾸기는 선택 파일만 대상",
            fg=COLORS["preview"] if matches else COLORS["text_muted"],
        )

    def _reset_find(self):
        self._find_matches = []
        self._find_index = -1

    def _find_next(self):
        find = self.find_text.get()
        if not find:
            messagebox.showinfo("안내", "찾을 내용을 입력해 주세요.")
            self.find_entry.focus_set()
            return

        selected = self._require_selection()
        if selected is None:
            return

        self._find_matches = self._get_matches(selected)
        if not self._find_matches:
            self._find_index = -1
            messagebox.showinfo("찾기", f"선택한 파일에서 '{find}'을(를) 찾을 수 없습니다.")
            return

        self._find_index = (self._find_index + 1) % len(self._find_matches)
        target = self._find_matches[self._find_index]

        self.tree.selection_set(target)
        self.tree.focus(target)
        self.tree.see(target)

        self.find_status.config(
            text=f"선택 파일 내 {self._find_index + 1} / {len(self._find_matches)} — '{target}'",
            fg=COLORS["preview"],
        )

    def _current_target(self) -> str | None:
        selected = self._get_selected_files()
        if not selected:
            return None

        tree_selection = self.tree.selection()
        if tree_selection and tree_selection[0] in self._find_matches:
            return tree_selection[0]

        if self._find_matches and 0 <= self._find_index < len(self._find_matches):
            return self._find_matches[self._find_index]

        return selected[0]

    def _replace_one(self):
        find = self.find_text.get()
        if not find:
            messagebox.showinfo("안내", "찾을 내용을 입력해 주세요.")
            self.find_entry.focus_set()
            return

        selected = self._require_selection()
        if selected is None:
            return

        if not self._find_matches:
            self._find_matches = self._get_matches(selected)
            self._find_index = -1

        target = self._current_target()
        if not target:
            self._find_next()
            target = self._current_target()
            if not target:
                return

        if target not in selected:
            messagebox.showinfo("안내", "선택된 파일에서만 바꾸기를 실행할 수 있습니다.")
            return

        old_name = target
        new_name = self._apply_replace(old_name, find, self.replace_text.get())
        if new_name is None or new_name == old_name:
            messagebox.showinfo("안내", f"'{old_name}'에서 '{find}'을(를) 찾을 수 없습니다.")
            return

        if not self._rename_file(old_name, new_name):
            return

        updated_selection = [new_name if f == old_name else f for f in selected]
        self._find_matches = self._get_matches(updated_selection)

        if self._find_matches:
            if self._find_index >= len(self._find_matches):
                self._find_index = 0
            next_target = self._find_matches[self._find_index]
            self.tree.selection_set(*updated_selection)
            self.tree.focus(next_target)
            self.tree.see(next_target)
            self.find_status.config(
                text=f"변경 완료 · 선택 파일 내 {self._find_index + 1} / {len(self._find_matches)} — '{next_target}'",
                fg=COLORS["accent"],
            )
        else:
            self._find_index = -1
            self.tree.selection_set(*updated_selection)
            self.find_status.config(text="선택한 파일에서 모든 일치 항목을 변경했습니다.", fg=COLORS["accent"])

        self._on_selection_changed()

    def _replace_all(self):
        find = self.find_text.get()
        if not find:
            messagebox.showinfo("안내", "찾을 내용을 입력해 주세요.")
            self.find_entry.focus_set()
            return

        selected = self._require_selection()
        if selected is None:
            return

        replace = self.replace_text.get()
        changes: list[tuple[str, str]] = []

        for name in selected:
            new_name = self._apply_replace(name, find, replace)
            if new_name is not None and new_name != name:
                changes.append((name, new_name))

        if not changes:
            messagebox.showinfo("안내", f"선택한 파일에서 '{find}'과(와) 일치하는 파일명이 없습니다.")
            return

        preview = "\n".join(f"  {old}  →  {new}" for old, new in changes[:15])
        if len(changes) > 15:
            preview += f"\n  ... 외 {len(changes) - 15}개"

        if not messagebox.askyesno(
            "선택 항목 모두 바꾸기",
            f"선택한 {len(selected)}개 파일 중 {len(changes)}개를 변경합니다.\n\n{preview}\n\n계속하시겠습니까?",
        ):
            return

        success = 0
        new_selection = list(selected)
        for old_name, new_name in changes:
            if self._rename_file(old_name, new_name):
                success += 1
                new_selection = [new_name if f == old_name else f for f in new_selection]

        self.tree.selection_set(*new_selection)
        self._reset_find()
        self._on_selection_changed()
        messagebox.showinfo("완료", f"선택한 파일 {success}개의 이름이 변경되었습니다.")

    def _start_inline_edit(self, event=None):
        if event and event.type == "4":
            if self.tree.identify_region(event.x, event.y) != "cell":
                return
            if self.tree.identify_column(event.x) != "#1":
                return

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

        children = self.tree.get_children()
        pos = children.index(old_name) if old_name in children else idx
        preview = self.tree.item(old_name, "values")[1] if old_name in children else ""
        self.tree.delete(old_name)
        self.tree.insert("", pos, iid=new_name, values=(new_name, preview))

        return True


def main():
    root = tk.Tk()
    FileRenameApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
