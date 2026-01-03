# -*- coding: utf-8 -*-
"""
DropZone - ファイルドロップエリアコンポーネント
"""
import tkinter as tk
from tkinter import ttk, filedialog
from pathlib import Path
from ...config import COLORS, FONTS


class DropZone(ttk.Frame):
    """ファイルドロップエリア"""

    def __init__(self, parent, on_file_selected=None, filetypes=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.on_file_selected = on_file_selected
        self.filetypes = filetypes or [("PowerPoint", "*.pptx"), ("All files", "*.*")]

        self._create_widgets()
        self._setup_dnd()

    def _create_widgets(self):
        """ウィジェットを作成"""
        # ドロップゾーンフレーム
        self.drop_frame = tk.Frame(
            self,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=2,
            highlightcolor=COLORS["primary"]
        )
        self.drop_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # 内側のコンテンツ
        inner = tk.Frame(self.drop_frame, bg=COLORS["surface"])
        inner.place(relx=0.5, rely=0.5, anchor="center")

        # アイコン
        icon_label = tk.Label(
            inner,
            text="📂",
            font=(FONTS["title"][0], 48),
            bg=COLORS["surface"]
        )
        icon_label.pack(pady=(0, 10))

        # メインテキスト
        main_text = tk.Label(
            inner,
            text="ファイルをドラッグ＆ドロップ",
            font=FONTS["heading"],
            fg=COLORS["text"],
            bg=COLORS["surface"]
        )
        main_text.pack()

        # サブテキスト
        sub_text = tk.Label(
            inner,
            text="または",
            font=FONTS["small"],
            fg=COLORS["text_muted"],
            bg=COLORS["surface"]
        )
        sub_text.pack(pady=10)

        # ファイル選択ボタン
        select_btn = tk.Button(
            inner,
            text="📁 ファイルを選択",
            font=FONTS["body"],
            bg=COLORS["primary"],
            fg="white",
            activebackground=COLORS["primary_dark"],
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=10,
            cursor="hand2",
            command=self._select_file
        )
        select_btn.pack(pady=(0, 15))

        # フォルダ一括処理ボタン
        batch_btn = tk.Button(
            inner,
            text="📂 フォルダー一括処理",
            font=FONTS["small"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text_secondary"],
            activebackground=COLORS["border"],
            activeforeground=COLORS["text"],
            relief="flat",
            padx=15,
            pady=5,
            cursor="hand2",
            command=self._select_folder
        )
        batch_btn.pack()

        # 対応形式の説明
        format_text = tk.Label(
            inner,
            text="対応形式: .pptx",
            font=FONTS["small"],
            fg=COLORS["text_muted"],
            bg=COLORS["surface"]
        )
        format_text.pack(pady=(20, 0))

    def _setup_dnd(self):
        """ドラッグ＆ドロップの設定"""
        # tkinterDnDがインストールされている場合のみ有効
        try:
            self.drop_frame.drop_target_register('DND_Files')
            self.drop_frame.dnd_bind('<<Drop>>', self._on_drop)
            self.drop_frame.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            self.drop_frame.dnd_bind('<<DragLeave>>', self._on_drag_leave)
        except:
            # tkinterDnDがない場合はスキップ
            pass

    def _on_drop(self, event):
        """ファイルドロップ時"""
        files = self._parse_drop_data(event.data)
        if files:
            self._process_file(files[0])
        self._reset_highlight()

    def _on_drag_enter(self, event):
        """ドラッグ開始時"""
        self.drop_frame.config(highlightbackground=COLORS["primary"])

    def _on_drag_leave(self, event):
        """ドラッグ終了時"""
        self._reset_highlight()

    def _reset_highlight(self):
        """ハイライトをリセット"""
        self.drop_frame.config(highlightbackground=COLORS["border"])

    def _parse_drop_data(self, data: str) -> list:
        """ドロップデータをパース"""
        # Windows/Macのパス形式に対応
        if data.startswith('{'):
            # 複数ファイルの場合
            files = data.strip('{}').split('} {')
        else:
            files = data.split()

        return [f.strip() for f in files if f.strip().lower().endswith('.pptx')]

    def _select_file(self):
        """ファイル選択ダイアログ"""
        file_path = filedialog.askopenfilename(
            title="PowerPointファイルを選択",
            filetypes=self.filetypes
        )
        if file_path:
            self._process_file(file_path)

    def _select_folder(self):
        """フォルダ選択ダイアログ"""
        folder_path = filedialog.askdirectory(
            title="処理するフォルダを選択"
        )
        if folder_path:
            # フォルダ内のpptxファイルを検索
            pptx_files = list(Path(folder_path).glob("*.pptx"))
            if pptx_files:
                # 最初のファイルを処理（将来的にはバッチ処理に対応）
                self._process_file(str(pptx_files[0]))
            else:
                tk.messagebox.showwarning(
                    "ファイルなし",
                    "選択したフォルダにPowerPointファイルが見つかりませんでした。"
                )

    def _process_file(self, file_path: str):
        """ファイルを処理"""
        if self.on_file_selected:
            self.on_file_selected(file_path)
