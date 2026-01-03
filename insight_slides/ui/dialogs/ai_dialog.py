# -*- coding: utf-8 -*-
"""
AIDialog - AI処理プロンプト入力ダイアログ
"""
import tkinter as tk
from tkinter import ttk, scrolledtext
from ...config import COLORS, FONTS, AI_PRESETS


class AIDialog:
    """AI処理プロンプト入力ダイアログ"""

    def __init__(self, parent, processor, callback, preset_name=None):
        self.processor = processor
        self.callback = callback
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("AI処理")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 画面中央に配置
        width, height = 550, 420
        screen_w = self.dialog.winfo_screenwidth()
        screen_h = self.dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        self.dialog.minsize(450, 350)

        self._create_widgets()

        # 指定されたプリセットがあれば選択
        if preset_name and preset_name in AI_PRESETS:
            self.preset_var.set(preset_name)
            self._on_preset_change(None)

    def _create_widgets(self):
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)

        # タイトル
        ttk.Label(
            frame,
            text="🤖 AI処理",
            font=FONTS["heading"]
        ).pack(anchor='w', pady=(0, 15))

        # プリセット選択
        ttk.Label(frame, text="プリセット:", font=FONTS["body"]).pack(anchor='w')
        presets = list(AI_PRESETS.keys())
        self.preset_var = tk.StringVar(value=presets[0] if presets else "")
        preset_combo = ttk.Combobox(
            frame,
            textvariable=self.preset_var,
            values=presets,
            width=50,
            state="readonly"
        )
        preset_combo.pack(fill='x', pady=(5, 15))
        preset_combo.bind("<<ComboboxSelected>>", self._on_preset_change)

        # プロンプト入力
        ttk.Label(frame, text="プロンプト:", font=FONTS["body"]).pack(anchor='w')
        self.prompt_text = scrolledtext.ScrolledText(
            frame,
            width=60,
            height=10,
            font=FONTS["small"],
            wrap=tk.WORD
        )
        self.prompt_text.pack(fill='both', expand=True, pady=5)

        # 初期プロンプト設定
        if presets:
            self.prompt_text.insert("1.0", AI_PRESETS[presets[0]])

        # ヒント
        hint_text = ttk.Label(
            frame,
            text="💡 ヒント: プロンプトを編集してカスタム処理も可能です",
            font=FONTS["small"],
            foreground=COLORS["text_muted"]
        )
        hint_text.pack(anchor='w', pady=(5, 0))

        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=15)

        ttk.Button(
            btn_frame,
            text="キャンセル",
            command=self.dialog.destroy
        ).pack(side='left', padx=5)

        execute_btn = tk.Button(
            btn_frame,
            text="▶ 実行",
            font=FONTS["body"],
            bg=COLORS["primary"],
            fg="white",
            relief="flat",
            padx=20,
            cursor="hand2",
            command=self._execute
        )
        execute_btn.pack(side='left', padx=5)

    def _on_preset_change(self, event):
        preset_prompt = AI_PRESETS.get(self.preset_var.get(), "")
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", preset_prompt)

    def _execute(self):
        prompt = self.prompt_text.get("1.0", "end").strip()
        if prompt:
            self.callback(prompt)
            self.dialog.destroy()


class EditDialog:
    """テキスト編集ダイアログ"""

    def __init__(self, parent, title, current_text, callback):
        self.callback = callback
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # サイズ設定
        width, height = 550, 450
        screen_w = self.dialog.winfo_screenwidth()
        screen_h = self.dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        self.dialog.resizable(True, True)
        self.dialog.minsize(400, 350)

        self._create_widgets(current_text)

    def _create_widgets(self, current_text):
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)

        # テキストエリア
        ttk.Label(frame, text="テキスト編集:", font=FONTS["body"]).pack(anchor='w')
        self.text_area = scrolledtext.ScrolledText(
            frame,
            width=60,
            height=15,
            font=FONTS["body"],
            wrap=tk.WORD
        )
        self.text_area.pack(fill='both', expand=True, pady=5)
        self.text_area.insert("1.0", current_text)

        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=15)

        ttk.Button(
            btn_frame,
            text="キャンセル",
            command=self.dialog.destroy
        ).pack(side='left', padx=5)

        save_btn = tk.Button(
            btn_frame,
            text="保存",
            font=FONTS["body"],
            bg=COLORS["primary"],
            fg="white",
            relief="flat",
            padx=20,
            command=self._save
        )
        save_btn.pack(side='left', padx=5)

    def _save(self):
        new_text = self.text_area.get("1.0", "end").strip()
        self.callback(new_text)
        self.dialog.destroy()
