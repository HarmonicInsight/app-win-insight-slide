# -*- coding: utf-8 -*-
"""
SettingsDialog - AI設定ダイアログ
"""
import tkinter as tk
from tkinter import ttk, messagebox
from ...config import COLORS, FONTS


class SettingsDialog:
    """AI設定ダイアログ"""

    def __init__(self, parent, processor):
        self.processor = processor
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("AI設定")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 画面中央に配置
        width, height = 450, 320
        screen_w = self.dialog.winfo_screenwidth()
        screen_h = self.dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")

        self._create_widgets()

    def _create_widgets(self):
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)

        # タイトル
        ttk.Label(
            frame,
            text="AI設定",
            font=FONTS["heading"]
        ).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 15))

        # プロバイダー選択
        ttk.Label(frame, text="AIプロバイダー:", font=FONTS["body"]).grid(row=1, column=0, sticky='w', pady=5)
        self.provider_var = tk.StringVar(value="mock")
        provider_combo = ttk.Combobox(
            frame,
            textvariable=self.provider_var,
            values=["mock", "openai", "claude"],
            width=30,
            state="readonly"
        )
        provider_combo.grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))

        # APIキー
        ttk.Label(frame, text="APIキー:", font=FONTS["body"]).grid(row=2, column=0, sticky='w', pady=5)
        self.api_key_var = tk.StringVar()
        ttk.Entry(
            frame,
            textvariable=self.api_key_var,
            width=32,
            show="*"
        ).grid(row=2, column=1, sticky='w', pady=5, padx=(10, 0))

        # モデル
        ttk.Label(frame, text="モデル:", font=FONTS["body"]).grid(row=3, column=0, sticky='w', pady=5)
        self.model_var = tk.StringVar(value="gpt-4o")
        ttk.Entry(
            frame,
            textvariable=self.model_var,
            width=32
        ).grid(row=3, column=1, sticky='w', pady=5, padx=(10, 0))

        # ヘルプテキスト
        help_text = ttk.Label(
            frame,
            text="※ mockはテスト用（API不要）\n※ APIキーは暗号化保存されません",
            font=FONTS["small"],
            foreground=COLORS["text_muted"]
        )
        help_text.grid(row=4, column=0, columnspan=2, sticky='w', pady=(10, 0))

        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=25)

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
            command=self._save_settings
        )
        save_btn.pack(side='left', padx=5)

    def _save_settings(self):
        provider = self.provider_var.get()
        api_key = self.api_key_var.get()
        model = self.model_var.get()

        try:
            if provider == "mock":
                self.processor.set_provider("mock")
            else:
                if not api_key:
                    messagebox.showwarning("入力エラー", "APIキーを入力してください")
                    return
                self.processor.set_provider(provider, api_key, model=model)

            messagebox.showinfo("成功", "AI設定を保存しました")
            self.dialog.destroy()
        except Exception as e:
            messagebox.showerror("エラー", str(e))
