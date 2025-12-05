# -*- coding: utf-8 -*-
"""
StepIndicator - 3ステップ進捗表示コンポーネント
"""
import tkinter as tk
from tkinter import ttk
from ...config import COLORS, FONTS, STEPS


class StepIndicator(ttk.Frame):
    """3ステップの進捗インジケータ"""

    def __init__(self, parent, on_step_click=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.on_step_click = on_step_click
        self.current_step = 0
        self.step_labels = []
        self.step_circles = []
        self.connectors = []

        self._create_widgets()

    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインコンテナ
        container = ttk.Frame(self)
        container.pack(expand=True)

        for i, step in enumerate(STEPS):
            # 接続線（最初のステップ以外）
            if i > 0:
                connector = tk.Canvas(
                    container,
                    width=80,
                    height=4,
                    bg=COLORS["bg"],
                    highlightthickness=0
                )
                connector.pack(side="left", pady=20)
                connector.create_line(0, 2, 80, 2, fill=COLORS["step_inactive"], width=3)
                self.connectors.append(connector)

            # ステップフレーム
            step_frame = ttk.Frame(container)
            step_frame.pack(side="left", padx=10)

            # 円形インジケータ（Canvas）
            circle = tk.Canvas(
                step_frame,
                width=36,
                height=36,
                bg=COLORS["bg"],
                highlightthickness=0
            )
            circle.pack()

            # 円を描画
            circle.create_oval(
                2, 2, 34, 34,
                fill=COLORS["step_inactive"] if i > 0 else COLORS["step_active"],
                outline="",
                tags="circle"
            )
            # 番号テキスト
            circle.create_text(
                18, 18,
                text=step["icon"],
                fill="white",
                font=(FONTS["body"][0], 12, "bold"),
                tags="text"
            )

            # クリックイベント
            if self.on_step_click:
                circle.bind("<Button-1>", lambda e, idx=i: self._on_click(idx))
                circle.config(cursor="hand2")

            self.step_circles.append(circle)

            # ラベル
            label = ttk.Label(
                step_frame,
                text=step["label"],
                font=FONTS["small"],
                foreground=COLORS["text"] if i == 0 else COLORS["text_muted"]
            )
            label.pack(pady=(5, 0))
            self.step_labels.append(label)

    def _on_click(self, step_index: int):
        """ステップクリック時の処理"""
        if self.on_step_click:
            self.on_step_click(step_index)

    def set_step(self, step_index: int):
        """現在のステップを設定"""
        self.current_step = step_index

        for i, (circle, label) in enumerate(zip(self.step_circles, self.step_labels)):
            if i < step_index:
                # 完了したステップ
                circle.itemconfig("circle", fill=COLORS["step_complete"])
                label.config(foreground=COLORS["step_complete"])
            elif i == step_index:
                # 現在のステップ
                circle.itemconfig("circle", fill=COLORS["step_active"])
                label.config(foreground=COLORS["text"])
            else:
                # 未完了のステップ
                circle.itemconfig("circle", fill=COLORS["step_inactive"])
                label.config(foreground=COLORS["text_muted"])

        # 接続線の色を更新
        for i, connector in enumerate(self.connectors):
            connector.delete("all")
            color = COLORS["step_complete"] if i < step_index else COLORS["step_inactive"]
            connector.create_line(0, 2, 80, 2, fill=color, width=3)


class StepManager:
    """3ステップの状態管理"""

    def __init__(self, main_window):
        self.main_window = main_window
        self.current_step = 0
        self.data = {}  # ステップ間で共有するデータ

    def go_to(self, step: int | str):
        """指定ステップに移動"""
        if isinstance(step, str):
            step = next((i for i, s in enumerate(STEPS) if s["id"] == step), 0)

        self.current_step = step
        self.main_window.show_step(step)
        self.main_window.update_indicator(step)

    def next(self):
        """次のステップへ"""
        if self.current_step < len(STEPS) - 1:
            self.go_to(self.current_step + 1)

    def prev(self):
        """前のステップへ"""
        if self.current_step > 0:
            self.go_to(self.current_step - 1)

    def get_current_step_id(self) -> str:
        """現在のステップIDを取得"""
        return STEPS[self.current_step]["id"]
