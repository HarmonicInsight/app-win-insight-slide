# -*- coding: utf-8 -*-
"""
EditableGrid - インライン編集対応グリッドコンポーネント
"""
import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable, Dict, List
from ...config import COLORS, FONTS


class UndoManager:
    """Undo/Redo管理"""

    def __init__(self, max_history: int = 50):
        self.undo_stack: List[Dict] = []
        self.redo_stack: List[Dict] = []
        self.max_history = max_history

    def push(self, action: Dict):
        """アクションを記録 - action = {"type": str, "item": str, "column": str, "old": str, "new": str}"""
        self.undo_stack.append(action)
        self.redo_stack.clear()
        if len(self.undo_stack) > self.max_history:
            self.undo_stack.pop(0)

    def undo(self) -> Optional[Dict]:
        """元に戻す"""
        if not self.undo_stack:
            return None
        action = self.undo_stack.pop()
        self.redo_stack.append(action)
        return action

    def redo(self) -> Optional[Dict]:
        """やり直す"""
        if not self.redo_stack:
            return None
        action = self.redo_stack.pop()
        self.undo_stack.append(action)
        return action

    def can_undo(self) -> bool:
        return len(self.undo_stack) > 0

    def can_redo(self) -> bool:
        return len(self.redo_stack) > 0

    def clear(self):
        """履歴をクリア"""
        self.undo_stack.clear()
        self.redo_stack.clear()


class EditableGrid(ttk.Frame):
    """ダブルクリックでセル編集可能なTreeview"""

    def __init__(
        self,
        parent,
        columns: List[tuple],  # [(id, header, width, editable), ...]
        on_change: Optional[Callable] = None,
        on_ai_process: Optional[Callable] = None,
        **kwargs
    ):
        super().__init__(parent, **kwargs)

        self.columns_config = columns
        self.on_change = on_change
        self.on_ai_process = on_ai_process
        self.undo_manager = UndoManager()
        self._edit_widget: Optional[tk.Entry] = None
        self._editing_item: Optional[str] = None
        self._editing_column: Optional[str] = None

        self._create_widgets()
        self._setup_bindings()

    def _create_widgets(self):
        """ウィジェットを作成"""
        # ツールバー
        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 5))

        # AI処理ボタン
        ai_btn = tk.Button(
            toolbar,
            text="🤖 AI処理",
            font=FONTS["body"],
            bg=COLORS["primary"],
            fg="white",
            relief="flat",
            padx=10,
            cursor="hand2",
            command=self._show_ai_menu
        )
        ai_btn.pack(side="left", padx=(0, 5))

        # 一括置換ボタン
        replace_btn = tk.Button(
            toolbar,
            text="🔄 一括置換",
            font=FONTS["body"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            padx=10,
            cursor="hand2",
            command=self._show_replace_dialog
        )
        replace_btn.pack(side="left", padx=(0, 5))

        # Undoボタン
        self.undo_btn = tk.Button(
            toolbar,
            text="↩ 元に戻す",
            font=FONTS["body"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            padx=10,
            cursor="hand2",
            command=self._do_undo
        )
        self.undo_btn.pack(side="left", padx=(0, 5))

        # Redoボタン
        self.redo_btn = tk.Button(
            toolbar,
            text="↪ やり直し",
            font=FONTS["body"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            padx=10,
            cursor="hand2",
            command=self._do_redo
        )
        self.redo_btn.pack(side="left")

        # Treeview
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True)

        # カラム設定
        col_ids = [c[0] for c in self.columns_config]
        self.tree = ttk.Treeview(tree_frame, columns=col_ids, show="headings")

        for col_id, header, width, _ in self.columns_config:
            self.tree.heading(col_id, text=header)
            self.tree.column(col_id, width=width, minwidth=50)

        # スクロールバー
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # グリッド配置
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # タグスタイル（変更ハイライト用）
        self.tree.tag_configure("modified", background=COLORS["highlight"])

    def _setup_bindings(self):
        """イベントバインディングを設定"""
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._show_context_menu)  # 右クリック

        # キーボードショートカット
        self.tree.bind("<Control-z>", lambda e: self._do_undo())
        self.tree.bind("<Control-y>", lambda e: self._do_redo())
        self.tree.bind("<Control-Z>", lambda e: self._do_undo())
        self.tree.bind("<Control-Y>", lambda e: self._do_redo())

    def _on_double_click(self, event):
        """ダブルクリックでセル編集開始"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)

        if not item:
            return

        # カラムインデックスを取得（#1, #2, ... 形式）
        col_idx = int(column.replace("#", "")) - 1
        if col_idx < 0 or col_idx >= len(self.columns_config):
            return

        # 編集可能カラムかチェック
        _, _, _, editable = self.columns_config[col_idx]
        if not editable:
            return

        self._start_edit(item, column)

    def _start_edit(self, item: str, column: str):
        """編集を開始"""
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return

        col_idx = int(column.replace("#", "")) - 1
        col_id = self.columns_config[col_idx][0]
        current_value = self.tree.set(item, col_id)

        self._editing_item = item
        self._editing_column = col_id

        # Entryウィジェットをオーバーレイ
        self._edit_widget = tk.Entry(
            self.tree,
            font=FONTS["body"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            relief="solid",
            borderwidth=1
        )
        self._edit_widget.insert(0, current_value)
        self._edit_widget.select_range(0, tk.END)
        self._edit_widget.place(
            x=bbox[0],
            y=bbox[1],
            width=bbox[2],
            height=bbox[3]
        )
        self._edit_widget.focus_set()

        # イベントバインド
        self._edit_widget.bind("<Return>", lambda e: self._finish_edit())
        self._edit_widget.bind("<Escape>", lambda e: self._cancel_edit())
        self._edit_widget.bind("<FocusOut>", lambda e: self._finish_edit())

    def _finish_edit(self):
        """編集を完了"""
        if not self._edit_widget or not self._editing_item:
            return

        new_value = self._edit_widget.get()
        old_value = self.tree.set(self._editing_item, self._editing_column)

        if new_value != old_value:
            # 値を更新
            self.tree.set(self._editing_item, self._editing_column, new_value)

            # 変更をハイライト
            self.tree.item(self._editing_item, tags=("modified",))

            # Undo履歴に追加
            self.undo_manager.push({
                "type": "edit",
                "item": self._editing_item,
                "column": self._editing_column,
                "old": old_value,
                "new": new_value
            })

            # コールバック
            if self.on_change:
                self.on_change(self._editing_item, self._editing_column, new_value)

        self._cancel_edit()

    def _cancel_edit(self):
        """編集をキャンセル"""
        if self._edit_widget:
            self._edit_widget.destroy()
            self._edit_widget = None
        self._editing_item = None
        self._editing_column = None

    def _show_context_menu(self, event):
        """右クリックメニューを表示"""
        item = self.tree.identify_row(event.y)
        if not item:
            return

        self.tree.selection_set(item)

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="✏️ 編集", command=lambda: self._start_edit_selected())
        menu.add_separator()
        menu.add_command(label="🇬🇧 英語に翻訳", command=lambda: self._ai_process_item(item, "翻訳（英語）"))
        menu.add_command(label="🇯🇵 日本語に翻訳", command=lambda: self._ai_process_item(item, "翻訳（日本語）"))
        menu.add_command(label="📝 校正", command=lambda: self._ai_process_item(item, "校正"))
        menu.add_command(label="🔧 カスタムプロンプト...", command=lambda: self._ai_custom(item))
        menu.add_separator()
        menu.add_command(label="↩ 元に戻す", command=self._do_undo)
        menu.add_command(label="📋 コピー", command=lambda: self._copy_cell(item))

        menu.tk_popup(event.x_root, event.y_root)

    def _start_edit_selected(self):
        """選択中のアイテムを編集"""
        selection = self.tree.selection()
        if selection:
            # 最初の編集可能カラムを探す
            for i, (col_id, _, _, editable) in enumerate(self.columns_config):
                if editable:
                    self._start_edit(selection[0], f"#{i+1}")
                    break

    def _ai_process_item(self, item: str, preset: str):
        """AIで処理"""
        if self.on_ai_process:
            self.on_ai_process(item, preset)

    def _ai_custom(self, item: str):
        """カスタムAI処理"""
        if self.on_ai_process:
            self.on_ai_process(item, None)  # Noneの場合はダイアログ表示

    def _copy_cell(self, item: str):
        """セルをコピー"""
        # 編集可能カラムの値をコピー
        for col_id, _, _, editable in self.columns_config:
            if editable:
                value = self.tree.set(item, col_id)
                self.clipboard_clear()
                self.clipboard_append(value)
                break

    def _do_undo(self):
        """元に戻す"""
        action = self.undo_manager.undo()
        if action:
            self.tree.set(action["item"], action["column"], action["old"])
            # ハイライトを解除（元の値に戻った場合）
            # 簡易実装：タグを維持

    def _do_redo(self):
        """やり直し"""
        action = self.undo_manager.redo()
        if action:
            self.tree.set(action["item"], action["column"], action["new"])
            self.tree.item(action["item"], tags=("modified",))

    def _show_ai_menu(self):
        """AI処理メニューを表示"""
        selection = self.tree.selection()
        if not selection:
            tk.messagebox.showinfo("選択なし", "処理する行を選択してください")
            return

        if self.on_ai_process:
            self.on_ai_process(selection[0], None)

    def _show_replace_dialog(self):
        """一括置換ダイアログ"""
        # TODO: 一括置換ダイアログを実装
        pass

    # === 公開API ===

    def load_data(self, data: List[Dict]):
        """データを読み込み"""
        self.clear()
        for row in data:
            values = [row.get(col[0], "") for col in self.columns_config]
            self.tree.insert("", "end", values=values)

    def clear(self):
        """全データをクリア"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.undo_manager.clear()

    def get_data(self) -> List[Dict]:
        """現在のデータを取得"""
        result = []
        for item in self.tree.get_children():
            row = {}
            for col_id, _, _, _ in self.columns_config:
                row[col_id] = self.tree.set(item, col_id)
            result.append(row)
        return result

    def update_cell(self, item: str, column: str, value: str, record_undo: bool = True):
        """セルを更新"""
        if record_undo:
            old_value = self.tree.set(item, column)
            if old_value != value:
                self.undo_manager.push({
                    "type": "edit",
                    "item": item,
                    "column": column,
                    "old": old_value,
                    "new": value
                })

        self.tree.set(item, column, value)
        self.tree.item(item, tags=("modified",))

        if self.on_change:
            self.on_change(item, column, value)
