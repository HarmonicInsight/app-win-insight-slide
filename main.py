"""
Insight Slides Python Edition - メインGUI
標準tkinterを使用（軽量版）
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
from pathlib import Path
from typing import Optional, List, Dict
import threading

from pptx_handler import extract_to_json, apply_from_json, save_json, load_json
from ai_processor import AIProcessor


# ============== 定数 ==============
APP_VERSION = "0.2.0"
APP_NAME = "Insight Slides"

# フォント設定
FONT_FAMILY = "Yu Gothic UI"
FONTS = {
    "title": (FONT_FAMILY, 16, "bold"),
    "heading": (FONT_FAMILY, 12, "bold"),
    "body": (FONT_FAMILY, 11),
    "small": (FONT_FAMILY, 10),
}

# 色設定
COLORS = {
    "primary": "#3367d6",
    "success": "#4caf50",
    "warning": "#e8871c",
    "bg": "#f5f5f5",
    "white": "#ffffff",
    "text": "#333333",
    "text_muted": "#666666",
    "diff_added": "#e6ffe6",      # 追加（薄緑）
    "diff_removed": "#ffe6e6",    # 削除（薄赤）
    "diff_changed": "#fff3e0",    # 変更（薄オレンジ）
}


def center_window(window, width, height, parent=None):
    """ウィンドウを中央に配置"""
    if parent:
        x = parent.winfo_x() + (parent.winfo_width() - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - height) // 2
    else:
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 2
    window.geometry(f"{width}x{height}+{max(0,x)}+{max(0,y)}")


class SettingsDialog:
    """AI設定ダイアログ"""
    
    def __init__(self, parent, processor: AIProcessor):
        self.processor = processor
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("AI設定")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 画面中央に配置
        width, height = 450, 280
        screen_w = self.dialog.winfo_screenwidth()
        screen_h = self.dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)
        
        # プロバイダー選択
        ttk.Label(frame, text="AIプロバイダー:", font=FONTS["body"]).grid(row=0, column=0, sticky='w', pady=5)
        self.provider_var = tk.StringVar(value="mock")
        provider_combo = ttk.Combobox(frame, textvariable=self.provider_var, values=["mock", "openai", "claude"], width=30, state="readonly")
        provider_combo.grid(row=0, column=1, sticky='w', pady=5, padx=(10, 0))
        
        # APIキー
        ttk.Label(frame, text="APIキー:", font=FONTS["body"]).grid(row=1, column=0, sticky='w', pady=5)
        self.api_key_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.api_key_var, width=32, show="*").grid(row=1, column=1, sticky='w', pady=5, padx=(10, 0))
        
        # モデル
        ttk.Label(frame, text="モデル:", font=FONTS["body"]).grid(row=2, column=0, sticky='w', pady=5)
        self.model_var = tk.StringVar(value="gpt-4o")
        ttk.Entry(frame, textvariable=self.model_var, width=32).grid(row=2, column=1, sticky='w', pady=5, padx=(10, 0))
        
        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="キャンセル", command=self.dialog.destroy).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="保存", command=self.save_settings).pack(side='left', padx=5)
    
    def save_settings(self):
        provider = self.provider_var.get()
        api_key = self.api_key_var.get()
        model = self.model_var.get()
        
        try:
            if provider == "mock":
                self.processor.set_provider("mock")
            else:
                self.processor.set_provider(provider, api_key, model=model)
            messagebox.showinfo("成功", "AI設定を保存しました")
            self.dialog.destroy()
        except Exception as e:
            messagebox.showerror("エラー", str(e))


class PromptDialog:
    """プロンプト入力ダイアログ"""
    
    def __init__(self, parent, processor: AIProcessor, callback):
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
        
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)
        
        # プリセット選択
        ttk.Label(frame, text="プリセット:", font=FONTS["body"]).pack(anchor='w')
        presets = list(processor.get_presets().keys())
        self.preset_var = tk.StringVar(value=presets[0] if presets else "")
        preset_combo = ttk.Combobox(frame, textvariable=self.preset_var, values=presets, width=50, state="readonly")
        preset_combo.pack(fill='x', pady=(5, 15))
        preset_combo.bind("<<ComboboxSelected>>", self.on_preset_change)
        
        # プロンプト入力
        ttk.Label(frame, text="プロンプト:", font=FONTS["body"]).pack(anchor='w')
        self.prompt_text = scrolledtext.ScrolledText(frame, width=60, height=10, font=FONTS["small"])
        self.prompt_text.pack(fill='both', expand=True, pady=5)
        
        # 初期プロンプト設定
        if presets:
            self.prompt_text.insert("1.0", processor.get_presets()[presets[0]])
        
        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="キャンセル", command=self.dialog.destroy).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="実行", command=self.execute).pack(side='left', padx=5)
    
    def on_preset_change(self, event):
        preset_prompt = self.processor.get_presets().get(self.preset_var.get(), "")
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", preset_prompt)
    
    def execute(self):
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
        
        # 画面中央に配置
        screen_w = self.dialog.winfo_screenwidth()
        screen_h = self.dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        self.dialog.resizable(True, True)
        self.dialog.minsize(400, 350)
        
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)
        
        # タイトル
        ttk.Label(frame, text=title, font=FONTS["heading"]).pack(anchor='w', pady=(0, 5))
        
        # 編集テキストラベル
        ttk.Label(frame, text="編集テキスト:", font=FONTS["body"], foreground=COLORS["primary"]).pack(anchor='w', pady=(5, 5))
        
        # テキストボックス（背景色を薄い青に）
        text_frame = ttk.Frame(frame)
        text_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        self.text_box = scrolledtext.ScrolledText(
            text_frame, 
            width=60, 
            height=12, 
            font=FONTS["body"],
            bg="#f0f8ff",
            relief="solid",
            borderwidth=1
        )
        self.text_box.pack(fill='both', expand=True)
        self.text_box.insert("1.0", current_text if current_text else "")
        
        # ボタンフレーム（固定高さ）
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        ttk.Button(btn_frame, text="キャンセル", command=self.dialog.destroy, width=12).pack(side='left', padx=(0, 10))
        ttk.Button(btn_frame, text="保存", command=self.save, width=12).pack(side='left')
    
    def save(self):
        new_text = self.text_box.get("1.0", "end").strip()
        self.callback(new_text)
        self.dialog.destroy()


class CompareDialog:
    """PPTX比較ダイアログ"""
    
    def __init__(self, parent, callback):
        self.callback = callback
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("PPTX比較")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 画面中央に配置
        width, height = 600, 300
        screen_w = self.dialog.winfo_screenwidth()
        screen_h = self.dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        self.dialog.resizable(False, False)
        
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)
        
        # 説明
        ttk.Label(
            frame, 
            text="2つのPowerPointファイルのテキストを比較します",
            font=FONTS["heading"]
        ).pack(anchor='w', pady=(0, 15))
        
        # ファイル1（元ファイル）
        file1_frame = ttk.Frame(frame)
        file1_frame.pack(fill='x', pady=5)
        ttk.Label(file1_frame, text="元ファイル（Before）:", font=FONTS["body"], width=18).pack(side='left')
        self.file1_var = tk.StringVar()
        ttk.Entry(file1_frame, textvariable=self.file1_var, width=40).pack(side='left', padx=5)
        ttk.Button(file1_frame, text="参照...", command=lambda: self.browse_file(self.file1_var)).pack(side='left')
        
        # ファイル2（新ファイル）
        file2_frame = ttk.Frame(frame)
        file2_frame.pack(fill='x', pady=5)
        ttk.Label(file2_frame, text="新ファイル（After）:", font=FONTS["body"], width=18).pack(side='left')
        self.file2_var = tk.StringVar()
        ttk.Entry(file2_frame, textvariable=self.file2_var, width=40).pack(side='left', padx=5)
        ttk.Button(file2_frame, text="参照...", command=lambda: self.browse_file(self.file2_var)).pack(side='left')
        
        # オプション
        opt_frame = ttk.Frame(frame)
        opt_frame.pack(fill='x', pady=15)
        self.ignore_whitespace = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="空白の違いを無視", variable=self.ignore_whitespace).pack(side='left')
        self.show_only_diff = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="差分のみ表示", variable=self.show_only_diff).pack(side='left', padx=20)
        
        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=(15, 0))
        ttk.Button(btn_frame, text="キャンセル", command=self.dialog.destroy, width=12).pack(side='left')
        ttk.Button(btn_frame, text="比較実行", command=self.execute, width=12).pack(side='left', padx=10)
    
    def browse_file(self, var):
        file_path = filedialog.askopenfilename(
            filetypes=[("PowerPoint", "*.pptx")]
        )
        if file_path:
            var.set(file_path)
    
    def execute(self):
        file1 = self.file1_var.get()
        file2 = self.file2_var.get()
        
        if not file1 or not file2:
            messagebox.showwarning("警告", "2つのファイルを選択してください")
            return
        
        self.callback(file1, file2, self.ignore_whitespace.get(), self.show_only_diff.get())
        self.dialog.destroy()


class CompareResultWindow:
    """比較結果ウィンドウ"""
    
    def __init__(self, parent, file1_path, file2_path, file1_name, file2_name, diff_data, stats, on_apply_callback=None):
        self.window = tk.Toplevel(parent)
        self.window.title(f"比較結果: {file1_name} ↔ {file2_name}")
        
        # サイズ設定（画面中央に配置）
        width, height = 1200, 800
        screen_w = self.window.winfo_screenwidth()
        screen_h = self.window.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2 - 50
        self.window.geometry(f"{width}x{height}+{x}+{max(0, y)}")
        self.window.minsize(900, 600)
        
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.file1_name = file1_name
        self.file2_name = file2_name
        self.diff_data = diff_data
        self.on_apply_callback = on_apply_callback
        
        # 各行の選択状態を管理（"before", "after", None）
        self.selections = {}
        for i, row in enumerate(diff_data):
            if row["status"] == "変更":
                self.selections[i] = None  # 未選択
            elif row["status"] == "追加":
                self.selections[i] = "after"  # デフォルトで新ファイル
            elif row["status"] == "削除":
                self.selections[i] = "before"  # デフォルトで元ファイル
            else:
                self.selections[i] = "same"  # 一致（選択不要）
        
        # ===== 上部: 統計情報 =====
        top_frame = ttk.Frame(self.window, padding=(10, 10, 10, 5))
        top_frame.pack(fill='x')
        
        ttk.Label(
            top_frame,
            text=f"📊 一致 {stats['same']} | 変更 {stats['changed']} | 追加 {stats['added']} | 削除 {stats['removed']}",
            font=FONTS["heading"]
        ).pack(side='left')
        
        ttk.Button(top_frame, text="CSVエクスポート", command=lambda: self.export_csv(diff_data)).pack(side='right', padx=5)
        ttk.Button(top_frame, text="変更のみ表示", command=self.toggle_filter).pack(side='right', padx=5)
        
        # 説明
        ttk.Label(
            self.window,
            text="  💡 クリックで採用を選択（未選択行は反映されません）",
            font=FONTS["small"],
            foreground=COLORS["text_muted"]
        ).pack(anchor='w', padx=10)
        
        # ===== 中央: グリッド =====
        grid_frame = ttk.Frame(self.window, padding=(10, 5, 10, 5))
        grid_frame.pack(fill='both', expand=True)
        
        columns = ("select", "slide", "shape", "status", "before", "after")
        self.tree = ttk.Treeview(grid_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("select", text="採用")
        self.tree.heading("slide", text="スライド")
        self.tree.heading("shape", text="シェイプ")
        self.tree.heading("status", text="状態")
        self.tree.heading("before", text=f"元: {file1_name}")
        self.tree.heading("after", text=f"新: {file2_name}")
        
        self.tree.column("select", width=80, anchor="center")
        self.tree.column("slide", width=60, anchor="center")
        self.tree.column("shape", width=100)
        self.tree.column("status", width=60, anchor="center")
        self.tree.column("before", width=380)
        self.tree.column("after", width=380)
        
        scrollbar_y = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        self.tree.pack(fill='both', expand=True)
        
        # タグ設定
        self.tree.tag_configure("same", background=COLORS["white"])
        self.tree.tag_configure("changed", background=COLORS["diff_changed"])
        self.tree.tag_configure("added", background=COLORS["diff_added"])
        self.tree.tag_configure("removed", background=COLORS["diff_removed"])
        self.tree.tag_configure("selected_before", background="#e3f2fd")
        self.tree.tag_configure("selected_after", background="#e8f5e9")
        
        self.show_all = True
        self.item_to_index = {}
        
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.show_detail)
        
        # ===== 下部: ボタン（固定高さ）=====
        bottom_frame = ttk.Frame(self.window, padding=10)
        bottom_frame.pack(fill='x', side='bottom')
        
        # 左側：一括選択
        ttk.Button(bottom_frame, text="全て元", command=lambda: self.select_all("before"), width=10).pack(side='left', padx=2)
        ttk.Button(bottom_frame, text="全て新", command=lambda: self.select_all("after"), width=10).pack(side='left', padx=2)
        ttk.Button(bottom_frame, text="クリア", command=self.clear_selections, width=8).pack(side='left', padx=2)
        
        # 右側：アクション
        ttk.Button(bottom_frame, text="選択を反映 →", command=self.apply_selections, width=14).pack(side='right', padx=5)
        ttk.Button(bottom_frame, text="閉じる", command=self.window.destroy, width=10).pack(side='right', padx=5)
        
        # 選択数表示（refresh_gridより先に作成）
        self.selection_label = ttk.Label(bottom_frame, text="", font=FONTS["small"])
        self.selection_label.pack(side='right', padx=20)
        
        # データ挿入（selection_label作成後に実行）
        self.refresh_grid()
    
    def refresh_grid(self):
        """グリッドを更新"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.item_to_index = {}
        
        for i, row in enumerate(self.diff_data):
            if not self.show_all and row["status"] == "一致":
                continue
            
            before_text = row["before"].replace("\n", " ↵ ")[:60] if row["before"] else ""
            after_text = row["after"].replace("\n", " ↵ ")[:60] if row["after"] else ""
            
            # 選択状態の表示
            selection = self.selections.get(i)
            if selection == "before":
                select_text = "◀ 元"
            elif selection == "after":
                select_text = "新 ▶"
            elif selection == "same":
                select_text = "─"
            else:
                select_text = "　"  # 未選択は空白
            
            # タグ決定
            base_tag = {"一致": "same", "変更": "changed", "追加": "added", "削除": "removed"}.get(row["status"], "same")
            if selection == "before" and row["status"] != "一致":
                tag = "selected_before"
            elif selection == "after" and row["status"] != "一致":
                tag = "selected_after"
            else:
                tag = base_tag
            
            item_id = self.tree.insert("", "end", values=(
                select_text,
                row["slide"],
                row["shape"],
                row["status"],
                before_text,
                after_text
            ), tags=(tag,))
            
            self.item_to_index[item_id] = i
        
        self.update_selection_count()
    
    def update_selection_count(self):
        """選択数を更新"""
        selected = sum(1 for i, row in enumerate(self.diff_data) 
                      if row["status"] != "一致" and self.selections.get(i) in ("before", "after"))
        total = sum(1 for row in self.diff_data if row["status"] != "一致")
        self.selection_label.configure(text=f"選択: {selected}/{total} 件")
    
    def on_click(self, event):
        """クリックで選択切り替え"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        if not item:
            return
        
        idx = self.item_to_index.get(item)
        if idx is None:
            return
        
        row = self.diff_data[idx]
        
        # 一致行は選択不可
        if row["status"] == "一致":
            return
        
        # 選択列または他の列クリックで選択切り替え
        current = self.selections.get(idx)
        
        if column == "#1":  # 採用列
            # before → after → before のトグル
            if current == "before":
                self.selections[idx] = "after"
            else:
                self.selections[idx] = "before"
        elif column == "#5":  # before列
            self.selections[idx] = "before"
        elif column == "#6":  # after列
            self.selections[idx] = "after"
        else:
            # その他の列はトグル
            if current == "before":
                self.selections[idx] = "after"
            else:
                self.selections[idx] = "before"
        
        self.refresh_grid()
    
    def select_all(self, choice):
        """一括選択"""
        for i, row in enumerate(self.diff_data):
            if row["status"] != "一致":
                self.selections[i] = choice
        self.refresh_grid()
    
    def clear_selections(self):
        """選択クリア"""
        for i, row in enumerate(self.diff_data):
            if row["status"] == "変更":
                self.selections[i] = None
            elif row["status"] == "追加":
                self.selections[i] = "after"
            elif row["status"] == "削除":
                self.selections[i] = "before"
        self.refresh_grid()
    
    def toggle_filter(self):
        """フィルター切り替え"""
        self.show_all = not self.show_all
        self.refresh_grid()
    
    def show_detail(self, event):
        """詳細表示"""
        item = self.tree.identify_row(event.y)
        if not item:
            return
        
        idx = self.item_to_index.get(item)
        if idx is None:
            return
        
        row = self.diff_data[idx]
        self.show_detail_dialog(idx, row)
    
    def show_detail_dialog(self, idx, row):
        """詳細ダイアログ"""
        dialog = tk.Toplevel(self.window)
        dialog.title(f"スライド {row['slide']} - {row['shape']} ({row['status']})")
        dialog.transient(self.window)
        dialog.grab_set()
        
        # サイズ設定（画面中央に配置）
        width, height = 900, 600
        screen_w = dialog.winfo_screenwidth()
        screen_h = dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.minsize(700, 450)
        
        # メインフレーム
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill='both', expand=True)
        
        # 選択ボタン（上部）
        if row["status"] != "一致":
            select_frame = ttk.Frame(main_frame)
            select_frame.pack(fill='x', pady=(0, 10))
            
            ttk.Label(select_frame, text="採用するテキスト:", font=FONTS["body"]).pack(side='left')
            
            self.detail_selection = tk.StringVar(value=self.selections.get(idx, "after") or "after")
            ttk.Radiobutton(select_frame, text="元ファイル", variable=self.detail_selection, value="before").pack(side='left', padx=10)
            ttk.Radiobutton(select_frame, text="新ファイル", variable=self.detail_selection, value="after").pack(side='left', padx=10)
        
        # テキスト比較エリア（中央、拡張可能）
        paned = ttk.PanedWindow(main_frame, orient='horizontal')
        paned.pack(fill='both', expand=True, pady=5)
        
        # Before
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        ttk.Label(left_frame, text=f"元ファイル: {self.file1_name}", font=FONTS["heading"], foreground=COLORS["text_muted"]).pack(anchor='w')
        before_text = scrolledtext.ScrolledText(left_frame, width=50, height=16, font=FONTS["body"], bg="#fff5f5")
        before_text.pack(fill='both', expand=True, pady=5)
        before_text.insert("1.0", row["before"] if row["before"] else "(なし)")
        before_text.configure(state='disabled')
        
        # After
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        ttk.Label(right_frame, text=f"新ファイル: {self.file2_name}", font=FONTS["heading"], foreground=COLORS["primary"]).pack(anchor='w')
        after_text = scrolledtext.ScrolledText(right_frame, width=50, height=16, font=FONTS["body"], bg="#f5fff5")
        after_text.pack(fill='both', expand=True, pady=5)
        after_text.insert("1.0", row["after"] if row["after"] else "(なし)")
        after_text.configure(state='disabled')
        
        # ボタン（下部、固定）
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        def save_and_close():
            if row["status"] != "一致":
                self.selections[idx] = self.detail_selection.get()
                self.refresh_grid()
            dialog.destroy()
        
        ttk.Button(btn_frame, text="閉じる", command=dialog.destroy, width=10).pack(side='left', padx=5)
        if row["status"] != "一致":
            ttk.Button(btn_frame, text="選択を保存", command=save_and_close, width=12).pack(side='left', padx=5)
    
    def apply_selections(self):
        """選択を反映（未選択行は反映しない）"""
        # 選択された行のみ抽出
        selected_data = []
        for i, row in enumerate(self.diff_data):
            selection = self.selections.get(i)
            
            # 一致または未選択はスキップ
            if selection == "same" or selection is None:
                continue
            
            if selection == "before":
                text = row["before"]
            else:
                text = row["after"]
            
            selected_data.append({
                "slide": row["slide"],
                "shape": row["shape"],
                "original": row["before"],  # 元ファイルの内容をoriginalに
                "text": text,               # 選択した内容をtextに
                "status": row["status"],
                "selection": selection
            })
        
        if not selected_data:
            messagebox.showwarning("警告", "反映する項目が選択されていません")
            return
        
        # 確認
        msg = f"{len(selected_data)} 件の選択を反映しますか？"
        if not messagebox.askyesno("確認", msg):
            return
        
        # コールバックで反映
        if self.on_apply_callback:
            self.on_apply_callback(self.file1_path, selected_data)
            messagebox.showinfo("完了", f"{len(selected_data)} 件をメイン画面に反映しました")
            self.window.destroy()
        else:
            messagebox.showwarning("警告", "反映先が設定されていません")
    
    def export_csv(self, diff_data):
        """CSVエクスポート"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not file_path:
            return
        
        import csv
        with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["スライド", "シェイプ", "状態", "採用", "元ファイル", "新ファイル"])
            for i, row in enumerate(diff_data):
                selection = self.selections.get(i, "")
                if selection == "before":
                    sel_text = "元ファイル"
                elif selection == "after":
                    sel_text = "新ファイル"
                elif selection == "same":
                    sel_text = "一致"
                else:
                    sel_text = "未選択"
                writer.writerow([row["slide"], row["shape"], row["status"], sel_text, row["before"], row["after"]])
        
        messagebox.showinfo("完了", f"CSVを保存しました:\n{file_path}")


class MainWindow:
    """メインウィンドウ"""
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1100x650")
        
        # 状態変数
        self.current_file: Optional[str] = None
        self.json_data: Optional[dict] = None
        self.ai_processor = AIProcessor()
        self.ai_processor.set_provider("mock")
        self.row_mapping = []
        
        # DPI対応
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
        
        self.setup_styles()
        self.setup_ui()
    
    def setup_styles(self):
        """スタイル設定"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Treeviewのスタイル
        style.configure("Treeview", rowheight=60, font=FONTS["body"])
        style.configure("Treeview.Heading", font=FONTS["heading"])
    
    def setup_ui(self):
        """UIのセットアップ"""
        
        # ヘッダー
        header = tk.Frame(self.root, bg=COLORS["primary"], height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text=f"📊 {APP_NAME}",
            font=FONTS["title"],
            bg=COLORS["primary"],
            fg=COLORS["white"]
        ).pack(side='left', padx=15, pady=10)
        
        ttk.Button(header, text="⚙ AI設定", command=self.open_settings).pack(side='right', padx=10, pady=8)
        
        # メインコンテンツ
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill='both', expand=True)
        
        # ツールバー
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill='x', pady=(0, 10))
        
        ttk.Button(toolbar, text="📂 ファイルを開く", command=self.open_file).pack(side='left', padx=2)
        ttk.Button(toolbar, text="💾 JSONを保存", command=self.save_json_file).pack(side='left', padx=2)
        ttk.Button(toolbar, text="🤖 AI処理", command=self.open_ai_dialog).pack(side='left', padx=2)
        ttk.Button(toolbar, text="✅ PPTXに反映", command=self.apply_to_pptx).pack(side='left', padx=2)
        
        # セパレーター
        ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=8)
        
        ttk.Button(toolbar, text="🔍 PPTX比較", command=self.open_compare_dialog).pack(side='left', padx=2)
        
        self.file_label = ttk.Label(toolbar, text="ファイル未選択", foreground=COLORS["text_muted"])
        self.file_label.pack(side='right', padx=10)
        
        # グリッド（Treeview）
        grid_frame = ttk.Frame(main_frame)
        grid_frame.pack(fill='both', expand=True)
        
        columns = ("slide", "shape", "original", "text", "status")
        self.tree = ttk.Treeview(grid_frame, columns=columns, show="headings", selectmode="extended")
        
        self.tree.heading("slide", text="スライド")
        self.tree.heading("shape", text="シェイプ")
        self.tree.heading("original", text="元テキスト")
        self.tree.heading("text", text="編集テキスト")
        self.tree.heading("status", text="状態")
        
        self.tree.column("slide", width=70, anchor="center")
        self.tree.column("shape", width=100)
        self.tree.column("original", width=350)
        self.tree.column("text", width=350)
        self.tree.column("status", width=60, anchor="center")
        
        # スクロールバー
        scrollbar_y = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        self.tree.pack(fill='both', expand=True)
        
        # ダブルクリックで編集
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # ステータスバー
        self.status_bar = ttk.Label(self.root, text="準備完了", anchor='w', padding=(10, 5))
        self.status_bar.pack(fill='x', side='bottom')
    
    def open_file(self):
        """ファイルを開く"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("PowerPoint", "*.pptx"),
                ("JSON", "*.json"),
                ("すべて", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith(".json"):
                self.json_data = load_json(file_path)
                self.current_file = None
            else:
                self.json_data = extract_to_json(file_path)
                self.current_file = file_path
            
            self.file_label.configure(text=Path(file_path).name)
            self.refresh_grid()
            self.set_status(f"読み込み完了: {len(self.json_data.get('slides', []))} スライド")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル読み込みエラー:\n{str(e)}")
    
    def refresh_grid(self):
        """グリッドを更新"""
        # 既存行をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.row_mapping = []
        
        if not self.json_data:
            return
        
        # データを挿入
        for slide in self.json_data.get("slides", []):
            slide_num = slide["slide"]
            for shape in slide.get("shapes", []):
                original = shape.get("original", shape.get("text", ""))
                text = shape.get("text", "")
                status = "─" if original == text else "✎"
                
                # 表示用に改行を置換し、長いテキストは省略
                display_original = original.replace("\n", " ↵ ")
                display_text = text.replace("\n", " ↵ ")
                if len(display_original) > 80:
                    display_original = display_original[:80] + "..."
                if len(display_text) > 80:
                    display_text = display_text[:80] + "..."
                
                item_id = self.tree.insert("", "end", values=(
                    slide_num,
                    shape.get("name", ""),
                    display_original,
                    display_text,
                    status
                ))
                
                self.row_mapping.append({
                    "item_id": item_id,
                    "slide": slide_num,
                    "shape_id": shape.get("id"),
                    "shape_name": shape.get("name", ""),
                    "original": original,
                    "text": text
                })
    
    def on_double_click(self, event):
        """セルダブルクリックで編集"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        if not item or column != "#4":  # text列のみ編集可能
            return
        
        # マッピングから元のテキストを取得
        for mapping in self.row_mapping:
            if mapping["item_id"] == item:
                title = f"スライド {mapping['slide']} - {mapping['shape_name']}"
                EditDialog(self.root, title, mapping["text"], lambda new_text, m=mapping: self.update_cell(m, new_text))
                break
    
    def update_cell(self, mapping, new_text):
        """セルの値を更新"""
        slide_num = mapping["slide"]
        shape_id = mapping["shape_id"]
        item_id = mapping["item_id"]
        original = mapping["original"]
        
        # マッピングを更新
        mapping["text"] = new_text
        
        # Treeviewを更新
        display_text = new_text.replace("\n", " ↵ ")
        if len(display_text) > 80:
            display_text = display_text[:80] + "..."
        status = "─" if original == new_text else "✎"
        
        values = list(self.tree.item(item_id, "values"))
        values[3] = display_text
        values[4] = status
        self.tree.item(item_id, values=values)
        
        # JSONデータを更新
        if self.json_data:
            for slide in self.json_data["slides"]:
                if slide["slide"] == slide_num:
                    for shape in slide["shapes"]:
                        if shape["id"] == shape_id:
                            shape["text"] = new_text
                            break
    
    def save_json_file(self):
        """JSONを保存"""
        if not self.json_data:
            messagebox.showwarning("警告", "データがありません")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        
        if file_path:
            save_json(self.json_data, file_path)
            self.set_status(f"保存完了: {file_path}")
    
    def apply_to_pptx(self):
        """PPTXに反映"""
        if not self.json_data:
            messagebox.showwarning("警告", "データがありません")
            return
        
        if not self.current_file:
            messagebox.showwarning("警告", "元のPPTXファイルがありません")
            return
        
        try:
            output_path = apply_from_json(self.current_file, self.json_data)
            self.set_status(f"反映完了: {output_path}")
            messagebox.showinfo("完了", f"PPTXファイルを作成しました:\n{output_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"反映エラー:\n{str(e)}")
    
    def open_settings(self):
        """AI設定ダイアログを開く"""
        SettingsDialog(self.root, self.ai_processor)
    
    def open_ai_dialog(self):
        """AI処理ダイアログを開く"""
        if not self.json_data:
            messagebox.showwarning("警告", "データがありません")
            return
        
        PromptDialog(self.root, self.ai_processor, self.execute_ai)
    
    def execute_ai(self, prompt: str):
        """AI処理を実行（全件対象、originalを入力としてtextに出力）"""
        self.set_status("AI処理中...")
        
        # originalフィールドを入力用JSONとして準備
        input_data = {
            "file": self.json_data.get("file", ""),
            "slides": []
        }
        for slide in self.json_data.get("slides", []):
            slide_copy = {
                "slide": slide["slide"],
                "shapes": []
            }
            for shape in slide.get("shapes", []):
                slide_copy["shapes"].append({
                    "id": shape["id"],
                    "name": shape["name"],
                    "text": shape.get("original", shape["text"])
                })
            input_data["slides"].append(slide_copy)
        
        def process():
            try:
                result = self.ai_processor.process_json(prompt, input_data)
                
                # 結果をtextフィールドに反映（originalは維持）
                for slide in self.json_data.get("slides", []):
                    for shape in slide.get("shapes", []):
                        for res_slide in result.get("slides", []):
                            if res_slide["slide"] == slide["slide"]:
                                for res_shape in res_slide.get("shapes", []):
                                    if res_shape["id"] == shape["id"]:
                                        shape["text"] = res_shape["text"]
                                        break
                
                self.root.after(0, self.refresh_grid)
                self.root.after(0, lambda: self.set_status("AI処理完了"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("エラー", str(e)))
                self.root.after(0, lambda: self.set_status("AI処理エラー"))
        
        thread = threading.Thread(target=process)
        thread.start()
    
    def set_status(self, message: str):
        """ステータスバーを更新"""
        self.status_bar.configure(text=message)
    
    def open_compare_dialog(self):
        """PPTX比較ダイアログを開く"""
        CompareDialog(self.root, self.execute_compare)
    
    def execute_compare(self, file1: str, file2: str, ignore_whitespace: bool, show_only_diff: bool):
        """PPTX比較を実行"""
        self.set_status("比較中...")
        
        try:
            # 両ファイルからテキストを抽出
            json1 = extract_to_json(file1)
            json2 = extract_to_json(file2)
            
            # テキストをフラット化して比較用辞書を作成
            def flatten_texts(json_data):
                texts = {}
                for slide in json_data.get("slides", []):
                    slide_num = slide["slide"]
                    for shape in slide.get("shapes", []):
                        key = (slide_num, shape.get("name", ""), shape.get("id"))
                        text = shape.get("text", "")
                        if ignore_whitespace:
                            text = " ".join(text.split())
                        texts[key] = {
                            "slide": slide_num,
                            "shape": shape.get("name", ""),
                            "shape_id": shape.get("id"),
                            "text": shape.get("text", ""),  # 元のテキスト（表示用）
                            "text_normalized": text  # 比較用
                        }
                return texts
            
            texts1 = flatten_texts(json1)
            texts2 = flatten_texts(json2)
            
            # 比較結果を生成
            diff_data = []
            stats = {"same": 0, "changed": 0, "added": 0, "removed": 0}
            
            all_keys = set(texts1.keys()) | set(texts2.keys())
            
            for key in sorted(all_keys, key=lambda k: (k[0], k[1])):
                slide_num, shape_name, shape_id = key
                
                in1 = key in texts1
                in2 = key in texts2
                
                if in1 and in2:
                    t1 = texts1[key]
                    t2 = texts2[key]
                    if t1["text_normalized"] == t2["text_normalized"]:
                        status = "一致"
                        stats["same"] += 1
                    else:
                        status = "変更"
                        stats["changed"] += 1
                    diff_data.append({
                        "slide": slide_num,
                        "shape": shape_name,
                        "shape_id": shape_id,
                        "status": status,
                        "before": t1["text"],
                        "after": t2["text"]
                    })
                elif in1:
                    status = "削除"
                    stats["removed"] += 1
                    diff_data.append({
                        "slide": slide_num,
                        "shape": shape_name,
                        "shape_id": texts1[key]["shape_id"],
                        "status": status,
                        "before": texts1[key]["text"],
                        "after": ""
                    })
                else:
                    status = "追加"
                    stats["added"] += 1
                    diff_data.append({
                        "slide": slide_num,
                        "shape": shape_name,
                        "shape_id": texts2[key]["shape_id"],
                        "status": status,
                        "before": "",
                        "after": texts2[key]["text"]
                    })
            
            # フィルター
            display_data = diff_data
            if show_only_diff:
                display_data = [d for d in diff_data if d["status"] != "一致"]
            
            # 結果ウィンドウを表示（コールバック付き）
            file1_name = Path(file1).name
            file2_name = Path(file2).name
            CompareResultWindow(
                self.root, 
                file1, file2,
                file1_name, file2_name, 
                diff_data,  # 全データを渡す（フィルターは内部で）
                stats,
                on_apply_callback=self.apply_compare_result
            )
            
            self.set_status(f"比較完了: 一致 {stats['same']} / 変更 {stats['changed']} / 追加 {stats['added']} / 削除 {stats['removed']}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"比較エラー:\n{str(e)}")
            self.set_status("比較エラー")
    
    def apply_compare_result(self, base_file: str, result_data: list):
        """比較結果をメイン画面に反映（選択行のみ）"""
        try:
            # 元ファイルのJSONを読み込み
            self.json_data = extract_to_json(base_file)
            self.current_file = base_file
            
            # 結果データを反映（選択されたもののみ）
            updated_count = 0
            for result in result_data:
                slide_num = result["slide"]
                shape_name = result["shape"]
                
                # JSONデータ内で該当するシェイプを探して更新
                for slide in self.json_data.get("slides", []):
                    if slide["slide"] == slide_num:
                        for shape in slide.get("shapes", []):
                            if shape.get("name", "") == shape_name:
                                # originalは元ファイルの内容を保持
                                shape["original"] = result["original"] if result["original"] else shape.get("text", "")
                                # textは選択された内容
                                shape["text"] = result["text"] if result["text"] else ""
                                updated_count += 1
                                break
            
            # グリッド更新
            self.file_label.configure(text=Path(base_file).name + " (比較結果)")
            self.refresh_grid()
            self.set_status(f"比較結果を反映: {updated_count} 項目を更新")
            
        except Exception as e:
            messagebox.showerror("エラー", f"反映エラー:\n{str(e)}")


def main():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
