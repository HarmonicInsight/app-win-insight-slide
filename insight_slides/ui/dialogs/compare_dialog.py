# -*- coding: utf-8 -*-
"""
CompareDialog - PPTX比較機能
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
from ...config import COLORS, FONTS


class CompareDialog:
    """PPTX比較ダイアログ"""

    def __init__(self, parent, callback):
        self.callback = callback
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("PPTX比較")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        width, height = 600, 300
        screen_w = self.dialog.winfo_screenwidth()
        screen_h = self.dialog.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        self.dialog.resizable(False, False)

        self._create_widgets()

    def _create_widgets(self):
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(
            frame,
            text="2つのPowerPointファイルのテキストを比較します",
            font=FONTS["heading"]
        ).pack(anchor='w', pady=(0, 15))

        # ファイル1
        file1_frame = ttk.Frame(frame)
        file1_frame.pack(fill='x', pady=5)
        ttk.Label(file1_frame, text="元ファイル（Before）:", font=FONTS["body"], width=18).pack(side='left')
        self.file1_var = tk.StringVar()
        ttk.Entry(file1_frame, textvariable=self.file1_var, width=40).pack(side='left', padx=5)
        ttk.Button(file1_frame, text="参照...", command=lambda: self._browse_file(self.file1_var)).pack(side='left')

        # ファイル2
        file2_frame = ttk.Frame(frame)
        file2_frame.pack(fill='x', pady=5)
        ttk.Label(file2_frame, text="新ファイル（After）:", font=FONTS["body"], width=18).pack(side='left')
        self.file2_var = tk.StringVar()
        ttk.Entry(file2_frame, textvariable=self.file2_var, width=40).pack(side='left', padx=5)
        ttk.Button(file2_frame, text="参照...", command=lambda: self._browse_file(self.file2_var)).pack(side='left')

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

        exec_btn = tk.Button(
            btn_frame,
            text="比較実行",
            font=FONTS["body"],
            bg=COLORS["primary"],
            fg="white",
            relief="flat",
            width=12,
            command=self._execute
        )
        exec_btn.pack(side='left', padx=10)

    def _browse_file(self, var):
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")])
        if file_path:
            var.set(file_path)

    def _execute(self):
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

        # 選択状態管理
        self.selections = {}
        for i, row in enumerate(diff_data):
            if row["status"] == "変更":
                self.selections[i] = None
            elif row["status"] == "追加":
                self.selections[i] = "after"
            elif row["status"] == "削除":
                self.selections[i] = "before"
            else:
                self.selections[i] = "same"

        self._create_widgets(stats)

    def _create_widgets(self, stats):
        # 上部: 統計
        top_frame = ttk.Frame(self.window, padding=(10, 10, 10, 5))
        top_frame.pack(fill='x')

        ttk.Label(
            top_frame,
            text=f"📊 一致 {stats['same']} | 変更 {stats['changed']} | 追加 {stats['added']} | 削除 {stats['removed']}",
            font=FONTS["heading"]
        ).pack(side='left')

        ttk.Button(top_frame, text="CSVエクスポート", command=self._export_csv).pack(side='right', padx=5)
        ttk.Button(top_frame, text="変更のみ表示", command=self._toggle_filter).pack(side='right', padx=5)

        ttk.Label(
            self.window,
            text="  💡 クリックで採用を選択（未選択行は反映されません）",
            font=FONTS["small"],
            foreground=COLORS["text_muted"]
        ).pack(anchor='w', padx=10)

        # グリッド
        grid_frame = ttk.Frame(self.window, padding=(10, 5, 10, 5))
        grid_frame.pack(fill='both', expand=True)

        columns = ("select", "slide", "shape", "status", "before", "after")
        self.tree = ttk.Treeview(grid_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("select", text="採用")
        self.tree.heading("slide", text="スライド")
        self.tree.heading("shape", text="シェイプ")
        self.tree.heading("status", text="状態")
        self.tree.heading("before", text=f"元: {self.file1_name}")
        self.tree.heading("after", text=f"新: {self.file2_name}")

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
        self.tree.tag_configure("same", background=COLORS["surface"])
        self.tree.tag_configure("changed", background=COLORS["diff_changed"])
        self.tree.tag_configure("added", background=COLORS["diff_added"])
        self.tree.tag_configure("removed", background=COLORS["diff_removed"])
        self.tree.tag_configure("selected_before", background="#e3f2fd")
        self.tree.tag_configure("selected_after", background="#e8f5e9")

        self.show_all = True
        self.item_to_index = {}

        self.tree.bind("<Button-1>", self._on_click)

        # 下部ボタン
        bottom_frame = ttk.Frame(self.window, padding=10)
        bottom_frame.pack(fill='x', side='bottom')

        ttk.Button(bottom_frame, text="全て元", command=lambda: self._select_all("before"), width=10).pack(side='left', padx=2)
        ttk.Button(bottom_frame, text="全て新", command=lambda: self._select_all("after"), width=10).pack(side='left', padx=2)
        ttk.Button(bottom_frame, text="クリア", command=self._clear_selections, width=8).pack(side='left', padx=2)

        apply_btn = tk.Button(
            bottom_frame,
            text="選択を反映 →",
            font=FONTS["body"],
            bg=COLORS["primary"],
            fg="white",
            relief="flat",
            width=14,
            command=self._apply_selections
        )
        apply_btn.pack(side='right', padx=5)
        ttk.Button(bottom_frame, text="閉じる", command=self.window.destroy, width=10).pack(side='right', padx=5)

        self.selection_label = ttk.Label(bottom_frame, text="", font=FONTS["small"])
        self.selection_label.pack(side='right', padx=20)

        self._refresh_grid()

    def _refresh_grid(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.item_to_index = {}

        for i, row in enumerate(self.diff_data):
            if not self.show_all and row["status"] == "一致":
                continue

            before_text = row["before"].replace("\n", " ↵ ")[:60] if row["before"] else ""
            after_text = row["after"].replace("\n", " ↵ ")[:60] if row["after"] else ""

            selection = self.selections.get(i)
            if selection == "before":
                select_text = "◀ 元"
            elif selection == "after":
                select_text = "新 ▶"
            elif selection == "same":
                select_text = "─"
            else:
                select_text = "　"

            base_tag = {"一致": "same", "変更": "changed", "追加": "added", "削除": "removed"}.get(row["status"], "same")
            if selection == "before" and row["status"] != "一致":
                tag = "selected_before"
            elif selection == "after" and row["status"] != "一致":
                tag = "selected_after"
            else:
                tag = base_tag

            item_id = self.tree.insert("", "end", values=(
                select_text, row["slide"], row["shape"], row["status"], before_text, after_text
            ), tags=(tag,))

            self.item_to_index[item_id] = i

        self._update_selection_count()

    def _update_selection_count(self):
        selected = sum(1 for i, row in enumerate(self.diff_data)
                       if row["status"] != "一致" and self.selections.get(i) in ("before", "after"))
        total = sum(1 for row in self.diff_data if row["status"] != "一致")
        self.selection_label.configure(text=f"選択: {selected}/{total} 件")

    def _on_click(self, event):
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
        if row["status"] == "一致":
            return

        current = self.selections.get(idx)
        if column == "#5":
            self.selections[idx] = "before"
        elif column == "#6":
            self.selections[idx] = "after"
        else:
            self.selections[idx] = "after" if current == "before" else "before"

        self._refresh_grid()

    def _select_all(self, choice):
        for i, row in enumerate(self.diff_data):
            if row["status"] != "一致":
                self.selections[i] = choice
        self._refresh_grid()

    def _clear_selections(self):
        for i, row in enumerate(self.diff_data):
            if row["status"] == "変更":
                self.selections[i] = None
            elif row["status"] == "追加":
                self.selections[i] = "after"
            elif row["status"] == "削除":
                self.selections[i] = "before"
        self._refresh_grid()

    def _toggle_filter(self):
        self.show_all = not self.show_all
        self._refresh_grid()

    def _apply_selections(self):
        selected_data = []
        for i, row in enumerate(self.diff_data):
            selection = self.selections.get(i)
            if selection == "same" or selection is None:
                continue

            text = row["before"] if selection == "before" else row["after"]
            selected_data.append({
                "slide": row["slide"],
                "shape": row["shape"],
                "original": row["before"],
                "text": text,
                "status": row["status"],
                "selection": selection
            })

        if not selected_data:
            messagebox.showwarning("警告", "反映する項目が選択されていません")
            return

        if not messagebox.askyesno("確認", f"{len(selected_data)} 件の選択を反映しますか？"):
            return

        if self.on_apply_callback:
            self.on_apply_callback(self.file1_path, selected_data)
            messagebox.showinfo("完了", f"{len(selected_data)} 件をメイン画面に反映しました")
            self.window.destroy()

    def _export_csv(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not file_path:
            return

        with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["スライド", "シェイプ", "状態", "採用", "元ファイル", "新ファイル"])
            for i, row in enumerate(self.diff_data):
                selection = self.selections.get(i, "")
                sel_text = {"before": "元ファイル", "after": "新ファイル", "same": "一致"}.get(selection, "未選択")
                writer.writerow([row["slide"], row["shape"], row["status"], sel_text, row["before"], row["after"]])

        messagebox.showinfo("完了", f"CSVを保存しました:\n{file_path}")
