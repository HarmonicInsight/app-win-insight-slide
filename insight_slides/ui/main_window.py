# -*- coding: utf-8 -*-
"""
MainWindow - 3ステップ構成のメインウィンドウ
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from typing import Optional, Dict, List
import threading

from ..config import APP_NAME, APP_VERSION, COLORS, FONTS, WINDOW_SIZE
from ..core.pptx_handler import extract_to_json, apply_from_json, save_json, load_json, load_excel, save_excel
from ..core.ai_processor import AIProcessor

from .components.step_indicator import StepIndicator, StepManager
from .components.drop_zone import DropZone
from .components.editable_grid import EditableGrid

from .dialogs.settings_dialog import SettingsDialog
from .dialogs.ai_dialog import AIDialog
from .dialogs.compare_dialog import CompareDialog, CompareResultWindow


class MainWindow:
    """3ステップ構成のメインウィンドウ"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry(f"{WINDOW_SIZE['main'][0]}x{WINDOW_SIZE['main'][1]}")
        self.root.minsize(900, 600)

        # 状態変数
        self.current_file: Optional[str] = None
        self.json_data: Optional[Dict] = None
        self.ai_processor = AIProcessor()
        self.ai_processor.set_provider("mock")

        # DPI対応
        self._setup_dpi()

        # スタイル設定
        self._setup_styles()

        # UI構築
        self._create_ui()

        # ステップマネージャー
        self.step_manager = StepManager(self)
        self.step_manager.go_to(0)

    def _setup_dpi(self):
        """DPI対応"""
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass

    def _setup_styles(self):
        """スタイル設定"""
        style = ttk.Style()
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("TLabel", background=COLORS["bg"], font=FONTS["body"])
        style.configure("Header.TFrame", background=COLORS["surface"])
        style.configure("Card.TFrame", background=COLORS["surface"])

    def _create_ui(self):
        """UIを構築"""
        # ヘッダー
        self._create_header()

        # ステップインジケータ
        self._create_step_indicator()

        # メインコンテンツ（ステップ別）
        self._create_main_content()

        # ステータスバー
        self._create_status_bar()

    def _create_header(self):
        """ヘッダーを作成"""
        header = ttk.Frame(self.root, style="Header.TFrame")
        header.pack(fill="x", pady=(0, 1))

        # 内側パディング
        inner = ttk.Frame(header, style="Header.TFrame")
        inner.pack(fill="x", padx=15, pady=10)

        # 左側: タイトル
        title_frame = ttk.Frame(inner, style="Header.TFrame")
        title_frame.pack(side="left")

        title_label = tk.Label(
            title_frame,
            text=f"◇ {APP_NAME}",
            font=FONTS["title"],
            fg=COLORS["primary"],
            bg=COLORS["surface"]
        )
        title_label.pack(side="left")

        version_label = tk.Label(
            title_frame,
            text=f"  v{APP_VERSION}",
            font=FONTS["small"],
            fg=COLORS["text_muted"],
            bg=COLORS["surface"]
        )
        version_label.pack(side="left", padx=(5, 0))

        # 右側: ボタン
        btn_frame = ttk.Frame(inner, style="Header.TFrame")
        btn_frame.pack(side="right")

        # 設定ボタン
        settings_btn = tk.Button(
            btn_frame,
            text="⚙",
            font=(FONTS["body"][0], 14),
            bg=COLORS["surface"],
            fg=COLORS["text_muted"],
            relief="flat",
            cursor="hand2",
            command=self._show_settings
        )
        settings_btn.pack(side="left", padx=5)

        # 比較ボタン
        compare_btn = tk.Button(
            btn_frame,
            text="📊 比較",
            font=FONTS["body"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            padx=10,
            cursor="hand2",
            command=self._show_compare
        )
        compare_btn.pack(side="left", padx=5)

    def _create_step_indicator(self):
        """ステップインジケータを作成"""
        indicator_frame = ttk.Frame(self.root)
        indicator_frame.pack(fill="x", pady=15)

        self.step_indicator = StepIndicator(
            indicator_frame,
            on_step_click=self._on_step_click
        )
        self.step_indicator.pack()

    def _create_main_content(self):
        """メインコンテンツエリアを作成"""
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # 各ステップのフレームを事前作成
        self.step_frames = []

        # Step 1: 読込
        self.step1_frame = self._create_step1()
        self.step_frames.append(self.step1_frame)

        # Step 2: 編集・AI処理
        self.step2_frame = self._create_step2()
        self.step_frames.append(self.step2_frame)

        # Step 3: 保存
        self.step3_frame = self._create_step3()
        self.step_frames.append(self.step3_frame)

    def _create_step1(self) -> ttk.Frame:
        """Step 1: 読込画面を作成"""
        frame = ttk.Frame(self.content_frame, style="Card.TFrame")

        # DropZone
        self.drop_zone = DropZone(
            frame,
            on_file_selected=self._on_file_loaded
        )
        self.drop_zone.pack(fill="both", expand=True)

        return frame

    def _create_step2(self) -> ttk.Frame:
        """Step 2: 編集・AI処理画面を作成"""
        frame = ttk.Frame(self.content_frame, style="Card.TFrame")

        # ファイル情報
        info_frame = ttk.Frame(frame)
        info_frame.pack(fill="x", padx=10, pady=10)

        self.file_label = ttk.Label(
            info_frame,
            text="ファイル: (未選択)",
            font=FONTS["body"]
        )
        self.file_label.pack(side="left")

        # 右側ボタン群
        right_btns = ttk.Frame(info_frame)
        right_btns.pack(side="right")

        # 戻るボタン
        back_btn = tk.Button(
            right_btns,
            text="← ファイル選択に戻る",
            font=FONTS["small"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text_secondary"],
            relief="flat",
            cursor="hand2",
            command=lambda: self.step_manager.go_to(0)
        )
        back_btn.pack(side="right", padx=(5, 0))

        # Excel読込ボタン
        load_excel_btn = tk.Button(
            right_btns,
            text="📥 Excel読込",
            font=FONTS["small"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            cursor="hand2",
            command=self._load_excel_file
        )
        load_excel_btn.pack(side="right", padx=5)

        # Excel保存ボタン
        save_excel_btn = tk.Button(
            right_btns,
            text="📊 Excel保存",
            font=FONTS["small"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            cursor="hand2",
            command=self._save_excel_file
        )
        save_excel_btn.pack(side="right", padx=5)

        # JSON読込ボタン
        load_json_btn = tk.Button(
            right_btns,
            text="📥 JSON読込",
            font=FONTS["small"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            cursor="hand2",
            command=self._load_json_file
        )
        load_json_btn.pack(side="right", padx=5)

        # JSON保存ボタン
        save_json_btn = tk.Button(
            right_btns,
            text="💾 JSON保存",
            font=FONTS["small"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            cursor="hand2",
            command=self._save_json_file
        )
        save_json_btn.pack(side="right", padx=5)

        # EditableGrid
        columns = [
            ("slide", "スライド", 60, False),
            ("shape", "シェイプ", 100, False),
            ("original", "元テキスト", 300, False),
            ("text", "編集後 ✏️", 350, True),
        ]

        self.grid = EditableGrid(
            frame,
            columns=columns,
            on_change=self._on_cell_change,
            on_ai_process=self._on_ai_process
        )
        self.grid.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # 下部ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", padx=10, pady=10)

        next_btn = tk.Button(
            btn_frame,
            text="保存へ進む →",
            font=FONTS["body"],
            bg=COLORS["primary"],
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=lambda: self.step_manager.go_to(2)
        )
        next_btn.pack(side="right")

        return frame

    def _create_step3(self) -> ttk.Frame:
        """Step 3: 保存画面を作成"""
        frame = ttk.Frame(self.content_frame, style="Card.TFrame")

        inner = ttk.Frame(frame)
        inner.pack(expand=True, pady=50)

        # タイトル
        tk.Label(
            inner,
            text="💾 保存設定",
            font=FONTS["title"],
            fg=COLORS["text"],
            bg=COLORS["bg"]
        ).pack(pady=(0, 20))

        # 出力パス
        path_frame = ttk.Frame(inner)
        path_frame.pack(fill="x", pady=10)

        ttk.Label(path_frame, text="出力先:", font=FONTS["body"]).pack(side="left")
        self.output_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.output_path_var, width=50).pack(side="left", padx=10)
        ttk.Button(path_frame, text="参照...", command=self._browse_output).pack(side="left")

        # オプション
        opt_frame = ttk.Frame(inner)
        opt_frame.pack(fill="x", pady=15)

        self.backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="元ファイルをバックアップ", variable=self.backup_var).pack(anchor="w")

        # ボタン
        btn_frame = ttk.Frame(inner)
        btn_frame.pack(pady=30)

        back_btn = tk.Button(
            btn_frame,
            text="← 編集に戻る",
            font=FONTS["body"],
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            relief="flat",
            padx=15,
            pady=8,
            cursor="hand2",
            command=lambda: self.step_manager.go_to(1)
        )
        back_btn.pack(side="left", padx=10)

        save_btn = tk.Button(
            btn_frame,
            text="保存実行",
            font=FONTS["heading"],
            bg=COLORS["success"],
            fg="white",
            relief="flat",
            padx=30,
            pady=10,
            cursor="hand2",
            command=self._save_file
        )
        save_btn.pack(side="left", padx=10)

        return frame

    def _create_status_bar(self):
        """ステータスバーを作成"""
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.pack(fill="x", side="bottom")

        self.status_label = ttk.Label(
            self.status_bar,
            text="準備完了",
            font=FONTS["small"],
            foreground=COLORS["text_muted"]
        )
        self.status_label.pack(side="left", padx=10, pady=5)

    # === ステップ管理 ===

    def show_step(self, step_index: int):
        """指定されたステップを表示"""
        for i, frame in enumerate(self.step_frames):
            if i == step_index:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()

    def update_indicator(self, step_index: int):
        """インジケータを更新"""
        self.step_indicator.set_step(step_index)

    def _on_step_click(self, step_index: int):
        """ステップクリック時"""
        # Step 0は常にアクセス可能
        # Step 1以降はファイル読込後のみ
        if step_index == 0:
            self.step_manager.go_to(step_index)
        elif self.json_data:
            self.step_manager.go_to(step_index)

    # === ファイル操作 ===

    def _on_file_loaded(self, file_path: str):
        """ファイル読み込み完了時"""
        try:
            self.current_file = file_path
            self.json_data = extract_to_json(file_path)

            # グリッドにデータを読み込み
            grid_data = []
            for slide in self.json_data.get("slides", []):
                for shape in slide.get("shapes", []):
                    grid_data.append({
                        "slide": str(slide["slide"]),
                        "shape": shape["name"],
                        "original": shape["text"],
                        "text": shape["text"],
                    })

            self.grid.load_data(grid_data)

            # ファイル名を表示
            file_name = Path(file_path).name
            self.file_label.config(text=f"ファイル: {file_name}")

            # 出力パスを設定
            p = Path(file_path)
            self.output_path_var.set(str(p.parent / f"{p.stem}_edited{p.suffix}"))

            # ステップ2へ
            self.step_manager.go_to(1)
            self._set_status(f"読み込み完了: {file_name}")

        except Exception as e:
            messagebox.showerror("エラー", f"ファイル読み込みエラー:\n{e}")

    def _browse_output(self):
        """出力先を参照"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")],
            initialfile=Path(self.output_path_var.get()).name if self.output_path_var.get() else "output.pptx"
        )
        if file_path:
            self.output_path_var.set(file_path)

    def _save_json_file(self):
        """JSONを保存（外部編集用）"""
        if not self.json_data:
            messagebox.showwarning("警告", "データがありません")
            return

        # グリッドからデータを取得してJSONを更新
        self._sync_grid_to_json()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile=f"{Path(self.current_file).stem if self.current_file else 'data'}.json"
        )

        if file_path:
            save_json(self.json_data, file_path)
            self._set_status(f"JSON保存完了: {Path(file_path).name}")
            messagebox.showinfo("完了", f"JSONを保存しました:\n{file_path}\n\n外部エディタで編集後、「JSON読込」で取り込めます。")

    def _load_json_file(self):
        """JSONを読込（外部編集後の取り込み）"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json"), ("すべて", "*.*")]
        )

        if not file_path:
            return

        try:
            self.json_data = load_json(file_path)

            # グリッドにデータを読み込み
            grid_data = []
            for slide in self.json_data.get("slides", []):
                for shape in slide.get("shapes", []):
                    grid_data.append({
                        "slide": str(slide["slide"]),
                        "shape": shape["name"],
                        "original": shape.get("original", shape["text"]),
                        "text": shape["text"],
                    })

            self.grid.load_data(grid_data)

            file_name = Path(file_path).name
            self.file_label.config(text=f"ファイル: {file_name} (JSON)")
            self._set_status(f"JSON読み込み完了: {file_name}")

        except Exception as e:
            messagebox.showerror("エラー", f"JSON読み込みエラー:\n{e}")

    def _save_excel_file(self):
        """Excelを保存（外部編集用）"""
        if not self.json_data:
            messagebox.showwarning("警告", "データがありません")
            return

        # グリッドからデータを取得してJSONを更新
        self._sync_grid_to_json()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"{Path(self.current_file).stem if self.current_file else 'data'}.xlsx"
        )

        if file_path:
            try:
                save_excel(self.json_data, file_path)
                self._set_status(f"Excel保存完了: {Path(file_path).name}")
                messagebox.showinfo("完了", f"Excelを保存しました:\n{file_path}\n\n外部エディタで編集後、「Excel読込」で取り込めます。")
            except Exception as e:
                messagebox.showerror("エラー", f"Excel保存エラー:\n{e}")

    def _load_excel_file(self):
        """Excelを読込（外部編集後の取り込み）"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx"), ("すべて", "*.*")]
        )

        if not file_path:
            return

        try:
            self.json_data = load_excel(file_path)

            # グリッドにデータを読み込み
            grid_data = []
            for slide in self.json_data.get("slides", []):
                for shape in slide.get("shapes", []):
                    grid_data.append({
                        "slide": str(slide["slide"]),
                        "shape": shape.get("name", ""),
                        "original": shape.get("original", shape["text"]),
                        "text": shape["text"],
                    })

            self.grid.load_data(grid_data)

            file_name = Path(file_path).name
            self.file_label.config(text=f"ファイル: {file_name} (Excel)")
            self._set_status(f"Excel読み込み完了: {file_name}")

        except Exception as e:
            messagebox.showerror("エラー", f"Excel読み込みエラー:\n{e}")

    def _sync_grid_to_json(self):
        """グリッドの内容をJSONデータに同期"""
        if not self.json_data:
            return

        grid_data = self.grid.get_data()
        data_idx = 0

        for slide in self.json_data.get("slides", []):
            for shape in slide.get("shapes", []):
                if data_idx < len(grid_data):
                    shape["text"] = grid_data[data_idx]["text"]
                    data_idx += 1

    def _save_file(self):
        """ファイルを保存"""
        if not self.current_file or not self.json_data:
            messagebox.showwarning("警告", "保存するデータがありません")
            return

        output_path = self.output_path_var.get()
        if not output_path:
            messagebox.showwarning("警告", "出力先を指定してください")
            return

        try:
            # グリッドからデータを取得してJSONを更新
            self._sync_grid_to_json()

            # バックアップ
            if self.backup_var.get() and Path(output_path).exists():
                backup_path = str(Path(output_path).with_suffix(".backup.pptx"))
                import shutil
                shutil.copy(output_path, backup_path)

            # 保存
            apply_from_json(self.current_file, self.json_data, output_path)

            messagebox.showinfo("完了", f"保存しました:\n{output_path}")
            self._set_status(f"保存完了: {Path(output_path).name}")

        except Exception as e:
            messagebox.showerror("エラー", f"保存エラー:\n{e}")

    # === AI処理 ===

    def _on_cell_change(self, item: str, column: str, value: str):
        """セル変更時"""
        self._set_status("編集中...")

    def _on_ai_process(self, item: str, preset_name: Optional[str]):
        """AI処理"""
        def callback(prompt: str):
            self._process_ai(item, prompt)

        AIDialog(self.root, self.ai_processor, callback, preset_name)

    def _process_ai(self, item: str, prompt: str):
        """AI処理を実行"""
        # 編集可能カラム（text）の値を取得
        current_text = self.grid.tree.set(item, "text")

        def process():
            try:
                result = self.ai_processor.process_text(prompt, current_text)
                self.root.after(0, lambda: self._update_ai_result(item, result))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("エラー", str(e)))

        threading.Thread(target=process, daemon=True).start()
        self._set_status("AI処理中...")

    def _update_ai_result(self, item: str, result: str):
        """AI処理結果を反映"""
        self.grid.update_cell(item, "text", result)
        self._set_status("AI処理完了")

    # === ダイアログ ===

    def _show_settings(self):
        """設定ダイアログを表示"""
        SettingsDialog(self.root, self.ai_processor)

    def _show_compare(self):
        """比較ダイアログを表示"""
        CompareDialog(self.root, self._do_compare)

    def _do_compare(self, file1: str, file2: str, ignore_whitespace: bool, show_only_diff: bool):
        """比較を実行"""
        try:
            data1 = extract_to_json(file1)
            data2 = extract_to_json(file2)

            # 差分を計算
            diff_data, stats = self._calculate_diff(data1, data2, ignore_whitespace)

            # 結果ウィンドウを表示
            CompareResultWindow(
                self.root,
                file1, file2,
                Path(file1).name, Path(file2).name,
                diff_data, stats,
                on_apply_callback=self._apply_compare_result
            )

        except Exception as e:
            messagebox.showerror("エラー", f"比較エラー:\n{e}")

    def _calculate_diff(self, data1: Dict, data2: Dict, ignore_whitespace: bool) -> tuple:
        """差分を計算"""
        diff_data = []
        stats = {"same": 0, "changed": 0, "added": 0, "removed": 0}

        # data1のシェイプをマッピング
        map1 = {}
        for slide in data1.get("slides", []):
            for shape in slide.get("shapes", []):
                key = (slide["slide"], shape["name"])
                map1[key] = shape["text"]

        # data2のシェイプをマッピング
        map2 = {}
        for slide in data2.get("slides", []):
            for shape in slide.get("shapes", []):
                key = (slide["slide"], shape["name"])
                map2[key] = shape["text"]

        # 全キーを取得
        all_keys = set(map1.keys()) | set(map2.keys())

        for key in sorted(all_keys):
            slide_num, shape_name = key
            text1 = map1.get(key)
            text2 = map2.get(key)

            # 比較
            if text1 is None:
                status = "追加"
                stats["added"] += 1
            elif text2 is None:
                status = "削除"
                stats["removed"] += 1
            else:
                t1 = text1.strip() if ignore_whitespace else text1
                t2 = text2.strip() if ignore_whitespace else text2
                if t1 == t2:
                    status = "一致"
                    stats["same"] += 1
                else:
                    status = "変更"
                    stats["changed"] += 1

            diff_data.append({
                "slide": slide_num,
                "shape": shape_name,
                "status": status,
                "before": text1 or "",
                "after": text2 or ""
            })

        return diff_data, stats

    def _apply_compare_result(self, file_path: str, selected_data: List[Dict]):
        """比較結果を反映"""
        # TODO: 比較結果をグリッドに反映
        pass

    # === ユーティリティ ===

    def _set_status(self, message: str):
        """ステータスバーを更新"""
        self.status_label.config(text=message)
