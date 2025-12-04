#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
InsightSlides License Manager
ライセンス管理ツール（管理者専用）

管理者: Erik
パスワード: admin123
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import random
import string
from datetime import datetime, timedelta
from pathlib import Path
import os

# === 設定 ===
ADMIN_USER = "administrator"
ADMIN_PASS = "admin123"
LICENSE_FILE = "licenses.json"

# === カラーパレット ===
COLORS = {
    "bg": "#F8FAFC",
    "bg_card": "#FFFFFF",
    "primary": "#2563EB",
    "success": "#10B981",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "text": "#1E293B",
    "text_secondary": "#64748B",
    "border": "#E2E8F0",
}

# === ライセンスデータ管理 ===
class LicenseStore:
    def __init__(self, filepath=LICENSE_FILE):
        self.filepath = filepath
        self.data = {"issued_keys": []}
        self.load()
    
    def load(self):
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except:
                self.data = {"issued_keys": []}
    
    def save(self):
        with open(self.filepath, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
    
    def add_key(self, key_info):
        self.data["issued_keys"].append(key_info)
        self.save()
    
    def delete_key(self, key):
        self.data["issued_keys"] = [k for k in self.data["issued_keys"] if k["key"] != key]
        self.save()
    
    def get_all_keys(self):
        return self.data["issued_keys"]
    
    def key_exists(self, key):
        return any(k["key"] == key for k in self.data["issued_keys"])
    
    def export_for_app(self):
        """本体アプリ用のキーリストをエクスポート"""
        result = {}
        for k in self.data["issued_keys"]:
            result[k["key"]] = (k["plan"], k.get("expires"))
        return result


# === キー生成 ===
def generate_key(plan: str, key_type: str, expires: str = None) -> str:
    """ライセンスキーを生成"""
    chars = string.ascii_uppercase + string.digits
    
    if key_type == "permanent":
        # 永続キー: PRO-XXXX-XXXX-XXXX
        prefix = plan.upper()[:3]
        parts = [''.join(random.choices(chars, k=4)) for _ in range(3)]
        return f"{prefix}-{'-'.join(parts)}"
    
    elif key_type == "annual":
        # 年間キー: STD-XXXX-XXXX-2025
        prefix = plan.upper()[:3]
        parts = [''.join(random.choices(chars, k=4)) for _ in range(2)]
        year = expires[:4] if expires else datetime.now().strftime("%Y")
        return f"{prefix}-{'-'.join(parts)}-{year}"
    
    else:  # trial
        # トライアル: TRIAL-XXXXXX-YYYYMMDD
        code = ''.join(random.choices(chars, k=6))
        date = expires.replace("-", "") if expires else datetime.now().strftime("%Y%m%d")
        return f"TRIAL-{code}-{date}"


# === ログイン画面 ===
class LoginWindow:
    def __init__(self, root, on_success):
        self.root = root
        self.on_success = on_success
        self.root.title("ライセンス管理 - ログイン")
        self.root.geometry("400x300")
        self.root.configure(bg=COLORS["bg"])
        self.root.resizable(False, False)
        
        # センタリング
        self.root.eval('tk::PlaceWindow . center')
        
        frame = tk.Frame(root, bg=COLORS["bg_card"], padx=30, pady=30)
        frame.pack(expand=True, fill='both', padx=20, pady=20)
        
        tk.Label(frame, text="🔐 管理者ログイン", font=("Yu Gothic UI", 14, "bold"),
                 bg=COLORS["bg_card"], fg=COLORS["text"]).pack(pady=(0, 20))
        
        # ユーザー名
        tk.Label(frame, text="ユーザー名", font=("Yu Gothic UI", 10),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(anchor='w')
        self.user_entry = tk.Entry(frame, font=("Yu Gothic UI", 11), width=30)
        self.user_entry.pack(pady=(0, 10))
        self.user_entry.insert(0, "administrator")
        
        # パスワード
        tk.Label(frame, text="パスワード", font=("Yu Gothic UI", 10),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(anchor='w')
        self.pass_entry = tk.Entry(frame, font=("Yu Gothic UI", 11), width=30, show="*")
        self.pass_entry.pack(pady=(0, 15))
        self.pass_entry.bind('<Return>', lambda e: self.login())
        
        # ログインボタン
        btn = tk.Button(frame, text="ログイン", font=("Yu Gothic UI", 11, "bold"),
                       bg=COLORS["primary"], fg="white", width=20, height=1,
                       relief='flat', cursor='hand2', command=self.login)
        btn.pack()
        
        self.pass_entry.focus()
    
    def login(self):
        if self.user_entry.get() == ADMIN_USER and self.pass_entry.get() == ADMIN_PASS:
            self.on_success()
        else:
            messagebox.showerror("エラー", "ユーザー名またはパスワードが違います")


# === メイン管理画面 ===
class LicenseManagerApp:
    def __init__(self, root):
        self.root = root
        self.store = LicenseStore()
        
        self.root.title("InsightSlides ライセンス管理")
        self.root.geometry("900x650")
        self.root.configure(bg=COLORS["bg"])
        self.root.resizable(True, True)
        
        self._create_ui()
        self._refresh_list()
    
    def _create_ui(self):
        # ヘッダー
        header = tk.Frame(self.root, bg=COLORS["primary"], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="🔐 InsightSlides ライセンス管理", 
                 font=("Yu Gothic UI", 16, "bold"), bg=COLORS["primary"], fg="white"
                ).pack(side='left', padx=20, pady=15)
        
        # メインコンテナ
        main = tk.Frame(self.root, bg=COLORS["bg"])
        main.pack(fill='both', expand=True, padx=20, pady=20)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=2)
        main.grid_rowconfigure(0, weight=1)
        
        # 左パネル：キー発行
        self._create_issue_panel(main)
        
        # 右パネル：キー一覧
        self._create_list_panel(main)
    
    def _create_issue_panel(self, parent):
        panel = tk.Frame(parent, bg=COLORS["bg_card"], relief='flat', bd=1)
        panel.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        
        # タイトル
        tk.Label(panel, text="📋 キー発行", font=("Yu Gothic UI", 13, "bold"),
                 bg=COLORS["bg_card"], fg=COLORS["text"]).pack(pady=(20, 15), padx=20, anchor='w')
        
        form = tk.Frame(panel, bg=COLORS["bg_card"], padx=20)
        form.pack(fill='x')
        
        # プラン選択
        tk.Label(form, text="プラン", font=("Yu Gothic UI", 10),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(anchor='w')
        self.plan_var = tk.StringVar(value="Pro")
        plan_combo = ttk.Combobox(form, textvariable=self.plan_var, 
                                   values=["Free", "Standard", "Pro"], state='readonly', width=25)
        plan_combo.pack(pady=(0, 10), anchor='w')
        
        # 種類選択
        tk.Label(form, text="種類", font=("Yu Gothic UI", 10),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(anchor='w')
        self.type_var = tk.StringVar(value="trial")
        type_combo = ttk.Combobox(form, textvariable=self.type_var,
                                   values=[("permanent", "永続"), ("annual", "年間"), ("trial", "トライアル")],
                                   state='readonly', width=25)
        type_combo['values'] = ["permanent (永続)", "annual (年間)", "trial (トライアル)"]
        type_combo.set("trial (トライアル)")
        type_combo.pack(pady=(0, 10), anchor='w')
        self.type_combo = type_combo
        
        # 有効期限
        tk.Label(form, text="有効期限", font=("Yu Gothic UI", 10),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(anchor='w')
        
        date_frame = tk.Frame(form, bg=COLORS["bg_card"])
        date_frame.pack(anchor='w', pady=(0, 10))
        
        # デフォルト: 30日後
        default_date = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")
        self.expires_var = tk.StringVar(value=default_date)
        self.expires_entry = tk.Entry(date_frame, textvariable=self.expires_var, 
                                       font=("Yu Gothic UI", 10), width=15)
        self.expires_entry.pack(side='left')
        
        tk.Label(date_frame, text="  (YYYY-MM-DD)", font=("Yu Gothic UI", 9),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(side='left')
        
        # クイック設定ボタン
        quick_frame = tk.Frame(form, bg=COLORS["bg_card"])
        quick_frame.pack(anchor='w', pady=(0, 10))
        
        for days, label in [(7, "7日"), (14, "14日"), (30, "30日"), (90, "90日"), (365, "1年")]:
            btn = tk.Button(quick_frame, text=label, font=("Yu Gothic UI", 9),
                           bg=COLORS["bg"], relief='flat', cursor='hand2',
                           command=lambda d=days: self._set_expires(d))
            btn.pack(side='left', padx=(0, 5))
        
        # メモ
        tk.Label(form, text="メモ（任意）", font=("Yu Gothic UI", 10),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(anchor='w')
        self.note_entry = tk.Entry(form, font=("Yu Gothic UI", 10), width=30)
        self.note_entry.pack(pady=(0, 10), anchor='w')
        self.note_entry.insert(0, "")
        
        # 発行数
        tk.Label(form, text="発行数", font=("Yu Gothic UI", 10),
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"]).pack(anchor='w')
        self.count_var = tk.StringVar(value="1")
        count_spin = tk.Spinbox(form, from_=1, to=100, textvariable=self.count_var,
                                 font=("Yu Gothic UI", 10), width=10)
        count_spin.pack(pady=(0, 20), anchor='w')
        
        # 発行ボタン
        issue_btn = tk.Button(form, text="🎫 キー発行", font=("Yu Gothic UI", 12, "bold"),
                              bg=COLORS["primary"], fg="white", width=20, height=2,
                              relief='flat', cursor='hand2', command=self._issue_keys)
        issue_btn.pack(pady=(10, 20))
        
        # 統計
        self.stats_label = tk.Label(panel, text="", font=("Yu Gothic UI", 10),
                                     bg=COLORS["bg_card"], fg=COLORS["text_secondary"])
        self.stats_label.pack(pady=10)
    
    def _create_list_panel(self, parent):
        panel = tk.Frame(parent, bg=COLORS["bg_card"], relief='flat', bd=1)
        panel.grid(row=0, column=1, sticky='nsew')
        
        # タイトル
        title_frame = tk.Frame(panel, bg=COLORS["bg_card"])
        title_frame.pack(fill='x', pady=(20, 10), padx=20)
        
        tk.Label(title_frame, text="📊 発行済みキー一覧", font=("Yu Gothic UI", 13, "bold"),
                 bg=COLORS["bg_card"], fg=COLORS["text"]).pack(side='left')
        
        # 更新ボタン
        refresh_btn = tk.Button(title_frame, text="🔄 更新", font=("Yu Gothic UI", 10),
                                bg=COLORS["bg"], relief='flat', cursor='hand2',
                                command=self._refresh_list)
        refresh_btn.pack(side='right')
        
        # リスト
        list_frame = tk.Frame(panel, bg=COLORS["bg_card"])
        list_frame.pack(fill='both', expand=True, padx=20, pady=(0, 10))
        
        # Treeview
        columns = ("key", "plan", "type", "expires", "status", "note")
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        self.tree.heading("key", text="ライセンスキー")
        self.tree.heading("plan", text="プラン")
        self.tree.heading("type", text="種類")
        self.tree.heading("expires", text="有効期限")
        self.tree.heading("status", text="状態")
        self.tree.heading("note", text="メモ")
        
        self.tree.column("key", width=220)
        self.tree.column("plan", width=70)
        self.tree.column("type", width=80)
        self.tree.column("expires", width=100)
        self.tree.column("status", width=80)
        self.tree.column("note", width=120)
        
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # アクションボタン
        btn_frame = tk.Frame(panel, bg=COLORS["bg_card"])
        btn_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        tk.Button(btn_frame, text="📋 コピー", font=("Yu Gothic UI", 10),
                  bg=COLORS["success"], fg="white", relief='flat', cursor='hand2',
                  command=self._copy_key).pack(side='left', padx=(0, 5))
        
        tk.Button(btn_frame, text="🗑 削除", font=("Yu Gothic UI", 10),
                  bg=COLORS["danger"], fg="white", relief='flat', cursor='hand2',
                  command=self._delete_key).pack(side='left', padx=(0, 5))
        
        tk.Button(btn_frame, text="📤 コードエクスポート", font=("Yu Gothic UI", 10),
                  bg=COLORS["warning"], fg="white", relief='flat', cursor='hand2',
                  command=self._export_code).pack(side='left', padx=(0, 5))
        
        tk.Button(btn_frame, text="📄 CSVエクスポート", font=("Yu Gothic UI", 10),
                  bg=COLORS["text_secondary"], fg="white", relief='flat', cursor='hand2',
                  command=self._export_csv).pack(side='left')
    
    def _set_expires(self, days):
        date = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")
        self.expires_var.set(date)
    
    def _get_key_type(self):
        val = self.type_combo.get()
        if "permanent" in val:
            return "permanent"
        elif "annual" in val:
            return "annual"
        else:
            return "trial"
    
    def _issue_keys(self):
        plan = self.plan_var.get()
        key_type = self._get_key_type()
        expires = self.expires_var.get() if key_type != "permanent" else None
        note = self.note_entry.get()
        count = int(self.count_var.get())
        
        # バリデーション
        if key_type != "permanent" and expires:
            try:
                exp_date = datetime.strptime(expires, "%Y-%m-%d")
                if exp_date < datetime.now():
                    messagebox.showerror("エラー", "有効期限が過去の日付です")
                    return
            except ValueError:
                messagebox.showerror("エラー", "日付形式が不正です (YYYY-MM-DD)")
                return
        
        issued = []
        for _ in range(count):
            # ユニークなキーを生成
            for _ in range(100):  # 最大100回試行
                key = generate_key(plan, key_type, expires)
                if not self.store.key_exists(key):
                    break
            
            key_info = {
                "key": key,
                "plan": plan,
                "type": key_type,
                "expires": expires,
                "issued_date": datetime.now().strftime("%Y-%m-%d"),
                "note": note
            }
            self.store.add_key(key_info)
            issued.append(key)
        
        self._refresh_list()
        
        # 発行したキーをクリップボードにコピー
        keys_text = "\n".join(issued)
        self.root.clipboard_clear()
        self.root.clipboard_append(keys_text)
        
        messagebox.showinfo("発行完了", 
                           f"{count}個のキーを発行しました\n\n" + 
                           f"クリップボードにコピーしました:\n{keys_text[:200]}...")
    
    def _refresh_list(self):
        # リストをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # データ読み込み
        keys = self.store.get_all_keys()
        
        stats = {"total": 0, "active": 0, "expired": 0}
        
        for k in keys:
            # 状態を判定
            status = "有効"
            if k.get("expires"):
                exp_date = datetime.strptime(k["expires"], "%Y-%m-%d")
                remaining = (exp_date - datetime.now()).days
                if remaining < 0:
                    status = "期限切れ"
                    stats["expired"] += 1
                else:
                    status = f"残{remaining}日"
                    stats["active"] += 1
            else:
                status = "永続"
                stats["active"] += 1
            
            stats["total"] += 1
            
            # 種類の表示
            type_display = {"permanent": "永続", "annual": "年間", "trial": "トライアル"}.get(k["type"], k["type"])
            
            self.tree.insert("", "end", values=(
                k["key"],
                k["plan"],
                type_display,
                k.get("expires", "-"),
                status,
                k.get("note", "")
            ))
        
        # 統計更新
        self.stats_label.config(text=f"発行総数: {stats['total']} / 有効: {stats['active']} / 期限切れ: {stats['expired']}")
    
    def _copy_key(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("選択なし", "コピーするキーを選択してください")
            return
        
        keys = []
        for item in selected:
            values = self.tree.item(item, 'values')
            keys.append(values[0])
        
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(keys))
        messagebox.showinfo("コピー", f"{len(keys)}個のキーをコピーしました")
    
    def _delete_key(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("選択なし", "削除するキーを選択してください")
            return
        
        if not messagebox.askyesno("確認", f"{len(selected)}個のキーを削除しますか？"):
            return
        
        for item in selected:
            values = self.tree.item(item, 'values')
            self.store.delete_key(values[0])
        
        self._refresh_list()
        messagebox.showinfo("削除完了", f"{len(selected)}個のキーを削除しました")
    
    def _export_code(self):
        """本体アプリに埋め込むPythonコードをエクスポート"""
        keys = self.store.export_for_app()
        
        code = "# ライセンスキー定義（自動生成）\n"
        code += f"# 生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        code += "LICENSE_KEYS = {\n"
        
        for key, (plan, expires) in keys.items():
            if expires:
                code += f'    "{key}": ("{plan}", "{expires}"),\n'
            else:
                code += f'    "{key}": ("{plan}", None),\n'
        
        code += "}\n"
        
        path = filedialog.asksaveasfilename(
            title="エクスポート先",
            defaultextension=".py",
            filetypes=[("Python", "*.py"), ("Text", "*.txt")],
            initialfile="license_keys.py"
        )
        
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(code)
            messagebox.showinfo("エクスポート完了", f"保存しました:\n{path}")
    
    def _export_csv(self):
        """CSVエクスポート"""
        keys = self.store.get_all_keys()
        
        path = filedialog.asksaveasfilename(
            title="CSVエクスポート先",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile=f"licenses_{datetime.now().strftime('%Y%m%d')}.csv"
        )
        
        if path:
            with open(path, 'w', encoding='utf-8-sig') as f:
                f.write("キー,プラン,種類,有効期限,発行日,メモ\n")
                for k in keys:
                    f.write(f'"{k["key"]}",{k["plan"]},{k["type"]},{k.get("expires", "")},{k.get("issued_date", "")},"{k.get("note", "")}"\n')
            messagebox.showinfo("エクスポート完了", f"保存しました:\n{path}")


# === メイン ===
def main():
    root = tk.Tk()
    
    def on_login_success():
        # ログイン画面を消してメイン画面を表示
        for widget in root.winfo_children():
            widget.destroy()
        root.geometry("900x650")
        LicenseManagerApp(root)
    
    LoginWindow(root, on_login_success)
    root.mainloop()


if __name__ == "__main__":
    main()
