import tkinter as tk
from tkinter import ttk, messagebox
import json
from PIL import Image
import os
from datetime import datetime

# 前回作成したモジュールをインポート
from module.imageFormDrawer import FormDrawer

# --- 設定 ---
IMAGE_PATH = '学外公欠申請書.jpg'
OUTPUT_PATH = '学外公欠申請書_作成済.jpg'
FONT_PATH = "C:\\Windows\\Fonts\\meiryo.ttc"

# 設定ファイルのパス
POS_LEFT = 'positions_left.json'
CIRC_LEFT = 'circles_left.json'
POS_RIGHT = 'positions_right.json'
CIRC_RIGHT = 'circles_right.json'

SUBJECT_JSON = 'subject.json'
STUDENT_JSON = 'student_data.json'

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("公欠届作成ツール")
        self.geometry("500x650")

        # データの読み込み
        self.subjects = self.load_json(SUBJECT_JSON)
        self.students = self.load_json(STUDENT_JSON)

        # UIの構築
        self.create_widgets()

    def load_json(self, path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            messagebox.showwarning("警告", f"{path} が見つかりません。")
            return []

    def create_widgets(self):
        # --- スタイル設定 ---
        pad_opt = {'padx': 10, 'pady': 5}
        
        # --- 1. 学生情報フレーム ---
        frame_student = ttk.LabelFrame(self, text="学生情報")
        frame_student.pack(fill="x", **pad_opt)

        ttk.Label(frame_student, text="氏名:").grid(row=0, column=0, sticky="e")
        
        # 氏名プルダウン
        self.var_name = tk.StringVar()
        self.cb_name = ttk.Combobox(frame_student, textvariable=self.var_name, state="readonly")
        self.cb_name['values'] = [d['name'] for d in self.students]
        self.cb_name.bind("<<ComboboxSelected>>", self.on_student_select)
        self.cb_name.grid(row=0, column=1, sticky="w")

        ttk.Label(frame_student, text="学籍番号:").grid(row=1, column=0, sticky="e")
        self.var_id = tk.StringVar()
        self.entry_id = ttk.Entry(frame_student, textvariable=self.var_id)
        self.entry_id.grid(row=1, column=1, sticky="w")

        # --- 2. 日時情報フレーム ---
        frame_date = ttk.LabelFrame(self, text="申請日時")
        frame_date.pack(fill="x", **pad_opt)

        # 今日の日付をデフォルトに
        today = datetime.now()
        
        ttk.Label(frame_date, text="月:").grid(row=0, column=0)
        self.var_month = tk.StringVar(value=str(today.month))
        ttk.Entry(frame_date, textvariable=self.var_month, width=5).grid(row=0, column=1)

        ttk.Label(frame_date, text="日:").grid(row=0, column=2)
        self.var_day = tk.StringVar(value=str(today.day))
        ttk.Entry(frame_date, textvariable=self.var_day, width=5).grid(row=0, column=3)

        # --- 3. 科目情報フレーム ---
        frame_subject = ttk.LabelFrame(self, text="公欠科目")
        frame_subject.pack(fill="x", **pad_opt)

        ttk.Label(frame_subject, text="科目名:").grid(row=0, column=0, sticky="e")
        self.var_subject = tk.StringVar()
        self.cb_subject = ttk.Combobox(frame_subject, textvariable=self.var_subject, state="readonly")
        self.cb_subject['values'] = [d['name'] for d in self.subjects]
        self.cb_subject.bind("<<ComboboxSelected>>", self.on_subject_select)
        self.cb_subject.grid(row=0, column=1, sticky="w")

        ttk.Label(frame_subject, text="担当教員:").grid(row=1, column=0, sticky="e")
        self.var_teacher = tk.StringVar()
        self.entry_teacher = ttk.Entry(frame_subject, textvariable=self.var_teacher)
        self.entry_teacher.grid(row=1, column=1, sticky="w")
        
        # --- 4. その他の情報 ---
        frame_detail = ttk.LabelFrame(self, text="詳細")
        frame_detail.pack(fill="x", **pad_opt)

        ttk.Label(frame_detail, text="理由:").grid(row=0, column=0, sticky="e")
        self.var_reason = tk.StringVar(value="通院のため")
        ttk.Entry(frame_detail, textvariable=self.var_reason, width=40).grid(row=0, column=1)

        # 交通手段（区分）の選択
        ttk.Label(frame_detail, text="行き手段:").grid(row=1, column=0, sticky="e")
        self.var_iki_type = tk.StringVar(value="a") # a:バス等, b:徒歩等 (JSONのキーに合わせる)
        frame_rad_iki = ttk.Frame(frame_detail)
        frame_rad_iki.grid(row=1, column=1, sticky="w")
        ttk.Radiobutton(frame_rad_iki, text="バス・電車(a)", variable=self.var_iki_type, value="a").pack(side="left")
        ttk.Radiobutton(frame_rad_iki, text="その他(b)", variable=self.var_iki_type, value="b").pack(side="left")

        # --- 5. 実行ボタン ---
        btn_run = ttk.Button(self, text="画像を作成", command=self.generate)
        btn_run.pack(pady=20, fill="x", padx=20)

    def on_student_select(self, event):
        """学生が選択されたら学籍番号を自動入力"""
        name = self.var_name.get()
        for s in self.students:
            if s['name'] == name:
                self.var_id.set(s['id'])
                break

    def on_subject_select(self, event):
        """科目が選択されたら教員名を自動入力"""
        sub_name = self.var_subject.get()
        for s in self.subjects:
            if s['name'] == sub_name:
                self.var_teacher.set(s['teacher'])
                break

    def get_form_data(self):
        """GUIの入力値から、FormDrawer用の辞書を作成する"""
        
        # 基本的なマッピング
        # GUIの入力値をJSONのキー（以前のdata.jsonのキー）に変換
        data = {
            "氏名": self.var_name.get(),
            "学籍番号": self.var_id.get(),
            "申請_月": self.var_month.get(),
            "申請_日": self.var_day.get(),
            "理由": self.var_reason.get(),
            
            # 科目情報のマッピング (例：1限目として登録する場合)
            "科目1_科目名": self.var_subject.get(),
            "科目1_教員名": self.var_teacher.get(),
            # ※必要に応じて科目2,3...もGUIに追加してください
            
            # 行き・帰りの設定
            "行き_区分": self.var_iki_type.get(),
            "帰り_区分": self.var_iki_type.get(), # 簡易的に行きと同じにしています
            
            # 時間のサンプル（GUIに入力欄を作れば動的にできます）
            "行き_時間": "8", "行き_分": "30",
            "帰り_時間": "17", "帰り_分": "00"
        }
        return data

    def generate(self):
        try:
            # 1. コンフィグ読み込み
            with open(POS_LEFT, 'r', encoding='utf-8') as f: pos_left = json.load(f)
            with open(CIRC_LEFT, 'r', encoding='utf-8') as f: circ_left = json.load(f)
            with open(POS_RIGHT, 'r', encoding='utf-8') as f: pos_right = json.load(f)
            with open(CIRC_RIGHT, 'r', encoding='utf-8') as f: circ_right = json.load(f)

            # 2. GUIからデータ取得
            form_data = self.get_form_data()

            # 3. 画像処理
            img = Image.open(IMAGE_PATH)
            drawer = FormDrawer(FONT_PATH)

            # 左側と右側に同じデータを書き込む（必要に応じて書き分けることも可能）
            drawer.draw(img, form_data, pos_left, circ_left)
            drawer.draw(img, form_data, pos_right, circ_right)

            # 4. 保存
            img.save(OUTPUT_PATH)
            messagebox.showinfo("成功", f"画像を作成しました:\n{OUTPUT_PATH}")
            
            # 自動で開く（オプション）
            os.startfile(OUTPUT_PATH)

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")

if __name__ == "__main__":
    app = Application()
    app.mainloop()