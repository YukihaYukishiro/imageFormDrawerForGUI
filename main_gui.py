import tkinter as tk
from tkinter import ttk, messagebox
import json
from PIL import Image
import os
from datetime import datetime

# 前回作成したモジュールをインポート
from module.imageFormDrawer.imageFormDrawer import FormDrawer

# --- 設定 ---
IMAGE_PATH = "./module/imageFormDrawer/image/A4.jpg"
OUTPUT_PATH = "作成済.jpg"
FONT_PATH = "./module/imageFormDrawer/fonts/ipaexg.ttf"

# 設定ファイルのパス
IMAGE_POSITIONS_OUT = "./module/imageFormDrawer/json/image_positions-out.json"
CIRC_POSITIONS_OUT = "./module/imageFormDrawer/json/circles_positions-out.json"
IMAGE_POSITIONS_IN = "./module/imageFormDrawer/json/image_positions-in.json"
CIRC_POSITIONS_IN = "./module/imageFormDrawer/json/circles_positions-in.json"
SUBJECT_JSON = "./data/subject.json"
STUDENT_JSON = "./data/student_data.json"
PROFILE_DIR = "profile"


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("公欠届作成ツール")
        self.geometry("1100x950")

        # データの読み込み
        self.subjects = self.load_json(SUBJECT_JSON)
        self.students = self.load_json(STUDENT_JSON)

        # 左右のフォームのデータを保持する辞書
        self.form_vars = {"left": {}, "right": {}}

        # プロファイルディレクトリの確認・作成
        if not os.path.exists(PROFILE_DIR):
            os.makedirs(PROFILE_DIR)

        # UIの構築
        self.create_widgets()

        # 下部の実行ボタンフレーム
        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", padx=10, pady=10)

        # 学外申請書ボタン
        btn_out = ttk.Button(button_frame, text="学外申請書を作成", command=lambda: self.generate("out"))
        btn_out.pack(side="left", expand=True, fill="x", padx=5)

        # 学内申請書ボタン
        btn_in = ttk.Button(button_frame, text="学内申請書を作成", command=lambda: self.generate("in"))
        btn_in.pack(side="left", expand=True, fill="x", padx=5)
    def load_json(self, path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            messagebox.showwarning("警告", f"{path} が見つかりません。")
            return []

    def create_widgets(self):
        # --- プロファイル管理フレーム ---
        profile_frame = ttk.LabelFrame(self, text="プロファイル管理")
        profile_frame.pack(fill="x", padx=10, pady=(10, 5))

        # 左プロファイル
        self.create_profile_panel(profile_frame, "left")
        # 中央のセパレータ
        ttk.Separator(profile_frame, orient=tk.VERTICAL).grid(row=0, column=3, sticky="ns", padx=10, rowspan=2)
        # 右プロファイル
        self.create_profile_panel(profile_frame, "right", start_column=4)

        profile_frame.columnconfigure(1, weight=1)
        profile_frame.columnconfigure(5, weight=1)

        # --- フォーム入力メインフレーム ---
        self.create_main_forms()

    def create_profile_panel(self, parent, side, start_column=0):
        ttk.Label(parent, text=f"{side.capitalize()} Profile:").grid(row=0, column=start_column, sticky="w", padx=5)
        self.form_vars[side]["profile_name"] = tk.StringVar()
        self.form_vars[side]["profile_cb"] = ttk.Combobox(parent, textvariable=self.form_vars[side]["profile_name"])
        self.form_vars[side]["profile_cb"].grid(row=0, column=start_column + 1, sticky="ew")
        ttk.Button(parent, text="読み込み", command=lambda: self.load_profile(side)).grid(row=0, column=start_column + 2, padx=5)
        ttk.Button(parent, text="保存", command=lambda: self.save_profile(side)).grid(row=1, column=start_column + 2, padx=5)
        self.update_profile_list(side)

    def create_main_forms(self):
        # メインウィンドウを左右に分割
        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        # 左側のフレーム
        frame_left = ttk.Frame(main_pane)
        main_pane.add(frame_left, weight=1)
        self.create_form_panel(frame_left, "left")

        # 右側のフレーム
        frame_right = ttk.Frame(main_pane)
        main_pane.add(frame_right, weight=1)
        self.create_form_panel(frame_right, "right")

    def create_form_panel(self, parent, side):
        """指定された親ウィジェットに、片側分の入力フォームを作成する"""
        pad_opt = {"padx": 10, "pady": 5}

        # --- 0. フォームのタイトル ---
        form_title = "左側のフォーム" if side == "left" else "右側のフォーム"
        ttk.Label(parent, text=form_title, font=("", 12, "bold")).pack(**pad_opt)

        # --- 1. 学生情報フレーム ---
        frame_student = ttk.LabelFrame(parent, text="学生情報")
        frame_student.pack(fill="x", padx=20, pady=5)

        ttk.Label(frame_student, text="氏名:").grid(row=0, column=0, sticky="e")
        self.form_vars[side]["name"] = tk.StringVar()
        cb_name = ttk.Combobox(frame_student, textvariable=self.form_vars[side]["name"], state="readonly")
        cb_name["values"] = [d["氏名"] for d in self.students]
        cb_name.bind("<<ComboboxSelected>>", lambda event, s=side: self.on_student_select(event, s))
        cb_name.grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Label(frame_student, text="クラス:").grid(row=1, column=0, sticky="e")
        self.form_vars[side]["class"] = tk.StringVar()
        entry_class = ttk.Entry(frame_student, textvariable=self.form_vars[side]["class"], state="readonly")
        entry_class.grid(row=1, column=1, sticky="ew", padx=5)

        ttk.Label(frame_student, text="出席番号:").grid(row=2, column=0, sticky="e")
        self.form_vars[side]["student_number"] = tk.StringVar()
        entry_student_number = ttk.Entry(frame_student, textvariable=self.form_vars[side]["student_number"], state="readonly")
        entry_student_number.grid(row=2, column=1, sticky="ew", padx=5)

        ttk.Label(frame_student, text="学科:").grid(row=3, column=0, sticky="e")
        self.form_vars[side]["department"] = tk.StringVar()
        entry_department = ttk.Entry(frame_student, textvariable=self.form_vars[side]["department"], state="readonly")
        entry_department.grid(row=3, column=1, sticky="ew", padx=5)

        ttk.Label(frame_student, text="学籍番号:").grid(row=4, column=0, sticky="e")
        self.form_vars[side]["gakuseki"] = tk.StringVar()
        entry_gakuseki = ttk.Entry(frame_student, textvariable=self.form_vars[side]["gakuseki"], state="readonly")
        entry_gakuseki.grid(row=4, column=1, sticky="ew", padx=5)

        frame_student.columnconfigure(1, weight=1)
        self.form_vars[side]["entries"] = {
            "class": entry_class, "student_number": entry_student_number,
            "department": entry_department, "gakuseki": entry_gakuseki
        }

        # --- 2. 日時情報フレーム ---
        frame_date = ttk.LabelFrame(parent, text="訪問日")
        frame_date.pack(fill="x", padx=20, pady=5)

        today = datetime.now()
        ttk.Label(frame_date, text="年:").grid(row=0, column=0)
        self.form_vars[side]["year"] = tk.StringVar(value=str(today.year))
        ttk.Entry(frame_date, textvariable=self.form_vars[side]["year"], width=7).grid(row=0, column=1)

        ttk.Label(frame_date, text="月:").grid(row=0, column=2)
        self.form_vars[side]["month"] = tk.StringVar(value=str(today.month))
        ttk.Entry(frame_date, textvariable=self.form_vars[side]["month"], width=5).grid(row=0, column=3)

        ttk.Label(frame_date, text="日:").grid(row=0, column=4)
        self.form_vars[side]["day"] = tk.StringVar(value=str(today.day))
        ttk.Entry(frame_date, textvariable=self.form_vars[side]["day"], width=5).grid(row=0, column=5)

        # --- 3. 科目情報フレーム ---
        frame_subject = ttk.LabelFrame(parent, text="公欠科目")
        frame_subject.pack(fill="x", padx=20, pady=5)

        self.form_vars[side]["subjects"] = []
        subject_names = [d["name"] for d in self.subjects]

        for i in range(6):
            row = i * 2
            ttk.Label(frame_subject, text=f"科目{i+1}:").grid(row=row, column=0, sticky="e", padx=5)
            var_sub = tk.StringVar()
            cb_sub = ttk.Combobox(frame_subject, textvariable=var_sub, state="readonly", values=subject_names)
            cb_sub.grid(row=row, column=1, columnspan=3, sticky="ew", padx=5, pady=2)

            ttk.Label(frame_subject, text="担当教員:").grid(row=row+1, column=0, sticky="e", padx=5)
            var_teacher = tk.StringVar()
            cb_teacher = ttk.Combobox(frame_subject, textvariable=var_teacher, state="readonly")
            cb_teacher.grid(row=row+1, column=1, columnspan=3, sticky="ew", padx=5, pady=2)

            cb_sub.bind("<<ComboboxSelected>>", lambda event, s=side, idx=i: self.on_subject_select(event, s, idx))
            self.form_vars[side]["subjects"].append({"subject": var_sub, "teacher_var": var_teacher, "teacher_cb": cb_teacher})

        frame_subject.columnconfigure(1, weight=1)

        # --- 4. その他の情報 ---
        frame_detail = ttk.LabelFrame(parent, text="詳細")
        frame_detail.pack(fill="x", padx=20, pady=5)

        ttk.Label(frame_detail, text="訪問先/活動名:").grid(row=0, column=0, sticky="e")
        self.form_vars[side]["visit_dest"] = tk.StringVar()
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["visit_dest"]).grid(row=0, column=1, sticky="ew", columnspan=4)

        ttk.Label(frame_detail, text="最寄り駅/活動場所:").grid(row=1, column=0, sticky="e")
        self.form_vars[side]["nearest_station"] = tk.StringVar()
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["nearest_station"]).grid(row=1, column=1, sticky="ew", columnspan=4)

        ttk.Label(frame_detail, text="内容:").grid(row=2, column=0, sticky="e")
        self.form_vars[side]["reason"] = tk.StringVar(value="通院のため")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["reason"], width=40).grid(row=2, column=1, sticky="ew", columnspan=4)

        self.form_vars[side]["iki_type"] = tk.StringVar(value="b")
        ttk.Radiobutton(frame_detail, text="行き(学校発)/開始時間", variable=self.form_vars[side]["iki_type"], value="a").grid(row=3, column=0, sticky="w", padx=5)
        self.form_vars[side]["iki_a_h"] = tk.StringVar(value="8")
        self.form_vars[side]["iki_a_m"] = tk.StringVar(value="30")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["iki_a_h"], width=3).grid(row=3, column=1, sticky="e")
        ttk.Label(frame_detail, text="時").grid(row=3, column=2, sticky="w")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["iki_a_m"], width=3).grid(row=3, column=3, sticky="e")
        ttk.Label(frame_detail, text="分").grid(row=3, column=4, sticky="w")

        ttk.Radiobutton(frame_detail, text="行き(自宅から直行)", variable=self.form_vars[side]["iki_type"], value="b").grid(row=4, column=0, sticky="w", padx=5)
        self.form_vars[side]["iki_b_h"] = tk.StringVar(value="9")
        self.form_vars[side]["iki_b_m"] = tk.StringVar(value="00")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["iki_b_h"], width=3).grid(row=4, column=1, sticky="e")
        ttk.Label(frame_detail, text="時").grid(row=4, column=2, sticky="w")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["iki_b_m"], width=3).grid(row=4, column=3, sticky="e")
        ttk.Label(frame_detail, text="分").grid(row=4, column=4, sticky="w")

        self.form_vars[side]["kaeri_type"] = tk.StringVar(value="b")
        ttk.Radiobutton(frame_detail, text="帰り(学校着)/終了時間", variable=self.form_vars[side]["kaeri_type"], value="a").grid(row=5, column=0, sticky="w", padx=5)
        self.form_vars[side]["kaeri_a_h"] = tk.StringVar(value="18")
        self.form_vars[side]["kaeri_a_m"] = tk.StringVar(value="00")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["kaeri_a_h"], width=3).grid(row=5, column=1, sticky="e")
        ttk.Label(frame_detail, text="時").grid(row=5, column=2, sticky="w")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["kaeri_a_m"], width=3).grid(row=5, column=3, sticky="e")
        ttk.Label(frame_detail, text="分").grid(row=5, column=4, sticky="w")

        ttk.Radiobutton(frame_detail, text="帰り(自宅へ直帰)", variable=self.form_vars[side]["kaeri_type"], value="b").grid(row=6, column=0, sticky="w", padx=5)
        self.form_vars[side]["kaeri_b_h"] = tk.StringVar(value="17")
        self.form_vars[side]["kaeri_b_m"] = tk.StringVar(value="30")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["kaeri_b_h"], width=3).grid(row=6, column=1, sticky="e")
        ttk.Label(frame_detail, text="時").grid(row=6, column=2, sticky="w")
        ttk.Entry(frame_detail, textvariable=self.form_vars[side]["kaeri_b_m"], width=3).grid(row=6, column=3, sticky="e")
        ttk.Label(frame_detail, text="分").grid(row=6, column=4, sticky="w")

        frame_detail.columnconfigure(1, weight=1)

    def update_profile_list(self, side=None):
        """プロファイルリストを更新する"""
        try:
            profiles = [f.replace(".json", "") for f in os.listdir(PROFILE_DIR) if f.endswith(".json")]
            if side:
                self.form_vars[side]["profile_cb"]["values"] = profiles
            else: # 両方更新
                self.form_vars["left"]["profile_cb"]["values"] = profiles
                self.form_vars["right"]["profile_cb"]["values"] = profiles
        except Exception as e:
            messagebox.showwarning("警告", f"プロファイルリストの読み込みに失敗しました:\n{e}")

    def save_profile(self, side):
        """フォームの現在の内容をJSONプロファイルとして保存する"""
        profile_name = self.form_vars[side]["profile_name"].get()
        if not profile_name:
            messagebox.showerror("エラー", "プロファイル名を入力してください。")
            return

        vars = self.form_vars[side]
        profile_data = {
            "name": vars["name"].get(),
            "class": vars["class"].get(),
            "student_number": vars["student_number"].get(),
            "department": vars["department"].get(),
            "gakuseki": vars["gakuseki"].get(),
            "year": vars["year"].get(),
            "month": vars["month"].get(),
            "day": vars["day"].get(),
            "visit_dest": vars["visit_dest"].get(),
            "nearest_station": vars["nearest_station"].get(),
            "reason": vars["reason"].get(),
            "iki_type": vars["iki_type"].get(),
            "iki_a_h": vars["iki_a_h"].get(), "iki_a_m": vars["iki_a_m"].get(),
            "iki_b_h": vars["iki_b_h"].get(), "iki_b_m": vars["iki_b_m"].get(),
            "kaeri_type": vars["kaeri_type"].get(),
            "kaeri_a_h": vars["kaeri_a_h"].get(), "kaeri_a_m": vars["kaeri_a_m"].get(),
            "kaeri_b_h": vars["kaeri_b_h"].get(), "kaeri_b_m": vars["kaeri_b_m"].get(),
            "subjects": [
                {"subject": s["subject"].get(), "teacher": s["teacher_var"].get()}
                for s in vars["subjects"]
            ]
        }

        filepath = os.path.join(PROFILE_DIR, f"{profile_name}.json")
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(profile_data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"プロファイル '{profile_name}' を保存しました。")
            self.update_profile_list() # 両方のリストを更新
        except Exception as e:
            messagebox.showerror("エラー", f"プロファイルの保存に失敗しました:\n{e}")

    def load_profile(self, side):
        """JSONプロファイルからフォームの内容を復元する"""
        profile_name = self.form_vars[side]["profile_name"].get()
        if not profile_name:
            messagebox.showerror("エラー", "読み込むプロファイルを選択してください。")
            return

        filepath = os.path.join(PROFILE_DIR, f"{profile_name}.json")
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                profile_data = json.load(f)
            self.set_form_from_data(side, profile_data)
            messagebox.showinfo("成功", f"プロファイル '{profile_name}' を読み込みました。")
        except FileNotFoundError:
            messagebox.showerror("エラー", f"プロファイル '{profile_name}' が見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"プロファイルの読み込みに失敗しました:\n{e}")

    def set_form_from_data(self, side, data):
        """辞書データからフォームの値を設定する"""
        vars = self.form_vars[side]
        # 単純なStringVar
        for key in ["name", "class", "student_number", "department", "gakuseki",
                    "year", "month", "day", "visit_dest", "nearest_station", "reason",
                    "iki_type", "iki_a_h", "iki_a_m", "iki_b_h", "iki_b_m",
                    "kaeri_type", "kaeri_a_h", "kaeri_a_m", "kaeri_b_h", "kaeri_b_m"]:
            if key in data:
                vars[key].set(data[key])

        # 科目情報
        if "subjects" in data:
            for i, sub_data in enumerate(data["subjects"]):
                if i < len(vars["subjects"]):
                    subject_info = vars["subjects"][i]
                    subject_info["subject"].set(sub_data.get("subject", ""))
                    # 科目選択イベントを手動でトリガーして教員リストを更新
                    self.on_subject_select(None, side, i)
                    subject_info["teacher_var"].set(sub_data.get("teacher", ""))

    def on_student_select(self, event, side):
        """学生が選択されたら学籍番号を自動入力"""
        name = self.form_vars[side]["name"].get()
        entries = self.form_vars[side]["entries"]
        for s in self.students:
            if s["氏名"] == name:
                for key, entry_widget in entries.items():
                    entry_widget.config(state="normal")

                self.form_vars[side]["class"].set(s.get("クラス", ""))
                self.form_vars[side]["student_number"].set(s.get("出席番号", ""))
                self.form_vars[side]["department"].set(s.get("学科", ""))
                self.form_vars[side]["gakuseki"].set(s.get("学籍番号", ""))

                for key, entry_widget in entries.items():
                    entry_widget.config(state="readonly")
                break

    def on_subject_select(self, event, side, index): # eventはNoneの場合がある
        """科目が選択されたら教員名を自動入力"""
        subject_info = self.form_vars[side]["subjects"][index]
        sub_name = subject_info["subject"].get()
        teacher_cb = subject_info["teacher_cb"]
        teacher_var = subject_info["teacher_var"]

        for s in self.subjects:
            if s["name"] == sub_name and "teacher" in s:
                teacher = s["teacher"]
                # 以前の値をクリア
                teacher_cb.set('')
                teacher_cb.config(values=[])
                if isinstance(teacher, list):
                    # 複数の教員がいる場合は選択式にする
                    teacher_cb.config(values=teacher, state="readonly")
                    teacher_var.set(teacher[0]) # デフォルトで最初の教員を選択
                else:
                    # 教員が一人だけの場合はそれを設定
                    teacher_cb.config(values=[teacher], state="readonly")
                    teacher_var.set(teacher)
                break
        else: # ループがbreakしなかった場合（科目が空欄など）
            teacher_cb.set('')
            teacher_cb.config(values=[], state="readonly")


    def get_form_data(self, side):
        """GUIの入力値から、FormDrawer用の辞書を作成する"""
        vars = self.form_vars[side]
        data = {
            "氏名": vars["name"].get(),
            "学籍番号": vars["gakuseki"].get(),
            "クラス": vars["class"].get(),
            "出席番号": vars["student_number"].get(),
            "学科": vars["department"].get(),
            "訪問日_年": vars["year"].get(),
            "訪問日_月": vars["month"].get(),
            "訪問日_日": vars["day"].get(),
            "訪問先": vars["visit_dest"].get(),
            "最寄り駅": vars["nearest_station"].get(),
            "内容": vars["reason"].get(),
        }

        try:
            year = int(vars["year"].get())
            month = int(vars["month"].get())
            day = int(vars["day"].get())
            weekday_str = "月火水木金土日"[datetime(year, month, day).weekday()]
            data["訪問日_曜日"] = weekday_str
        except (ValueError, TypeError):
            data["訪問日_曜日"] = ""

        for i, sub_var in enumerate(vars["subjects"], 1):
            data[f"科目{i}_科目名"] = sub_var["subject"].get()
            data[f"科目{i}_担当教員"] = sub_var["teacher_var"].get()

        iki_type = vars["iki_type"].get()
        kaeri_type = vars["kaeri_type"].get()

        data["行き_区分"] = iki_type
        data["帰り_区分"] = kaeri_type

        if iki_type == 'a':
            data["行きa_時間"] = vars["iki_a_h"].get()
            data["行きa_分"] = vars["iki_a_m"].get()
        else:
            data["行きb_時間"] = vars["iki_b_h"].get()
            data["行きb_分"] = vars["iki_b_m"].get()

        if kaeri_type == 'a':
            data["帰りa_時間"] = vars["kaeri_a_h"].get()
            data["帰りa_分"] = vars["kaeri_a_m"].get()
        else:
            data["帰りb_時間"] = vars["kaeri_b_h"].get()
            data["帰りb_分"] = vars["kaeri_b_m"].get()

        return data

    def generate(self, application_type): # 'out' or 'in'
        try:
            if application_type == "out":
                image_pos_path = IMAGE_POSITIONS_OUT
                circ_pos_path = CIRC_POSITIONS_OUT
            else: # "in"
                image_pos_path = IMAGE_POSITIONS_IN
                circ_pos_path = CIRC_POSITIONS_IN

            with open(image_pos_path, "r", encoding="utf-8") as f:
                image_pos = json.load(f)
            with open(circ_pos_path, "r", encoding="utf-8") as f:
                circ_pos = json.load(f)

            form_data_left = self.get_form_data("left")
            form_data_right = self.get_form_data("right")

            img = Image.open(IMAGE_PATH)
            drawer = FormDrawer(FONT_PATH)

            drawer.draw(img, form_data_left, image_pos["left"], circ_pos["left"])
            drawer.draw(img, form_data_right, image_pos["right"], circ_pos["right"])

            img.save(OUTPUT_PATH)
            messagebox.showinfo("成功", f"画像を作成しました:\n{OUTPUT_PATH}")

            os.startfile(OUTPUT_PATH)

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")


if __name__ == "__main__":
    app = Application()
    app.mainloop()
