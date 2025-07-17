
# ui/viewmodel.py

import tkinter as tk
from tkinter import filedialog, messagebox  # ファイル選択ダイアログとメッセージボックス機能
import threading  # 重い処理をバックグラウンドで実行し、UIのフリーズを防ぐ
import os         # ファイルパスの操作や存在確認
import sys        # コマンドライン引数の取得

# 実際のデータ処理ロジックをインポート
from services import task_runner

class AppViewModel:
    """
    UIの状態(State)と操作(Logic)を管理するクラス (ViewModel)。
    MVVM (Model-View-ViewModel) アーキテクチャパターンにおけるViewModelの役割を担う。
    - View (view.py): 見た目の定義。ViewModelへの操作を指示する。
    - ViewModel (このファイル): Viewからの指示を受け、UIの状態を更新し、Modelを呼び出す。
    - Model (services/, repositories/): ビジネスロジックやデータ永続化。
    """
    def __init__(self, master):
        """
        ViewModelの初期化。UIの各部品が持つべき「状態」をTkinterの変数として定義する。
        master(ルートウィンドウ)への参照は、ウィンドウ自体を操作（例: 終了）するために必要。
        """
        # --- インスタンス変数の設定 ---
        self.master = master  # ウィンドウを閉じる処理などで使用

        # --- UIの状態を保持するプロパティ (Tkinter Variable) ---
        # これらはViewのウィジェットと「データバインディング」され、値が変わるとUIも自動的に更新される。
        self.csv_path = tk.StringVar(master=master)
        self.excel_path = tk.StringVar(master=master)
        self.status_text_log = tk.StringVar(master=master)   # ステータスログ表示用
        self.progress_value = tk.DoubleVar(master=master, value=0.0) # プログレスバーの進捗

        # チェックボックスの状態
        self.process_transfer = tk.BooleanVar(master=master, value=True)
        self.process_grades = tk.BooleanVar(master=master, value=True)
        self.process_attendance = tk.BooleanVar(master=master, value=True)
        self.term_zenki = tk.BooleanVar(master=master, value=True)
        self.term_tsuki = tk.BooleanVar(master=master, value=True)

        # 処理が実行中かどうかを示すフラグ。二重実行の防止や、安全な終了処理に使われる。
        self.is_running = tk.BooleanVar(master=master, value=False)

    # --- UIからのイベントに対応するメソッド群 ---

    def request_quit(self):
        """
        ウィンドウの「×」ボタンが押されたときに呼ばれる、安全な終了処理。
        """
        # 処理が実行中の場合
        if self.is_running.get():
            # ユーザーに本当に終了して良いか確認する。ファイル破損のリスクを伝える。
            if messagebox.askyesno("確認", "処理を実行中です。本当にアプリケーションを終了しますか？\n(ファイルが破損する可能性があります)"):
                self.master.destroy()  # 「はい」が押されたらウィンドウを破棄して終了
        else:
            # 処理が実行中でなければ、そのまま終了する。
            self.master.destroy()

    def select_csv_file(self):
        """「CSVファイルを参照...」ボタンのコマンド"""
        # ファイル選択ダイアログを開き、選択されたファイルのパスを取得する。
        path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        # パスが取得できたら（キャンセルされなかったら）、状態変数に設定する。
        if path:
            self.csv_path.set(path)

    def select_excel_file(self):
        """「Excelファイルを参照...」ボタンのコマンド"""
        path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excelファイル", "*.xlsx;*.xlsm"), ("すべてのファイル", "*.*")]
        )
        if path:
            self.excel_path.set(path)

    def process_dropped_files(self, dropped_paths: list):
        """ドラッグ＆ドロップされたファイルのパスを処理する"""
        found_csv = False
        found_excel = False
        # ドロップされた各ファイルパスについてループ
        for path in dropped_paths:
            clean_path = path.strip()  # パス前後の空白を除去
            if not os.path.exists(clean_path): continue # ファイルが存在しなければスキップ

            # 拡張子を見て、CSVファイルかExcelファイルかを判断し、対応する変数に設定する。
            if not found_csv and clean_path.lower().endswith('.csv'):
                self.csv_path.set(clean_path)
                found_csv = True
            elif not found_excel and clean_path.lower().endswith(('.xlsx', '.xlsm')):
                self.excel_path.set(clean_path)
                found_excel = True
            
            if found_csv and found_excel: break # 両方見つかったらループを抜ける

    def start_processing(self):
        """「実行開始」ボタンのメインコマンド"""
        if self.is_running.get(): return # 処理中なら何もしない（二重実行防止）

        # --- 【エラー処理①】入力検証 (Validation) ---
        # 処理を開始する前に、必要な条件が満たされているかチェックする。
        # 問題があればユーザーにメッセージを出し、処理を中断する。
        if not self.csv_path.get() or not os.path.exists(self.csv_path.get()):
            messagebox.showerror("入力エラー", "有効なCSVファイルが指定されていません。")
            return
        if not self.excel_path.get() or not os.path.exists(self.excel_path.get()):
            messagebox.showerror("入力エラー", "有効なExcelファイルが指定されていません。")
            return

        tasks = {
            'transfer': self.process_transfer.get(),
            'grades': self.process_grades.get(),
            'attendance': self.process_attendance.get()
        }
        if not any(tasks.values()):
            messagebox.showerror("入力エラー", "実行する処理を少なくとも1つ選択してください。")
            return

        selected_terms = []
        if tasks['transfer']:
            if self.term_zenki.get(): selected_terms.append("前期")
            if self.term_tsuki.get(): selected_terms.append("通期")
            if not selected_terms:
                messagebox.showerror("入力エラー", "「成績一覧を更新」が選択されていますが、対象学期が未選択です。")
                return

        # --- 処理の準備とバックグラウンド実行 ---
        self.is_running.set(True) # 処理中フラグを立てる
        self.status_text_log.set("") # ステータスログをクリア
        self.progress_value.set(0)   # プログレスバーをリセット

        # 重い処理(task_runner)を別スレッドで実行する。
        # これにより、処理中でもUIが固まらず、応答可能な状態を保つ。
        # daemon=True は、メインウィンドウが閉じられたらこのスレッドも強制終了する設定。
        thread = threading.Thread(
            target=self._run_in_thread,
            args=(tasks, selected_terms),
            daemon=True
        )
        thread.start() # スレッドを開始

    def _run_in_thread(self, tasks, terms):
        """バックグラウンドスレッドで実行される実処理"""
        
        # --- スレッドからUIへ安全に情報を送るためのコールバック関数 ---
        def status_callback(message: str):
            # ログメッセージを追記する
            self.status_text_log.set(self.status_text_log.get() + message + "\n")

        def progress_callback(value: int):
            # プログレスバーの値を更新する
            self.progress_value.set(float(value))

        files = (self.csv_path.get(), self.excel_path.get())

        # --- データ処理の本体(task_runner)を呼び出し、結果を受け取る ---
        success, final_message = task_runner.run_all_tasks(
            tasks, files, terms, status_callback, progress_callback
        )

        # --- 【エラー処理②】結果の表示 ---
        # task_runnerからの結果(成功/失敗)に応じて、ユーザーに最終メッセージを表示する。
        if success:
            messagebox.showinfo("完了", final_message)
        else:
            # 失敗した場合、final_messageにはエラーの理由が入っている。
            messagebox.showerror("エラー", final_message)

        # 処理が完了したので、実行中フラグを降ろす。
        self.is_running.set(False)

