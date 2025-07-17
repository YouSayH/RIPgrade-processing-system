

# ui/view.py

import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES  # ドラッグ＆ドロップ機能のためにインポート
import os
import sys

# このViewのロジックを担当するViewModelをインポート
from .viewmodel import AppViewModel

class AppView(ttk.Frame):
    """
    アプリケーションの見た目(View)を定義・構築するクラス。
    ttk.Frameを継承しており、自身がウィンドウ上の1つの大きな部品として振る舞う。
    ロジック（どう動くか）はすべてViewModelに委譲する。
    """
    def __init__(self, master: tk.Tk, viewmodel: AppViewModel):
        super().__init__(master)
        self.vm = viewmodel  # ViewModelへの参照をインスタンス変数として保持

        # --- 【エラー処理】テーマライブラリの適用 ---
        # ttkthemesライブラリが存在すれば、UIの見た目をモダンなスタイルに変更する。
        # 存在しなくてもエラーで停止せず、標準のスタイルでアプリケーションが動作するように
        # try...except構文で囲む（Graceful Fallback）。
        try:
            from ttkthemes import ThemedStyle
            style = ThemedStyle(self)
            style.set_theme("arc")  # "arc"テーマを適用
        except ImportError:
            # ライブラリが見つからなかった場合は何もしない
            pass

        # --- ウィジェットの構築と配置 ---
        self._setup_widgets()
        
        # --- ViewModelの変更をViewに反映させる設定 (データバインディング) ---
        self._bind_viewmodel_to_view()

        # --- D&Dとコマンドライン引数のハンドリング設定 ---
        self._setup_dnd()
        self._handle_command_line_args()

    def _setup_widgets(self):
        """UIの部品（ウィジェット）を生成し、画面に配置する。"""
        # --- メインフレーム ---
        # 全てのウィジェットを乗せる土台となるフレーム
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- ファイル選択エリア ---
        csv_lf = ttk.LabelFrame(main_frame, text="① CSVファイル (成績データ)", padding=(10,5))
        csv_lf.pack(fill=tk.X, padx=5, pady=5)
        # textvariableにViewModelの変数を指定することで、値が自動的に同期される
        csv_entry = ttk.Entry(csv_lf, textvariable=self.vm.csv_path)
        csv_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5), pady=5)
        
        excel_lf = ttk.LabelFrame(main_frame, text="② Excelファイル (転記先)", padding=(10,5))
        excel_lf.pack(fill=tk.X, padx=5, pady=(5, 10))
        excel_entry = ttk.Entry(excel_lf, textvariable=self.vm.excel_path)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5), pady=5)

        # --- ファイル参照ボタン ---
        file_btn_frame = ttk.Frame(main_frame)
        file_btn_frame.pack(fill=tk.X, padx=5)
        # commandにViewModelのメソッドを指定し、クリック時の動作を委譲する
        ttk.Button(file_btn_frame, text="CSVファイルを参照...", command=self.vm.select_csv_file).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,2))
        ttk.Button(file_btn_frame, text="Excelファイルを参照...", command=self.vm.select_excel_file).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2,0))

        # --- 実行処理選択エリア ---
        process_lf = ttk.LabelFrame(main_frame, text="③ 実行する処理を選択", padding=(10, 5))
        process_lf.pack(fill=tk.X, padx=5, pady=(10,5))
        # variableにViewModelの変数を指定することで、チェック状態を同期させる
        cb_transfer = ttk.Checkbutton(process_lf, text="成績一覧を更新", variable=self.vm.process_transfer, command=self._toggle_term_widgets)
        cb_transfer.grid(row=0, column=0, sticky="w", padx=10, pady=5)
        cb_grades = ttk.Checkbutton(process_lf, text="評定一覧を作成", variable=self.vm.process_grades)
        cb_grades.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        cb_attendance = ttk.Checkbutton(process_lf, text="出席状況一覧を作成", variable=self.vm.process_attendance)
        cb_attendance.grid(row=0, column=2, sticky="w", padx=10, pady=5)

        # --- 対象学期選択エリア ---
        self.term_lf = ttk.LabelFrame(main_frame, text="④ 成績一覧の対象学期", padding=(10,5))
        self.term_lf.pack(fill=tk.X, padx=5, pady=5)
        self.cb_zenki = ttk.Checkbutton(self.term_lf, text="前期", variable=self.vm.term_zenki)
        self.cb_zenki.pack(side=tk.LEFT, padx=10, pady=5)
        self.cb_tsuki = ttk.Checkbutton(self.term_lf, text="通期", variable=self.vm.term_tsuki)
        self.cb_tsuki.pack(side=tk.LEFT, padx=10, pady=5)

        # --- アクションエリア（実行ボタン、プログレスバー） ---
        action_frame = ttk.Frame(main_frame, padding=(0,10))
        action_frame.pack(fill=tk.X, pady=10)
        self.run_button = ttk.Button(action_frame, text="実行開始", command=self.vm.start_processing, style="Accent.TButton")
        self.run_button.pack(pady=5)
        s = ttk.Style()
        s.configure("Accent.TButton", font=('Helvetica', 10, 'bold'), padding=6)
        
        progress_bar = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate", variable=self.vm.progress_value)
        progress_bar.pack(fill=tk.X, padx=5, pady=5)

        # --- 処理状況表示エリア ---
        status_lf = ttk.LabelFrame(main_frame, text="処理状況", padding=(10,5))
        status_lf.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # ユーザーに編集させないテキストエリア
        self.status_text = tk.Text(status_lf, height=10, wrap=tk.WORD, relief="sunken", borderwidth=1, font=(" Meiryo UI", 9) if os.name == 'nt' else ("TkDefaultFont", 9))
        status_scrollbar = ttk.Scrollbar(status_lf, orient=tk.VERTICAL, command=self.status_text.yview)
        self.status_text.config(yscrollcommand=status_scrollbar.set)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.insert(tk.END, "ファイルと実行したい処理を選択し、「実行開始」ボタンを押してください。\n（ウィンドウ内のどこにでもファイルをドラッグ＆ドロップできます）\n")
        self.status_text.config(state=tk.DISABLED) # 初期状態は読み取り専用

    def _bind_viewmodel_to_view(self):
        """ViewModelの状態変更を監視し、対応するViewの更新処理を呼び出す設定。"""
        # self.vm.is_running の値が書き換わったら(trace_add 'write')、self._toggle_run_button_state を呼び出す
        self.vm.is_running.trace_add('write', self._toggle_run_button_state)
        # self.vm.status_text_log の値が書き換わったら、self._update_status_log を呼び出す
        self.vm.status_text_log.trace_add('write', self._update_status_log)

    def _setup_dnd(self):
        """ドラッグ＆ドロップの有効化とイベントのバインド"""
        # このViewが配置されているトップレベルウィンドウ（メインウィンドウ）を取得
        toplevel = self.winfo_toplevel()
        # メインウィンドウをファイルのドロップ先として登録
        toplevel.drop_target_register(DND_FILES)
        # ファイルがドロップされた際のイベント(<<Drop>>)と、実行するメソッド(_on_file_drop)を関連付ける
        toplevel.dnd_bind('<<Drop>>', self._on_file_drop)

    def _on_file_drop(self, event):
        """ファイルがドロップされたときに実行されるイベントハンドラ"""
        # event.data にはドロップされたファイルのパスが文字列として格納されている
        # 複数のファイルがドロップされると '{パス1} {パス2}' のような形式になるため、解析処理を行う
        filepaths_str = event.data.strip()
        if filepaths_str.startswith('{') and filepaths_str.endswith('}'):
            paths = [p.strip() for p in filepaths_str[1:-1].split('} {')]
        else:
            paths = filepaths_str.split()
        # 解析したパスのリストをViewModelに渡して、実際の処理を依頼
        self.vm.process_dropped_files(paths)

    def _handle_command_line_args(self):
        """コマンドライン引数として渡されたファイルを処理する"""
        # sys.argv[1:] で、スクリプト名以降の引数をリストとして取得
        args = sys.argv[1:]
        if args:
            self.vm.process_dropped_files(args)
    
    # --- ViewModelからの通知で実行されるUI更新メソッド ---
    
    def _toggle_term_widgets(self, *args):
        """「成績一覧を更新」チェックボックスの状態に応じて、学期選択UIの有効/無効を切り替える"""
        state = tk.NORMAL if self.vm.process_transfer.get() else tk.DISABLED
        self.cb_zenki.config(state=state)
        self.cb_tsuki.config(state=state)

    def _toggle_run_button_state(self, *args):
        """処理中フラグ(is_running)に応じて、実行ボタンの有効/無効を切り替える"""
        state = tk.DISABLED if self.vm.is_running.get() else tk.NORMAL
        self.run_button.config(state=state)
    
    def _update_status_log(self, *args):
        """ViewModelのステータスログ変数の内容を、画面のテキストエリアに反映させる"""
        # tk.Textウィジェットは、内容を変更するために一時的に state を 'normal' にする必要がある
        self.status_text.config(state=tk.NORMAL)
        # テキストエリアの内容を一旦すべて削除
        self.status_text.delete(1.0, tk.END)
        # ViewModelが保持している最新のログを挿入
        self.status_text.insert(tk.END, self.vm.status_text_log.get())
        # 自動で最下部にスクロールする
        self.status_text.see(tk.END)
        # 再び読み取り専用に戻す
        self.status_text.config(state=tk.DISABLED)

