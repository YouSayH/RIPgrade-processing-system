
"""
アプリケーションのエントリーポイント（開始地点）。
このスクリプトを直接実行すると、GUIアプリケーションが起動します。
"""
# --- ライブラリのインポート ---
import tkinter as tk
from tkinterdnd2 import TkinterDnD  # ファイルのドラッグ＆ドロップ機能を提供
from ui.view import AppView          # UIの見た目を定義するクラス
from ui.viewmodel import AppViewModel    # UIの動作と状態を管理するクラス
import logging                       # エラーや処理状況をログファイルに記録する機能を提供

# --- メイン処理 ---
if __name__ == '__main__':
    """
    Pythonスクリプトが直接実行された場合にのみ、以下の処理を行うブロック。
    （他のファイルからこのファイルがインポートされた場合には実行されない）
    """
    # --- 1. ログ機能の初期設定 ---
    # アプリケーション全体で発生したエラーや主要な動作を 'error.log' ファイルに記録するための設定。
    # 問題が発生した際に、原因を追跡しやすくする目的がある。
    logging.basicConfig(
        level=logging.INFO,  # INFOレベル以上の重要度のログから記録する
        format='%(asctime)s [%(levelname)s] %(message)s',  # '日時 [ログレベル] メッセージ' の形式で記録
        filename='error.log',  # ログを保存するファイル名
        encoding='utf-8',      # 日本語の文字化けを防ぐための文字コード指定
        filemode='a'           # 'a' (append)モード: 既存のログに追記する / 'w' (write)モード: 毎回ファイルを上書きする
    )
    logging.info("アプリケーションを起動しました。")

    # --- 2. メインウィンドウの作成 ---
    # TkinterDnD.Tk() を使い、ドラッグ＆ドロップが可能なウィンドウを生成する。
    # これがアプリケーションの土台となる。
    root = TkinterDnD.Tk()
    root.title("成績処理ツール") # ウィンドウの上部に表示されるタイトル
    root.geometry("700x650")   # ウィンドウの初期サイズ（幅x高さ）700x650

    # --- 3. ViewModel（ロジック担当）のインスタンス化 ---
    # UIの裏側で動くロジックや、UIが持つべきデータ（ファイルパスなど）を管理する
    # ViewModelオブジェクトを生成する。
    # ウィンドウ自身への参照 (master=root) を渡すことで、ViewModelからウィンドウを
    # 直接操作（例: 終了処理）できるようにする。
    viewmodel = AppViewModel(master=root)

    # --- 4. View（見た目担当）のインスタンス化 ---
    # ボタンやテキストボックスなど、画面の部品を組み立てるViewオブジェクトを生成する。
    # master=root で、このViewがどのウィンドウに属するかを指定する。
    # viewmodel=viewmodel で、View（見た目）とViewModel（ロジック）を接続し、
    # ボタンが押されたらViewModelのメソッドを呼び出す、といった連携を可能にする。
    view = AppView(master=root, viewmodel=viewmodel)
    # pack()メソッドで、作成したViewをウィンドウ内に配置し、表示する。
    view.pack(fill=tk.BOTH, expand=True)

    # --- 5. ウィンドウ終了処理の上書き ---
    # ウィンドウ右上の「×」ボタンが押されたときの標準の動作を上書きする。
    # 標準では即座に終了するが、ViewModelが持つ安全確認機能付きのメソッド(request_quit)に
    # 処理を差し替えることで、処理中の意図しない終了を防ぐ。
    root.protocol("WM_DELETE_WINDOW", viewmodel.request_quit)

    # --- 6. アプリケーションの実行（イベントループ開始） ---
    # この行で、アプリケーションはユーザーからの操作（クリック、入力など）を待ち受ける状態に入る。
    # ウィンドウが閉じられるまで、プログラムはここで待機し続ける。
    try:
        root.mainloop()
    except Exception as e:
        # 通常は発生しないが、GUIのイベントループ自体で致命的なエラーが起きた場合に備え、
        # その情報をログに記録して、追跡できるようにする。
        logging.exception("GUIのメインループで予期せぬエラーが発生しました。")
    finally:
        # mainloopが終了した（＝ウィンドウが閉じられた）後に必ず実行される。
        logging.info("アプリケーションを終了しました。")
































# """
# アプリケーションのエントリーポイント（開始地点）。
# このスクリプトを直接実行すると、GUIアプリケーションが起動します。
# """
# import tkinter as tk
# from tkinterdnd2 import TkinterDnD
# from ui.view import AppView
# from ui.viewmodel import AppViewModel
# import logging

# if __name__ == '__main__':
#     """
#     アプリケーションのエントリーポイント
#     """
#     # ★★★ アプリケーション起動時にログファイルの設定を行う ★★★
#     logging.basicConfig(
#         level=logging.INFO,
#         format='%(asctime)s [%(levelname)s] %(message)s',
#         filename='error.log', # ログファイル名
#         encoding='utf-8',
#         filemode='a' # 追記モード
#     )
#     logging.info("アプリケーションを起動しました。")

#     # --- 1. ルートウィンドウの作成 ---
#     root = TkinterDnD.Tk()
#     root.title("成績処理ツール")
#     root.geometry("700x650")

#     # --- 2. ViewModelのインスタンス化 ---
#     viewmodel = AppViewModel(master=root)

#     # ★★★ ウィンドウを閉じるボタンの動作をViewModelのメソッドに差し替え ★★★
#     root.protocol("WM_DELETE_WINDOW", viewmodel.request_quit)

#     # --- 3. Viewのインスタンス化 ---
#     view = AppView(master=root, viewmodel=viewmodel)
#     view.pack(fill=tk.BOTH, expand=True)

#     # --- 4. アプリケーションの実行 ---
#     try:
#         root.mainloop()
#     except Exception as e:
#         logging.exception("GUIのメインループで予期せぬエラーが発生しました。")
#     finally:
#         logging.info("アプリケーションを終了しました。")



























# """
# アプリケーションのエントリーポイント（開始地点）。
# このスクリプトを直接実行すると、GUIアプリケーションが起動します。
# """
# import tkinter as tk
# # ドラッグ＆ドロップ機能を有効にするために、標準のTkinterの代わりにTkinterDnDをインポート
# from tkinterdnd2 import TkinterDnD
# from ui.view import AppView
# from ui.viewmodel import AppViewModel

# if __name__ == '__main__':
#     """
#     アプリケーションのエントリーポイント
#     """
#     # --- 1. ルートウィンドウの作成 ---
#     # D&D機能が有効化されたルートウィンドウを最初に生成する
#     # 通常の tk.Tk() の代わりに TkinterDnD.Tk() を使用
#     root = TkinterDnD.Tk()
#     root.title("成績処理ツール")
#     root.geometry("700x650")
#     # ウィンドウの閉じるボタンが押されたときにアプリケーションを終了させる
#     root.protocol("WM_DELETE_WINDOW", root.quit)

#     # --- 2. ViewModelのインスタンス化 ---
#     # UIの状態とロジックを管理するViewModelを生成し、ルートウィンドウを渡す
#     # (ViewModel内でTkinterの変数型を使うため、master=root が必要)
#     viewmodel = AppViewModel(master=root)

#     # --- 3. Viewのインスタンス化 ---
#     # アプリケーションの見た目を担当するView（メインフレーム）を生成し、
#     # ルートウィンドウとViewModelを渡して、UI要素とロジックを接続する
#     view = AppView(master=root, viewmodel=viewmodel)
#     # Viewをウィンドウ全体に表示
#     view.pack(fill=tk.BOTH, expand=True)

#     # --- 4. アプリケーションの実行 ---
#     # GUIのイベントループを開始し、ユーザーの操作を待ち受ける
#     root.mainloop()
