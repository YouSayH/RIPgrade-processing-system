


# services/task_runner.py
"""
アプリケーションのメインロジックを統括するモジュール。
UIからの要求に応じて、データの読み込み、処理、保存の一連の流れを管理する。
"""

import traceback  # エラー発生時の詳細情報を取得するために使用（現在はloggingに移行）
import logging    # エラーや処理状況をログファイルに記録するために使用
from models.student import load_students_from_csv
from repositories.excel_repository import ExcelRepository
from . import transfer_service, summary_service

def run_all_tasks(tasks: dict, files: tuple, terms: list, status_callback, progress_callback):
    """
    UIからの指示に基づき、全てのビジネスタスクを統括して実行する。

    Args:
        tasks (dict): 実行するタスクのフラグ {'transfer': bool, 'grades': bool, ...}
        files (tuple): (CSVパス, Excelパス)
        terms (list): 対象学期のリスト
        status_callback (function): UIのステータス表示を更新するためのコールバック関数。
        progress_callback (function): UIのプログレスバーを更新するためのコールバック関数。

    Returns:
        (bool, str): (成功/失敗, UIへ表示する最終メッセージ) のタプル
    """
    csv_path, excel_path = files
    try:
        # --- 【通常処理】ステップ1: データの読み込み ---
        # student.pyの関数を呼び出し、CSVファイルから学生データを読み込む。
        # この時点でCSVのフォーマットが不正な場合、例外が発生し、下のexceptブロックで捕捉される。
        status_callback("ステップ1/4: CSVから学生データを読み込み中...")
        students, original_columns = load_students_from_csv(csv_path)
        progress_callback(10) # UIに進捗を通知

        # --- 【通常処理】ステップ2: 永続化層（リポジトリ）の準備 ---
        # excel_repository.pyのクラスを使い、Excelファイルを操作するための準備を行う。
        # この時点でExcelファイルが存在しない、または破損している場合、例外が発生する。
        status_callback("ステップ2/4: Excelファイルを準備中...")
        repo = ExcelRepository(excel_path)
        progress_callback(20)

        # --- 【通常処理】ステップ3: 各ビジネスタスクの実行 ---
        # UIでチェックされた処理を、対応するサービスを呼び出して順次実行する。
        status_callback("ステップ3/4: データ処理を実行中...")
        if tasks.get('transfer'):
            transfer_service.run(repo, students, terms, status_callback)
        progress_callback(50)

        if tasks.get('grades'):
            summary_service.create_grade_summary(repo, students, original_columns, status_callback)
        progress_callback(75)

        if tasks.get('attendance'):
            summary_service.create_attendance_summary(repo, students, original_columns, status_callback)
        progress_callback(90)

        # --- 【通常処理】ステップ4: 変更の保存 ---
        # これまでの処理でメモリ上で行われた変更を、実際にExcelファイルに上書き保存する。
        # この時点でファイルが他のプログラムで開かれていたり、書き込み権限がない場合、例外が発生する。
        status_callback("ステップ4/4: 変更をExcelファイルに保存しています...")
        repo.save()
        progress_callback(100)

        # --- 【通常処理】正常終了 ---
        # 全ての処理が成功した場合、UIに表示する成功メッセージを作成し、
        # 成功を示すフラグ(True)と共に返す。
        final_message = f"処理が正常に完了しました。\nファイル: {excel_path}"
        status_callback(f"\n✅ {final_message}")
        return (True, final_message)

    # --- 【エラー処理】ここから下で、tryブロック内で発生した様々なエラーを捕捉する ---

    except MemoryError:
        # メモリ不足エラー: 非常に巨大なCSV/Excelファイルを読み込もうとした場合に発生。
        error_message = "メモリ不足のため、処理を中断しました。\n処理しようとしたファイルが大きすぎる可能性があります。"
        status_callback(f"\n❌ 致命的なエラー: {error_message}")
        logging.exception("MemoryErrorが発生しました。")  # エラー詳細をログファイルに記録
        return (False, error_message) # UIに表示するメッセージを返す

    except (FileNotFoundError, KeyError, ValueError, IOError, PermissionError) as e:
        # 予測可能なエラー群:
        # FileNotFoundError: 指定されたファイルが存在しない。
        # KeyError:          CSVの必須ヘッダー（例: '氏名'）が見つからない。
        # ValueError:        CSVのデータ形式や文字コードが不正。
        # IOError:           Excelファイルが破損している、またはディスク容量不足。
        # PermissionError:   Excelファイルが他のアプリで開かれていて書き込めない。
        #
        # student.pyやexcel_repository.pyで生成された分かりやすいエラーメッセージ(e)をそのまま利用する。
        error_message = f"処理を中断しました。\n\n理由:\n{e}"
        status_callback(f"\n❌ エラーが発生しました:\n{e}")
        logging.exception(f"処理中にハンドリング済みのエラーが発生しました: {e}") # エラー詳細をログファイルに記録
        return (False, error_message)

    except Exception as e:
        # 上記以外の予期せぬエラー: プログラムのバグなど、開発者が想定していない問題。
        # これを捕捉することで、アプリケーション全体がクラッシュするのを防ぐ最終的なセーフティネット。
        error_info = f"予期せぬエラーが発生しました: {e}"
        status_callback(f"\n❌ 致命的なエラー: {error_info}")
        logging.exception("予期せぬエラーが発生しました。") # エラー詳細をログファイルに記録
        return (False, f"重大なエラーが発生しました。\n詳細はステータス欄やログファイルを確認してください。\n\n詳細情報: {e}")

    finally:
        # finallyブロックは、tryブロックが正常に終了しても、エラーで中断しても、必ず最後に実行される。
        # プログレスバーをリセットして、次の操作に備える。
        progress_callback(0)
































# # services/task_runner.py

# import traceback
# import logging
# from models.student import load_students_from_csv
# from repositories.excel_repository import ExcelRepository
# from . import transfer_service, summary_service

# def run_all_tasks(tasks: dict, files: tuple, terms: list, status_callback, progress_callback):
#     """
#     UIからの指示に基づき、全てのビジネスタスクを統括して実行する。
#     この関数がアプリケーションのメインロジックの司令塔となる。
#     """
#     csv_path, excel_path = files
#     try:
#         # --- ステップ1: データの読み込み ---
#         status_callback("ステップ1/4: CSVから学生データを読み込み中...")
#         students, original_columns = load_students_from_csv(csv_path)
#         progress_callback(10)

#         # --- ステップ2: 永続化層の準備 ---
#         status_callback("ステップ2/4: Excelファイルを準備中...")
#         repo = ExcelRepository(excel_path)
#         progress_callback(20)

#         # --- ステップ3: 各ビジネスタスクの実行 ---
#         status_callback("ステップ3/4: データ処理を実行中...")
#         if tasks.get('transfer'):
#             transfer_service.run(repo, students, terms, status_callback)
#         progress_callback(50)

#         if tasks.get('grades'):
#             summary_service.create_grade_summary(repo, students, original_columns, status_callback)
#         progress_callback(75)

#         if tasks.get('attendance'):
#             summary_service.create_attendance_summary(repo, students, original_columns, status_callback)
#         progress_callback(90)

#         # --- ステップ4: 変更の保存 ---
#         status_callback("ステップ4/4: 変更をExcelファイルに保存しています...")
#         repo.save()
#         progress_callback(100)

#         # --- 正常終了時のメッセージ作成 ---
#         final_message = f"処理が正常に完了しました。\nファイル: {excel_path}"
#         status_callback(f"\n✅ {final_message}")
#         return (True, final_message)

#     # --- エラーハンドリング (強化) ---
#     except MemoryError:
#         error_message = "メモリ不足のため、処理を中断しました。\n処理しようとしたファイルが大きすぎる可能性があります。"
#         status_callback(f"\n❌ 致命的なエラー: {error_message}")
#         logging.exception("MemoryErrorが発生しました。")
#         return (False, error_message)
#     except (FileNotFoundError, KeyError, ValueError, IOError, PermissionError) as e:
#         error_message = f"処理を中断しました。\n\n理由:\n{e}"
#         status_callback(f"\n❌ エラーが発生しました:\n{e}")
#         logging.exception(f"処理中にハンドリング済みのエラーが発生しました: {e}")
#         return (False, error_message)
#     except Exception as e:
#         error_info = f"予期せぬエラーが発生しました: {e}"
#         status_callback(f"\n❌ 致命的なエラー: {error_info}")
#         logging.exception("予期せぬエラーが発生しました。")
#         return (False, f"重大なエラーが発生しました。\n詳細はステータス欄やログファイルを確認してください。\n\n詳細情報: {e}")
#     finally:
#         # 処理が成功しても失敗しても、最後に必ずプログレスバーをリセットする
#         progress_callback(0)




























# # # services/task_runner.py

# # import traceback
# # from models.student import load_students_from_csv
# # from repositories.excel_repository import ExcelRepository
# # from . import transfer_service, summary_service

# # def run_all_tasks(tasks: dict, files: tuple, terms: list, status_callback, progress_callback):
# #     """
# #     UIからの指示に基づき、全てのビジネスタスクを統括して実行する。
# #     この関数がアプリケーションのメインロジックの司令塔となる。

# #     Args:
# #         tasks (dict): 実行するタスクのフラグ {'transfer': bool, 'grades': bool, ...}
# #         files (tuple): (CSVパス, Excelパス)
# #         terms (list): 対象学期のリスト
# #         status_callback (function): UIへの状況通知コールバック
# #         progress_callback (function): UIへの進捗通知コールバック

# #     Returns:
# #         (bool, str): (成功/失敗, UIへ表示する最終メッセージ) のタプル
# #     """
# #     csv_path, excel_path = files
# #     try:
# #         # --- ステップ1: データの読み込み ---
# #         status_callback("ステップ1/4: CSVから学生データを読み込み中...")
# #         students, original_columns = load_students_from_csv(csv_path)
# #         progress_callback(10)

# #         # --- ステップ2: 永続化層の準備 ---
# #         status_callback("ステップ2/4: Excelファイルを準備中...")
# #         repo = ExcelRepository(excel_path)
# #         progress_callback(20)

# #         # --- ステップ3: 各ビジネスタスクの実行 ---
# #         status_callback("ステップ3/4: データ処理を実行中...")
# #         if tasks.get('transfer'):
# #             transfer_service.run(repo, students, terms, status_callback)
# #         progress_callback(50)

# #         if tasks.get('grades'):
# #             summary_service.create_grade_summary(repo, students, original_columns, status_callback)
# #         progress_callback(75)

# #         if tasks.get('attendance'):
# #             summary_service.create_attendance_summary(repo, students, original_columns, status_callback)
# #         progress_callback(90)

# #         # --- ステップ4: 変更の保存 ---
# #         status_callback("ステップ4/4: 変更をExcelファイルに保存しています...")
# #         repo.save()
# #         progress_callback(100)

# #         # --- 正常終了時のメッセージ作成 ---
# #         final_message = f"処理が正常に完了しました。\nファイル: {excel_path}"
# #         status_callback(f"\n✅ {final_message}")
# #         return (True, final_message)

# #     # --- エラーハンドリング (強化) ---
# #     except MemoryError:
# #         error_message = "メモリ不足のため、処理を中断しました。\n処理しようとしたファイルが大きすぎる可能性があります。"
# #         status_callback(f"\n❌ 致命的なエラー: {error_message}")
# #         traceback.print_exc()
# #         return (False, error_message)
# #     except (FileNotFoundError, KeyError, ValueError, IOError, PermissionError) as e:
# #         # FileNotFoundError: CSV/Excelファイルが見つからない
# #         # KeyError, ValueError: CSVのヘッダー/データ形式が不正
# #         # IOError, PermissionError: Excelファイルの読み書きに関する問題
# #         error_message = f"処理を中断しました。\n\n理由:\n{e}"
# #         status_callback(f"\n❌ エラーが発生しました:\n{e}")
# #         traceback.print_exc() # コンソールに詳細なエラー情報を出力
# #         return (False, error_message)
# #     except Exception as e:
# #         # 上記以外の予期せぬエラーをキャッチ
# #         error_info = f"予期せぬエラーが発生しました: {e}"
# #         status_callback(f"\n❌ 致命的なエラー: {error_info}")
# #         traceback.print_exc() # コンソールに詳細なエラー情報を出力
# #         return (False, f"重大なエラーが発生しました。\n詳細はステータス欄を確認してください。\n\n詳細情報: {e}")
# #     finally:
# #         # 処理が成功しても失敗しても、最後に必ずプログレスバーをリセットする
# #         progress_callback(0)

















# # # services/task_runner.py

# # import traceback
# # from models.student import load_students_from_csv
# # from repositories.excel_repository import ExcelRepository
# # from . import transfer_service, summary_service

# # def run_all_tasks(tasks: dict, files: tuple, terms: list, status_callback, progress_callback):
# #     """
# #     UIからの指示に基づき、全てのビジネスタスクを統括して実行する。
# #     この関数がアプリケーションのメインロジックの司令塔となる。

# #     Args:
# #         tasks (dict): 実行するタスクのフラグ {'transfer': bool, 'grades': bool, ...}
# #         files (tuple): (CSVパス, Excelパス)
# #         terms (list): 対象学期のリスト
# #         status_callback (function): UIへの状況通知コールバック
# #         progress_callback (function): UIへの進捗通知コールバック

# #     Returns:
# #         (bool, str): (成功/失敗, UIへ表示する最終メッセージ) のタプル
# #     """
# #     csv_path, excel_path = files
# #     try:
# #         # --- ステップ1: データの読み込み ---
# #         status_callback("ステップ1/4: CSVから学生データを読み込み中...")
# #         students, original_columns = load_students_from_csv(csv_path)
# #         progress_callback(10)

# #         # --- ステップ2: 永続化層の準備 ---
# #         status_callback("ステップ2/4: Excelファイルを準備中...")
# #         repo = ExcelRepository(excel_path)
# #         progress_callback(20)

# #         # --- ステップ3: 各ビジネスタスクの実行 ---
# #         status_callback("ステップ3/4: データ処理を実行中...")
# #         if tasks.get('transfer'):
# #             transfer_service.run(repo, students, terms, status_callback)
# #         progress_callback(50)

# #         if tasks.get('grades'):
# #             summary_service.create_grade_summary(repo, students, original_columns, status_callback)
# #         progress_callback(75)

# #         if tasks.get('attendance'):
# #             summary_service.create_attendance_summary(repo, students, original_columns, status_callback)
# #         progress_callback(90)

# #         # --- ステップ4: 変更の保存 ---
# #         status_callback("ステップ4/4: 変更をExcelファイルに保存しています...")
# #         repo.save()
# #         progress_callback(100)

# #         # --- 正常終了時のメッセージ作成 ---
# #         final_message = f"処理が正常に完了しました。\nファイル: {excel_path}"
# #         status_callback(f"\n✅ {final_message}")
# #         # ★★★ 成功時は (True, メッセージ) のタプルを返す ★★★
# #         return (True, final_message)

# #     # --- エラーハンドリング ---
# #     except (FileNotFoundError, KeyError) as e:
# #         # FileNotFoundError: CSVやExcelファイルが見つからない場合
# #         # KeyError: CSVの必須列が見つからないなど、データ形式が不正な場合
# #         status_callback(f"エラー: {e}")
# #         traceback.print_exc() # コンソールに詳細なエラー情報を出力
# #         # ★★★ 失敗時も (False, メッセージ) のタプルを返す ★★★
# #         return (False, f"処理を中断しました。\n理由: {e}")
# #     except Exception as e:
# #         # 上記以外の予期せぬエラーをキャッチ
# #         error_info = f"予期せぬエラーが発生しました: {e}"
# #         status_callback(f"致命的なエラー: {error_info}")
# #         traceback.print_exc() # コンソールに詳細なエラー情報を出力
# #         # ★★★ 失敗時も (False, メッセージ) のタプルを返す ★★★
# #         return (False, f"重大なエラーが発生しました。\n詳細はステータス欄を確認してください。\n\n詳細: {e}")
# #     finally:
# #         # 処理が成功しても失敗しても、最後に必ずプログレスバーをリセットする
# #         progress_callback(0)




















