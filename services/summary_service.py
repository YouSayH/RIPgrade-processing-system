



# services/summary_service.py
"""
各種の集計シートを作成するビジネスロジックを提供するモジュール。
task_runnerからの指示に基づき、データモデル(Student)と永続化層(ExcelRepository)を
組み合わせて、特定の目的のExcelシートを生成する。
"""

from typing import List
from collections import OrderedDict  # 順序を保持したまま重複を削除するために使用
import pandas as pd
from models.student import Student
from repositories.excel_repository import ExcelRepository
import config

# --- 評定一覧サービス ---

def create_grade_summary(repository: ExcelRepository, students: List[Student], original_columns: pd.Index, status_callback):
    """
    「評定一覧」シートを作成するサービス。
    元のCSVファイルの科目順を維持して一覧を作成する。

    【エラーハンドリングについて】
    この関数内では、`try...except`によるエラー捕捉は行いません。
    これは設計上の意図であり、リポジトリ層（例: Excelシート作成時）で発生したエラーは、
    呼び出し元である`task_runner.py`まで伝播させ、そこで一元的にエラーハンドリングを行います。

    Args:
        repository (ExcelRepository): データ書き込みを担当するリポジトリ。
        students (List[Student]): 学生データのリスト。
        original_columns (pd.Index): 元のCSVの列情報。科目の順序を維持するために使用。
        status_callback (function): UIへの状況通知コールバック。
    """
    status_callback("\n処理中: 評定一覧シートを作成しています...")

    # --- 1. ヘッダーとなる科目リストの抽出 ---
    # 元のCSVの列情報から、ヘッダーの2行目が「評」である列の科目名（1行目）を抽出する。
    grade_cols_subjects = [col[0] for col in original_columns if len(col) == 2 and col[1] == config.KEY_GRADE]
    # 抽出した科目リストから重複を削除しつつ、元の順序は維持する (OrderedDictの特性を利用)。
    subjects = list(OrderedDict.fromkeys(grade_cols_subjects))

    # --- 2. 入力値の検証 ---
    # 評定データを持つ科目が一つも見つからない場合は、警告を出力して処理を中断する。
    if not subjects:
        status_callback("警告(評定一覧): 評定データを持つ科目が見つかりません。スキップします。")
        return

    # --- 3. Excelに出力するデータの作成 ---
    # ヘッダー行を作成: ['学籍番号', '氏名', '科目A', '科目B', ...]
    headers = ['学籍番号', '氏名'] + subjects

    # データ行を作成: 各学生について、科目リストの順に評定データを格納していく。
    data_rows = []
    for s in students:
        row = [s.id, s.name] # 各行の先頭は学籍番号と氏名
        for sub in subjects:
            # studentオブジェクトのscores辞書から、(科目名, "評") をキーにして評定を取得。
            # 存在しない場合は空文字 '' を設定する。
            row.append(s.scores.get((sub, config.KEY_GRADE), ''))
        data_rows.append(row)

    # --- 4. Excelへの書き込み ---
    sheet_name = config.SHEET_NAME_GRADE_SUMMARY
    # 準備したヘッダーとデータ行をリポジトリに渡し、実際のシート作成と書き込みを依頼する。
    repository.create_summary_sheet(sheet_name, headers, data_rows)

    # ▼▼▼ 修正箇所 ▼▼▼
    # ヘッダー行(1行目)を中央揃えにするよう、リポジトリに依頼する。
    repository.align_header_center(sheet_name, 1, 1)
    
    # 作成したシートの全セルに罫線を引くよう、リポジトリに依頼する。
    repository.apply_borders_to_all_cells(sheet_name)
    status_callback("-> 評定一覧シートの作成完了。")


# --- 出席状況一覧サービス ---

def create_attendance_summary(repository: ExcelRepository, students: List[Student], original_columns: pd.Index, status_callback):
    """
    「科目別個人出席状況一覧」シートを作成するサービス。
    元のCSVファイルの列順を維持し、2段ヘッダーを持つ一覧を作成する。
    """
    status_callback("\n処理中: 科目別個人出席状況シートを作成しています...")

    # --- 1. ヘッダー情報の準備 ---
    attendance_map = config.KEY_ATTENDANCE  # {'欠': '欠席', ...}
    short_types = list(attendance_map.keys())  # ['欠', '遅', '早']

    # 元の列情報から、ヘッダー2行目が出欠関連キー('欠', '遅', '早')である列を全て抽出する。
    attendance_keys = [col for col in original_columns if len(col) == 2 and col[1] in short_types]

    # --- 2. 入力値の検証 ---
    if not attendance_keys:
        status_callback("警告(出席状況): 出席関連データが見つかりません。スキップします。")
        return

    # --- 3. Excelに出力するデータ（2段ヘッダーとデータ行）の作成 ---
    # 1段目のヘッダー（科目名）を作成。['科目名', '', '科目A', '科目A', '科目B', ...]
    header_row1 = ['科目名', '']
    # 2段目のヘッダー（出欠種別）を作成。['学籍番号', '氏名', '欠席', '遅刻', '欠席', ...]
    header_row2 = ['学籍番号', '氏名']
    for subject, att_type_short in attendance_keys:
        header_row1.append(subject)
        header_row2.append(attendance_map.get(att_type_short, att_type_short))

    # データ行を作成
    data_rows = []
    for s in students:
        row = [s.id, s.name]
        for key in attendance_keys:
            # studentオブジェクトから (科目名, '欠') などをキーに出欠データを取得
            row.append(s.attendance.get(key, ''))
        data_rows.append(row)

    # --- 4. Excelへの書き込みと整形 ---
    sheet_name = config.SHEET_NAME_ATTENDANCE_SUMMARY
    # まず2段目のヘッダーとデータ行を書き込む
    repository.create_summary_sheet(sheet_name, header_row2, data_rows)

    # 1行目を挿入して、1段目のヘッダー（科目名）を書き込む
    ws = repository.workbook[sheet_name]
    ws.insert_rows(1)
    for col_idx, value in enumerate(header_row1, 1):
        ws.cell(row=1, column=col_idx, value=value)

    # --- 5. スタイル設定とセルの結合 ---
    # ヘッダー行にスタイル（背景色など）を適用
    repository.style_row(sheet_name, 1, is_header=True)
    repository.style_row(sheet_name, 2, is_header=True)

    # 1段目のヘッダーで、同じ科目が続く部分のセルを結合する
    if len(header_row1) > 2:
        start_col = 3 # 結合を開始する列
        # 3列目から最後までループ
        for col_idx in range(start_col + 1, len(header_row1) + 2):
            # 隣のセルと科目名が異なるか、または最終列まで達した場合
            if col_idx > len(header_row1) or header_row1[col_idx-1] != header_row1[start_col-1]:
                # 結合範囲が1セルより大きい場合
                if col_idx - 1 > start_col:
                    # start_col から col_idx - 1 までを結合するようリポジトリに依頼
                    repository.merge_header_cells(sheet_name, 1, start_col, col_idx - 1)
                start_col = col_idx # 新しい結合開始列を更新

    # ヘッダー全体を中央揃えにし、シート全体に罫線を引く
    repository.align_header_center(sheet_name, 1, 2)
    repository.apply_borders_to_all_cells(sheet_name)
    status_callback("-> 科目別個人出席状況シートの作成完了。")
















# # services/summary_service.py
# """
# 各種の集計シートを作成するビジネスロジックを提供するモジュール。
# task_runnerからの指示に基づき、データモデル(Student)と永続化層(ExcelRepository)を
# 組み合わせて、特定の目的のExcelシートを生成する。
# """

# from typing import List
# from collections import OrderedDict  # 順序を保持したまま重複を削除するために使用
# import pandas as pd
# from models.student import Student
# from repositories.excel_repository import ExcelRepository
# import config

# # --- 評定一覧サービス ---

# def create_grade_summary(repository: ExcelRepository, students: List[Student], original_columns: pd.Index, status_callback):
#     """
#     「評定一覧」シートを作成するサービス。
#     元のCSVファイルの科目順を維持して一覧を作成する。

#     【エラーハンドリングについて】
#     この関数内では、`try...except`によるエラー捕捉は行いません。
#     これは設計上の意図であり、リポジトリ層（例: Excelシート作成時）で発生したエラーは、
#     呼び出し元である`task_runner.py`まで伝播させ、そこで一元的にエラーハンドリングを行います。

#     Args:
#         repository (ExcelRepository): データ書き込みを担当するリポジトリ。
#         students (List[Student]): 学生データのリスト。
#         original_columns (pd.Index): 元のCSVの列情報。科目の順序を維持するために使用。
#         status_callback (function): UIへの状況通知コールバック。
#     """
#     status_callback("\n処理中: 評定一覧シートを作成しています...")

#     # --- 1. ヘッダーとなる科目リストの抽出 ---
#     # 元のCSVの列情報から、ヘッダーの2行目が「評」である列の科目名（1行目）を抽出する。
#     grade_cols_subjects = [col[0] for col in original_columns if len(col) == 2 and col[1] == config.KEY_GRADE]
#     # 抽出した科目リストから重複を削除しつつ、元の順序は維持する (OrderedDictの特性を利用)。
#     subjects = list(OrderedDict.fromkeys(grade_cols_subjects))

#     # --- 2. 入力値の検証 ---
#     # 評定データを持つ科目が一つも見つからない場合は、警告を出力して処理を中断する。
#     if not subjects:
#         status_callback("警告(評定一覧): 評定データを持つ科目が見つかりません。スキップします。")
#         return

#     # --- 3. Excelに出力するデータの作成 ---
#     # ヘッダー行を作成: ['学籍番号', '氏名', '科目A', '科目B', ...]
#     headers = ['学籍番号', '氏名'] + subjects

#     # データ行を作成: 各学生について、科目リストの順に評定データを格納していく。
#     data_rows = []
#     for s in students:
#         row = [s.id, s.name] # 各行の先頭は学籍番号と氏名
#         for sub in subjects:
#             # studentオブジェクトのscores辞書から、(科目名, "評") をキーにして評定を取得。
#             # 存在しない場合は空文字 '' を設定する。
#             row.append(s.scores.get((sub, config.KEY_GRADE), ''))
#         data_rows.append(row)

#     # --- 4. Excelへの書き込み ---
#     # 準備したヘッダーとデータ行をリポジトリに渡し、実際のシート作成と書き込みを依頼する。
#     repository.create_summary_sheet(config.SHEET_NAME_GRADE_SUMMARY, headers, data_rows)
#     # 作成したシートの全セルに罫線を引くよう、リポジトリに依頼する。
#     repository.apply_borders_to_all_cells(config.SHEET_NAME_GRADE_SUMMARY)
#     status_callback("-> 評定一覧シートの作成完了。")


# # --- 出席状況一覧サービス ---

# def create_attendance_summary(repository: ExcelRepository, students: List[Student], original_columns: pd.Index, status_callback):
#     """
#     「科目別個人出席状況一覧」シートを作成するサービス。
#     元のCSVファイルの列順を維持し、2段ヘッダーを持つ一覧を作成する。
#     """
#     status_callback("\n処理中: 科目別個人出席状況シートを作成しています...")

#     # --- 1. ヘッダー情報の準備 ---
#     attendance_map = config.KEY_ATTENDANCE  # {'欠': '欠席', ...}
#     short_types = list(attendance_map.keys())  # ['欠', '遅', '早']

#     # 元の列情報から、ヘッダー2行目が出欠関連キー('欠', '遅', '早')である列を全て抽出する。
#     attendance_keys = [col for col in original_columns if len(col) == 2 and col[1] in short_types]

#     # --- 2. 入力値の検証 ---
#     if not attendance_keys:
#         status_callback("警告(出席状況): 出席関連データが見つかりません。スキップします。")
#         return

#     # --- 3. Excelに出力するデータ（2段ヘッダーとデータ行）の作成 ---
#     # 1段目のヘッダー（科目名）を作成。['科目名', '', '科目A', '科目A', '科目B', ...]
#     header_row1 = ['科目名', '']
#     # 2段目のヘッダー（出欠種別）を作成。['学籍番号', '氏名', '欠席', '遅刻', '欠席', ...]
#     header_row2 = ['学籍番号', '氏名']
#     for subject, att_type_short in attendance_keys:
#         header_row1.append(subject)
#         header_row2.append(attendance_map.get(att_type_short, att_type_short))

#     # データ行を作成
#     data_rows = []
#     for s in students:
#         row = [s.id, s.name]
#         for key in attendance_keys:
#             # studentオブジェクトから (科目名, '欠') などをキーに出欠データを取得
#             row.append(s.attendance.get(key, ''))
#         data_rows.append(row)

#     # --- 4. Excelへの書き込みと整形 ---
#     sheet_name = config.SHEET_NAME_ATTENDANCE_SUMMARY
#     # まず2段目のヘッダーとデータ行を書き込む
#     repository.create_summary_sheet(sheet_name, header_row2, data_rows)

#     # 1行目を挿入して、1段目のヘッダー（科目名）を書き込む
#     ws = repository.workbook[sheet_name]
#     ws.insert_rows(1)
#     for col_idx, value in enumerate(header_row1, 1):
#         ws.cell(row=1, column=col_idx, value=value)

#     # --- 5. スタイル設定とセルの結合 ---
#     # ヘッダー行にスタイル（背景色など）を適用
#     repository.style_row(sheet_name, 1, is_header=True)
#     repository.style_row(sheet_name, 2, is_header=True)

#     # 1段目のヘッダーで、同じ科目が続く部分のセルを結合する
#     if len(header_row1) > 2:
#         start_col = 3 # 結合を開始する列
#         # 3列目から最後までループ
#         for col_idx in range(start_col + 1, len(header_row1) + 2):
#             # 隣のセルと科目名が異なるか、または最終列まで達した場合
#             if col_idx > len(header_row1) or header_row1[col_idx-1] != header_row1[start_col-1]:
#                 # 結合範囲が1セルより大きい場合
#                 if col_idx - 1 > start_col:
#                     # start_col から col_idx - 1 までを結合するようリポジトリに依頼
#                     repository.merge_header_cells(sheet_name, 1, start_col, col_idx - 1)
#                 start_col = col_idx # 新しい結合開始列を更新

#     # ヘッダー全体を中央揃えにし、シート全体に罫線を引く
#     repository.align_header_center(sheet_name, 1, 2)
#     repository.apply_borders_to_all_cells(sheet_name)
#     status_callback("-> 科目別個人出席状況シートの作成完了。")





































# # from typing import List
# # from collections import OrderedDict
# # import pandas as pd
# # from models.student import Student
# # from repositories.excel_repository import ExcelRepository
# # import config

# # def create_grade_summary(repository: ExcelRepository, students: List[Student], original_columns: pd.Index, status_callback):
# #     """
# #     【課題2】評定一覧を作成するサービス。
# #     元のCSVファイルの科目順を維持して一覧を作成する。
# #     """
# #     status_callback("\n処理中: 評定一覧シートを作成しています...")

# #     # 元の列順序(original_columns)から、ヘッダーの第2レベルが「評」である科目を抽出
# #     grade_cols_subjects = [col[0] for col in original_columns if len(col) == 2 and col[1] == config.KEY_GRADE]
# #     # 重複する科目を削除しつつ、順序は維持する (OrderedDictの特性を利用)
# #     subjects = list(OrderedDict.fromkeys(grade_cols_subjects))

# #     if not subjects:
# #         status_callback("警告(評定一覧): 評定データを持つ科目が見つかりません。スキップします。")
# #         return

# #     # Excelシートのヘッダー行を作成
# #     headers = ['学籍番号', '氏名'] + subjects

# #     # Excelシートのデータ部分を作成
# #     data_rows = []
# #     for s in students:
# #         row = [s.id, s.name]
# #         for sub in subjects:
# #             # student.scoresから (科目名, "評") をキーにして値を取得。存在しない場合は空文字。
# #             row.append(s.scores.get((sub, config.KEY_GRADE), ''))
# #         data_rows.append(row)

# #     # Repositoryにシート作成を依頼
# #     repository.create_summary_sheet(config.SHEET_NAME_GRADE_SUMMARY, headers, data_rows)
# #     # 作成したシートに罫線を適用
# #     repository.apply_borders_to_all_cells(config.SHEET_NAME_GRADE_SUMMARY)
# #     status_callback("-> 評定一覧シートの作成完了。")


# # def create_attendance_summary(repository: ExcelRepository, students: List[Student], original_columns: pd.Index, status_callback):
# #     """
# #     【課題3】科目別個人出席状況一覧を作成するサービス。
# #     元のCSVファイルの列順を維持し、2段ヘッダーを持つ一覧を作成する。
# #     """
# #     status_callback("\n処理中: 科目別個人出席状況シートを作成しています...")

# #     attendance_map = config.KEY_ATTENDANCE # {'欠': '欠席', ...}
# #     short_types = list(attendance_map.keys()) # ['欠', '遅', '早']

# #     # 元の列順序から、ヘッダーの第2レベルが出欠関連のキーであるものを抽出
# #     attendance_keys = [col for col in original_columns if len(col) == 2 and col[1] in short_types]

# #     if not attendance_keys:
# #         status_callback("警告(出席状況): 出席関連データが見つかりません。スキップします。")
# #         return

# #     # --- 2段ヘッダーの生成 ---
# #     header_row1 = ['科目名', '']
# #     header_row2 = ['学籍番号', '氏名']
# #     for subject, att_type_short in attendance_keys:
# #         header_row1.append(subject)
# #         header_row2.append(attendance_map.get(att_type_short, att_type_short))

# #     # --- データ行の生成 ---
# #     data_rows = []
# #     for s in students:
# #         row = [s.id, s.name]
# #         for key in attendance_keys:
# #             row.append(s.attendance.get(key, ''))
# #         data_rows.append(row)

# #     # --- Excelへの書き込みと整形 ---
# #     sheet_name = config.SHEET_NAME_ATTENDANCE_SUMMARY
# #     repository.create_summary_sheet(sheet_name, header_row2, data_rows)

# #     ws = repository.workbook[sheet_name]
# #     ws.insert_rows(1)
# #     for col_idx, value in enumerate(header_row1, 1):
# #         ws.cell(row=1, column=col_idx, value=value)

# #     # --- スタイルとセル結合 ---
# #     repository.style_row(sheet_name, 1, is_header=True)
# #     repository.style_row(sheet_name, 2, is_header=True)

# #     if len(header_row1) > 2:
# #         start_col = 3
# #         for col_idx in range(start_col + 1, len(header_row1) + 2):
# #             if col_idx > len(header_row1) or header_row1[col_idx-1] != header_row1[start_col-1]:
# #                 if col_idx - 1 > start_col:
# #                     repository.merge_header_cells(sheet_name, 1, start_col, col_idx - 1)
# #                 start_col = col_idx

# #     # ヘッダー全体を中央揃えにする
# #     repository.align_header_center(sheet_name, 1, 2)
# #     # 作成したシートに罫線を適用
# #     repository.apply_borders_to_all_cells(sheet_name)
# #     status_callback("-> 科目別個人出席状況シートの作成完了。")
























