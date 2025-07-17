




# models/student.py
"""
アプリケーションのデータモデル（Studentクラス）と、
CSVからのデータ読み込みロジックを定義するモジュール。
"""

import pandas as pd  # CSVファイルの読み込みとデータ操作のために使用
import config        # アプリケーション共通の設定値（キーワードなど）をインポート
from typing import List, Dict, Tuple, Any # 型ヒント（コードの可読性向上）のために使用

class Student:
    """
    学生一人分のデータを保持するクラス（データモデル）。
    CSVから読み込んだ情報を、プログラム内で扱いやすいように整理して格納する。
    """
    def __init__(self, student_id: str, name: str):
        """Studentオブジェクトの初期化"""
        self.id: str = student_id  # 学籍番号
        self.name: str = name      # 氏名
        # 成績データ（例: {('数学', '本試'): 80, ('英語', '再試'): 65}）
        self.scores: Dict[Tuple[str, str], Any] = {}
        # 出欠データ（例: {('数学', '欠'): 2, ('物理', '遅'): 1}）
        self.attendance: Dict[Tuple[str, str], Any] = {}
        # その他の集計データ（例: {'備考': '特記事項あり', '総点': 550}）
        self.summary_data: Dict[str, Any] = {}

    def add_score(self, subject: str, test_type: str, score: Any):
        """成績データを追加する"""
        self.scores[(subject, test_type)] = score

    def add_attendance(self, subject: str, att_type: str, value: Any):
        """出欠データを追加する"""
        self.attendance[(subject, att_type)] = value

    def __repr__(self) -> str:
        """
        このオブジェクトをprint()したときなどに表示される、開発者向けの文字列。
        デバッグ時に中身を確認しやすくする目的で定義する。
        """
        return f"Student(id={self.id}, name='{self.name}')"


def load_students_from_csv(csv_path: str) -> Tuple[List[Student], pd.Index]:
    """
    CSVファイルを読み込み、Studentオブジェクトのリストと元の列順序を生成して返す。

    【エラーハンドリングについて】
    この関数は、外部ファイルであるCSVを扱うため、堅牢なエラー処理が実装されている。
    ファイル形式、文字コード、必須列の欠落など、様々な問題を検知し、
    ユーザーが原因を理解しやすい具体的なエラーメッセージを生成する。

    Args:
        csv_path (str): 読み込むCSVファイルのパス。

    Returns:
        Tuple[List[Student], pd.Index]:
            - Studentオブジェクトのリスト。
            - 元のCSVの列情報。後の集計処理で列の順序を維持するために使用。
    """
    # --- 1. CSVファイルの読み込み ---
    try:
        # pandasを使いCSVを読み込む。このアプリケーションで扱うCSVは特殊なフォーマットを持つため、
        # いくつかのオプションを指定している。
        df = pd.read_csv(
            csv_path,
            encoding="cp932",      # 文字コードをShift-JISに指定
            skiprows=[1, 2],       # 読み飛ばす行を指定 (ExcelでCSVを開いた際に追加されがちな空白行など)
            header=[0, 1]          # ヘッダーが2行にまたがっていることを指定
        )
    except UnicodeDecodeError:
        # 【エラー処理】文字コードがShift-JIS(cp932)以外で保存されている場合
        raise ValueError("CSVファイルの文字コードエラーです。\nファイルがShift-JIS (cp932) 形式で保存されているか確認してください。")
    except FileNotFoundError:
        # 【エラー処理】ファイルが存在しない場合。より上位のtask_runner.pyで処理するため、ここではそのままエラーを送出。
        raise
    except Exception as e:
        # 【エラー処理】上記以外のpandas関連エラー（例: CSVのフォーマットが壊れている）。
        raise ValueError(f"CSVファイルの読み込みに失敗しました。\nファイル形式が正しいか、破損していないか確認してください。\n詳細: {e}")

    # --- 2. CSV構造の検証 ---
    # 【エラー処理】読み込んだCSVのヘッダーが期待通りの2段ヘッダーになっているか検証する。
    if not isinstance(df.columns, pd.MultiIndex) or df.columns.nlevels < 2:
        raise KeyError("CSVファイルのヘッダー形式が不正です。1行目と2行目で構成される2段ヘッダーが必要です。")

    # 【エラー処理】列が一つも存在しない空のCSVでないか検証する。
    if len(df.columns) < 1:
        raise KeyError("CSVファイルに列が存在しません。")
    
    # --- 3. 必須列の特定 ---
    # 学籍番号の列は、規約として常に最初の列とする。
    student_id_col_tuple = df.columns[0]
    # '氏名' 列を探す。ヘッダー2行目に '氏名' という文字列を持つ列を特定する。
    student_name_col_tuple = next((c for c in df.columns if str(c[1]).strip() == config.KEY_STUDENT_NAME), None)
    
    # 【エラー処理】'氏名' 列が見つからなかった場合。
    if student_name_col_tuple is None:
        raise KeyError(f"必須列 '{config.KEY_STUDENT_NAME}' がCSVヘッダー(2行目)に見つかりません。")

    # --- 4. データ行の解析とStudentオブジェクトの生成 ---
    students: List[Student] = []
    # DataFrameの各行をループ処理し、Studentオブジェクトに変換していく。
    for index, row in df.iterrows():
        try:
            # 学籍番号が空の行（データのない空白行など）は処理対象外としてスキップする。
            student_id = str(row[student_id_col_tuple])
            if pd.isna(student_id) or not student_id.strip():
                continue

            # 氏名が空でも、学籍番号があればデータとして処理を続行する。
            student_name = row[student_name_col_tuple]
            if pd.isna(student_name):
                student_name = ""

            # Studentオブジェクトを生成。
            student = Student(student_id, student_name)

            # --- 5. 各科目のデータをStudentオブジェクトに格納 ---
            # 行内の各列をループ処理する。
            for col_lv1, col_lv2 in df.columns:
                # 学籍番号と氏名の列は既に処理済みなのでスキップ。
                if (col_lv1, col_lv2) in [student_id_col_tuple, student_name_col_tuple]:
                    continue

                # セルの値を取得。欠損値(NaN)の場合は空文字に変換する。
                value = row[(col_lv1, col_lv2)]
                if pd.isna(value): value = ''
                
                # 列のヘッダー（2行目）を見て、データの種類を判別し、適切なメソッドで格納する。
                if col_lv2 in config.KEY_TEST_TYPES:   # '本試', '再試', '評' のいずれか
                    student.add_score(col_lv1, col_lv2, value)
                elif col_lv2 in config.KEY_ATTENDANCE: # '欠', '遅', '早' のいずれか
                    student.add_attendance(col_lv1, col_lv2, value)
                elif col_lv1 in config.KEY_OTHER_COLS: # '備考', '総点' など
                    student.summary_data[col_lv1] = value

            students.append(student) # 完成したStudentオブジェクトをリストに追加
            
        except Exception as e:
            # 【エラー処理】特定のデータ行の処理中に予期せぬエラーが発生した場合。
            # エラーが発生した行番号をメッセージに含めることで、原因特定を容易にする。
            # (indexは0から始まるので、Excel上の行番号に合わせるため +4 する)
            raise ValueError(f"CSVファイルの {index + 4} 行目のデータ処理中にエラーが発生しました。\nデータ形式を確認してください。\n詳細: {e}")

    return students, df.columns































# # models/student.py (エラー処理追加版)

# import pandas as pd
# import config
# from typing import List, Dict, Tuple, Any

# class Student:
#     """学生一人分のデータを保持するクラス"""
#     def __init__(self, student_id: str, name: str):
#         self.id: str = student_id
#         self.name: str = name
#         self.scores: Dict[Tuple[str, str], Any] = {}
#         self.attendance: Dict[Tuple[str, str], Any] = {}
#         self.summary_data: Dict[str, Any] = {}

#     def add_score(self, subject: str, test_type: str, score: Any):
#         self.scores[(subject, test_type)] = score

#     def add_attendance(self, subject: str, att_type: str, value: Any):
#         self.attendance[(subject, att_type)] = value

#     def __repr__(self) -> str:
#         return f"Student(id={self.id}, name='{self.name}')"


# def load_students_from_csv(csv_path: str) -> Tuple[List[Student], pd.Index]:
#     """
#     【エラー処理追加版】CSVファイルを読み込み、Studentオブジェクトのリストと元の列順序を生成して返す。
#     CSVのフォーマットや必須列のチェックを行い、問題があれば例外を送出する。
#     """
#     try:
#         # ヘッダーが2行にまたがっていることを指定して読み込む
#         df = pd.read_csv(csv_path, encoding="cp932", skiprows=[1, 2], header=[0, 1])
#     except UnicodeDecodeError:
#         # 文字コードが原因で読み込めない場合
#         raise ValueError("CSVファイルの文字コードエラーです。\nファイルがShift-JIS (cp932) 形式で保存されているか確認してください。")
#     except FileNotFoundError:
#         # ファイルが見つからない場合は、呼び出し元で処理するため、そのまま送出
#         raise
#     except Exception as e:
#         # その他のpandas関連エラー（フォーマット不正など）
#         raise ValueError(f"CSVファイルの読み込みに失敗しました。\nファイル形式が正しいか、破損していないか確認してください。\n詳細: {e}")

#     # --- 必須列の存在チェック ---
#     if not isinstance(df.columns, pd.MultiIndex) or df.columns.nlevels < 2:
#         raise KeyError("CSVファイルのヘッダー形式が不正です。1行目と2行目で構成される2段ヘッダーが必要です。")

#     # 学籍番号の列を取得 (通常は最初の列)
#     if len(df.columns) < 1:
#         raise KeyError("CSVファイルに列が存在しません。")
#     student_id_col_tuple = df.columns[0]

#     # 氏名列の存在をチェック
#     student_name_col_tuple = next((c for c in df.columns if str(c[1]).strip() == config.KEY_STUDENT_NAME), None)
#     if student_name_col_tuple is None:
#         raise KeyError(f"必須列 '{config.KEY_STUDENT_NAME}' がCSVヘッダー(2行目)に見つかりません。")

#     students: List[Student] = []
#     # DataFrameの各行をStudentオブジェクトに変換
#     for index, row in df.iterrows():
#         try:
#             # 欠損値(NaN)など、不正な値でないかチェック
#             student_id = str(row[student_id_col_tuple])
#             if pd.isna(student_id) or not student_id.strip():
#                 # 学籍番号が空の行はデータとして扱わずスキップする
#                 continue

#             student_name = row[student_name_col_tuple]
#             if pd.isna(student_name):
#                 student_name = "" # 名前が空でも処理は続行

#             student = Student(student_id, student_name)

#             for col_lv1, col_lv2 in df.columns:
#                 if (col_lv1, col_lv2) == student_id_col_tuple or (col_lv1, col_lv2) == student_name_col_tuple:
#                     continue

#                 # 値を取得し、NaNの場合は空文字に変換
#                 value = row[(col_lv1, col_lv2)]
#                 if pd.isna(value):
#                     value = ''

#                 if col_lv2 in config.KEY_TEST_TYPES:
#                     student.add_score(col_lv1, col_lv2, value)
#                 elif col_lv2 in config.KEY_ATTENDANCE:
#                     student.add_attendance(col_lv1, col_lv2, value)
#                 elif col_lv1 in config.KEY_OTHER_COLS:
#                     student.summary_data[col_lv1] = value

#             students.append(student)
#         except Exception as e:
#             # 各行の処理中に予期せぬエラーが発生した場合、行番号を添えて例外を送出
#             raise ValueError(f"CSVファイルの {index + 4} 行目のデータ処理中にエラーが発生しました。\nデータ形式を確認してください。\n詳細: {e}")

#     return students, df.columns




































# # # models/student.py (単純化・エラー処理削除版)

# # import pandas as pd
# # import config
# # from typing import List, Dict, Tuple, Any

# # class Student:
# #     """学生一人分のデータを保持するクラス"""
# #     def __init__(self, student_id: str, name: str):
# #         self.id: str = student_id
# #         self.name: str = name
# #         self.scores: Dict[Tuple[str, str], Any] = {}
# #         self.attendance: Dict[Tuple[str, str], Any] = {}
# #         self.summary_data: Dict[str, Any] = {}

# #     def add_score(self, subject: str, test_type: str, score: Any):
# #         self.scores[(subject, test_type)] = score

# #     def add_attendance(self, subject: str, att_type: str, value: Any):
# #         self.attendance[(subject, att_type)] = value

# #     def __repr__(self) -> str:
# #         return f"Student(id={self.id}, name='{self.name}')"


# # def load_students_from_csv(csv_path: str) -> Tuple[List[Student], pd.Index]:
# #     """
# #     【単純化版】CSVファイルを読み込み、Studentオブジェクトのリストと元の列順序を生成して返す。
# #     """
# #     df = pd.read_csv(csv_path, encoding="cp932", skiprows=[1, 2], header=[0, 1])
# #     students: List[Student] = []
    
# #     # 必須の列情報を取得
# #     student_id_col_tuple = df.columns[0]
# #     student_name_col_tuple = next((c for c in df.columns if c[1].strip() == config.KEY_STUDENT_NAME), None)

# #     # 必須列(氏名)の存在チェックを削除。
# #     # 存在しない場合、この後の処理でエラーとなり停止する。

# #     # DataFrameの各行をStudentオブジェクトに変換
# #     for _, row in df.iterrows():
# #         student_id = str(row[student_id_col_tuple])
# #         student_name = row[student_name_col_tuple]
        
# #         student = Student(student_id, student_name)

# #         for col_lv1, col_lv2 in df.columns:
# #             if (col_lv1, col_lv2) == student_id_col_tuple or (col_lv1, col_lv2) == student_name_col_tuple:
# #                 continue
            
# #             value = row[(col_lv1, col_lv2)]
            
# #             if col_lv2 in config.KEY_TEST_TYPES:
# #                 student.add_score(col_lv1, col_lv2, value)
# #             elif col_lv2 in config.KEY_ATTENDANCE:
# #                 student.add_attendance(col_lv1, col_lv2, value)
# #             elif col_lv1 in config.KEY_OTHER_COLS:
# #                 student.summary_data[col_lv1] = value
        
# #         students.append(student)
        
# #     return students, df.columns























