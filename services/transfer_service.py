

# services/transfer_service.py

from typing import List
from models.student import Student
from repositories.excel_repository import ExcelRepository

def run(repository: ExcelRepository, students: List[Student], terms: List[str], status_callback):
    """
    「成績一覧を更新」処理を実行するサービス。

    この関数は、UIからの要求を受けて、指定された学期（前期、通期）の成績一覧シートを
    更新するための処理を統括します。実際のExcelファイルへの書き込みは、
    引数で受け取ったリポジトリ(repository)に任せます。

    【エラーハンドリングについて】
    この関数内では、`try...except`によるエラー捕捉は行いません。
    これは設計上の意図であり、リポジトリ層（例: Excelファイル書き込み時）で発生したエラーは、
    呼び出し元である`task_runner.py`まで伝播させ、そこで一元的にエラーハンドリングを行います。
    これにより、エラー処理のロジックが分散せず、管理しやすくなります。

    Args:
        repository (ExcelRepository): データ書き込みを担当するリポジトリ。
        students (List[Student]): 転記する学生データのリスト。
        terms (List[str]): 対象となる学期（例: ["前期", "通期"]）。
        status_callback (function): UIのステータス表示を更新するためのコールバック関数。
    """
    # --- 1. 処理開始の通知 ---
    # UIに対し、これから何の処理を始めるかを通知します。
    status_callback("\n処理中: 成績一覧シートを更新しています...")
    
    # --- 2. 入力値の検証 ---
    # 処理対象となる学期が一つも選択されていない場合は、警告メッセージをUIに表示し、
    # それ以降の処理を行わずにここで終了します。
    # これも一種の事前エラーハンドリングです。
    if not terms:
        status_callback("警告: 成績転記の対象学期が選択されていません。")
        return

    # --- 3. メイン処理の実行 ---
    # 選択された各学期（"前期", "通期"など）についてループ処理を行います。
    for term in terms:
        # Excelへの具体的な書き込み処理は、リポジトリのメソッドに完全に任せます（処理の委譲）。
        # これにより、このサービスは「何をするか」だけを管理し、「どうやってやるか」は
        # リポジトリが担当するという役割分担が明確になります。
        repository.update_grades_sheet(term, students, status_callback)

