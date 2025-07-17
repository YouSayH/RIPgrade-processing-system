# --- シート名 ---
SHEET_NAME_GRADES_TEMPLATE = "成績一覧 ({term})"
SHEET_NAME_GRADE_SUMMARY = "評定一覧"
SHEET_NAME_ATTENDANCE_SUMMARY = "科目別個人出席状況"

# --- CSV/Excel内のキーワード ---
# データモデルや各サービスで参照される
KEY_STUDENT_NAME = "氏名"
KEY_TEST_TYPES = ("本試", "再試", "評")
KEY_GRADE = "評"
KEY_ATTENDANCE = {'欠': '欠席', '遅': '遅刻', '早': '早退'}

# KEY_OTHER_COLSは主に成績一覧シートで扱われる汎用的な列
KEY_OTHER_COLS = ["備考", "科目数", "総点", "平均点", "欠課合計", "順位", "再試数"]


# --- Excelスタイル設定 ---
# リポジトリ層で参照される
HEADER_FONT_COLOR = "000000"
HEADER_FILL_COLOR = "CEE6C1"