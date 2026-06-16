from datetime import date

from core.file_namer import build_output_filename


def test_build_output_filename() -> None:
    assert build_output_filename(date(2026, 3, 9)) == "【開示用未収リストALL(売却予定リスト除外済)】20260309.xlsx"
