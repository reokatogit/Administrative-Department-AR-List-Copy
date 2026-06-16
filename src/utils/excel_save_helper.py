from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client


XL_OPEN_XML_WORKBOOK = 51
XL_REPAIR_FILE = 1


def normalize_with_excel(input_path: str | Path, output_path: str | Path | None = None) -> None:
    """
    openpyxl で一度保存した xlsx を Excel 本体で開き直して保存し直す。
    修復ダイアログが出る系のメタ情報を Excel 側で正規化する目的。

    input_path:
        openpyxl が保存した一時ファイル
    output_path:
        最終出力先。省略時は input_path を上書き保存
    """
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve() if output_path else input_path

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        workbook = excel.Workbooks.Open(
            str(input_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Local=True,
            CorruptLoad=XL_REPAIR_FILE,
        )

        if output_path == input_path:
            workbook.Save()
        else:
            if output_path.exists():
                output_path.unlink()

            workbook.SaveAs(
                str(output_path),
                FileFormat=XL_OPEN_XML_WORKBOOK,
                AddToMru=False,
                Local=True,
            )

        workbook.Close(SaveChanges=False)
        workbook = None

    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()