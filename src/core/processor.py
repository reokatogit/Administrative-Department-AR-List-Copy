from __future__ import annotations

import re
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

from core.constants import DATA_START_ROW, HEADER_ROW
from core.mapping import ConditionalCopyRule, CopyRule, DEFAULT_MAPPING, MappingConfig
from core.validator import validate_inputs
from utils.excel_helpers import (
    build_row_index_by_key,
    col_to_index,
    copy_row_style_and_formulas,
    find_last_data_row,
    get_cell_value,
    is_blank,
    iter_non_blank_rows,
    next_sequential_number,
    normalize_key,
)


MONTHLY_KEY_COL = "A"
MASTER_KEY_COL = "E"


class ProcessingError(Exception):
    pass


@dataclass(frozen=True)
class ProcessingResult:
    output_path: Path
    source_row_count: int
    updated_row_count: int
    appended_row_count: int


class WorkbookProcessor:
    def __init__(self, mapping: MappingConfig | None = None) -> None:
        self.mapping = mapping or DEFAULT_MAPPING

    def process(
        self,
        monthly_path: str | Path,
        master_path: str | Path,
        output_path: str | Path,
        extract_date: str,
        progress_callback: Callable[[int, str], None] | None = None,
    ) -> ProcessingResult:
        if not self.mapping.is_configured:
            raise ProcessingError("マッピング定義が未設定です。src/core/mapping.py を埋めてください。")

        output_path = Path(output_path)
        validated = validate_inputs(monthly_path, master_path, output_path.parent, self.mapping)

        source_ws = validated.monthly_workbook.worksheets[0]
        target_ws = validated.master_workbook.worksheets[0]

        source_rows = list(iter_non_blank_rows(source_ws, DATA_START_ROW))
        if not source_rows:
            raise ProcessingError("月次ファイルに追記対象のデータ行がありません。")

        total_rows = len(source_rows)
        self._report_progress(progress_callback, 0, "処理を開始しています…")

        key_to_row = build_row_index_by_key(target_ws, MASTER_KEY_COL, DATA_START_ROW)
        append_row = find_last_data_row(target_ws, DATA_START_ROW) + 1

        updated_count = 0
        appended_count = 0
        progress_step = max(1, total_rows // 100)

        for i, source_row in enumerate(source_rows, start=1):
            source_key = normalize_key(get_cell_value(source_ws, source_row, MONTHLY_KEY_COL))

            if source_key and source_key in key_to_row:
                target_row = key_to_row[source_key]
                self._update_existing_row(source_ws, target_ws, source_row, target_row)
                updated_count += 1
            else:
                target_row = append_row
                self._append_new_row(source_ws, target_ws, source_row, target_row)
                appended_count += 1
                append_row += 1

                appended_key = normalize_key(get_cell_value(target_ws, target_row, MASTER_KEY_COL))
                if appended_key:
                    key_to_row[appended_key] = target_row

            if i % progress_step == 0 or i == total_rows:
                percent = int((i / total_rows) * 88)
                self._report_progress(
                    progress_callback,
                    percent,
                    f"処理中… {i}/{total_rows}（更新 {updated_count} 件 / 追加 {appended_count} 件）",
                )

        self._report_progress(progress_callback, 90, "I列の抽出日を更新しています…")
        self._update_i_column_extract_date(target_ws, extract_date)

        self._report_progress(progress_callback, 93, "並び替えをしています…")
        self._sort_all_sheet(target_ws)

        self._report_progress(progress_callback, 96, "フィルタを更新しています…")
        self._refresh_auto_filter(target_ws)

        self._enable_recalc_on_open(validated.master_workbook)

        self._report_progress(progress_callback, 97, "数式参照を調整しています…")
        self._sanitize_formula_references(validated.master_workbook)

        self._report_progress(progress_callback, 98, "保存しています…")
        self._save_output_workbook(validated.master_workbook, output_path)

        self._report_progress(progress_callback, 100, "完了しました。")

        return ProcessingResult(
            output_path=output_path,
            source_row_count=total_rows,
            updated_row_count=updated_count,
            appended_row_count=appended_count,
        )

    def _update_existing_row(self, source_ws, target_ws, source_row: int, target_row: int) -> None:
        """
        既存行更新:
        - 転記対象列を更新
        - 派生列を更新
        - 固定値を反映
        - 採番、書式コピー、数式コピーはしない
        """
        self._apply_copy_rules(source_ws, target_ws, source_row, target_row)
        self._apply_conditional_copy_rules(source_ws, target_ws, source_row, target_row)
        self._apply_value_map_copy_rules(source_ws, target_ws, source_row, target_row)
        self._apply_fixed_value_rules(target_ws, target_row)

    def _append_new_row(self, source_ws, target_ws, source_row: int, target_row: int) -> None:
        """
        新規追加:
        - 前行から書式/数式コピー
        - 転記
        - 派生列
        - 採番
        - 固定値
        """
        template_row = target_row - 1 if target_row > DATA_START_ROW else DATA_START_ROW
        copy_row_style_and_formulas(target_ws, template_row, target_row)

        self._apply_copy_rules(source_ws, target_ws, source_row, target_row)
        self._apply_conditional_copy_rules(source_ws, target_ws, source_row, target_row)
        self._apply_value_map_copy_rules(source_ws, target_ws, source_row, target_row)
        self._apply_auto_number_rules(target_ws, target_row)
        self._apply_fixed_value_rules(target_ws, target_row)

    def _apply_copy_rules(self, source_ws, target_ws, source_row: int, target_row: int) -> None:
        for rule in self.mapping.always_copy_rules:
            self._copy_value(rule, source_ws, target_ws, source_row, target_row)

    def _apply_conditional_copy_rules(self, source_ws, target_ws, source_row: int, target_row: int) -> None:
        for rule in self.mapping.conditional_copy_rules:
            condition_value = source_ws.cell(
                row=source_row,
                column=col_to_index(rule.condition_source_col),
            ).value
            if condition_value in rule.allowed_values:
                self._copy_value(rule, source_ws, target_ws, source_row, target_row)

    def _apply_value_map_copy_rules(self, source_ws, target_ws, source_row: int, target_row: int) -> None:
        for rule in self.mapping.value_map_copy_rules:
            source_value = source_ws.cell(
                row=source_row,
                column=col_to_index(rule.source_col),
            ).value
            mapped_value = rule.value_map.get(source_value, rule.default_value)
            target_ws.cell(
                row=target_row,
                column=col_to_index(rule.target_col),
            ).value = mapped_value

    def _apply_fixed_value_rules(self, target_ws, target_row: int) -> None:
        for rule in self.mapping.fixed_value_rules:
            target_cell = target_ws.cell(
                row=target_row,
                column=col_to_index(rule.target_col),
            )
            if rule.only_if_blank and not is_blank(target_cell.value):
                continue
            target_cell.value = rule.value

    def _apply_auto_number_rules(self, target_ws, target_row: int) -> None:
        for rule in self.mapping.auto_number_rules:
            target_ws.cell(
                row=target_row,
                column=col_to_index(rule.target_col),
            ).value = next_sequential_number(
                ws=target_ws,
                target_col=rule.target_col,
                current_append_row=target_row,
                start_number=rule.start_number,
            )

    @staticmethod
    def _copy_value(
        rule: CopyRule | ConditionalCopyRule,
        source_ws,
        target_ws,
        source_row: int,
        target_row: int,
    ) -> None:
        source_value = source_ws.cell(
            row=source_row,
            column=col_to_index(rule.source_col),
        ).value
        target_ws.cell(
            row=target_row,
            column=col_to_index(rule.target_col),
        ).value = source_value

    @staticmethod
    def _update_i_column_extract_date(target_ws, extract_date: str) -> None:
        """
        I列の数式内にある >=yyyy/m/d または >=yyyy/mm/dd の日付条件を、
        画面入力の抽出日に置換する。
        """
        pattern = re.compile(r'">=\d{4}/\d{1,2}/\d{1,2}"')
        replacement = f'">={extract_date}"'

        last_row = find_last_data_row(target_ws, DATA_START_ROW)

        for row in range(DATA_START_ROW, last_row + 1):
            cell = target_ws.cell(row=row, column=col_to_index("I"))
            value = cell.value

            if not isinstance(value, str):
                continue
            if not value.startswith("="):
                continue
            if ">=" not in value:
                continue

            cell.value = pattern.sub(replacement, value)

    @staticmethod
    def _sort_all_sheet(target_ws) -> None:
        """
        ALLシートを以下の順で並び替える。
        第一基準: BU列（対応方針コード）昇順
        第二基準: G列（未収金額）降順

        行の値・書式・数式をまとめて並び替える。
        数式は移動先行に合わせて可能な範囲で翻訳する。
        """
        start_row = DATA_START_ROW
        last_row = find_last_data_row(target_ws, DATA_START_ROW)
        max_col = target_ws.max_column

        if last_row <= start_row:
            return

        def parse_amount(value) -> float:
            if value is None:
                return 0.0

            if isinstance(value, (int, float)):
                return float(value)

            if isinstance(value, str):
                text = value.strip()
                if not text:
                    return 0.0

                # 金額文字列（カンマ、円記号、全角カンマ）を並び替え用数値へ正規化
                text = text.replace(",", "").replace("，", "").replace("円", "")
                text = text.replace(" ", "")

                if text.startswith("(") and text.endswith(")"):
                    text = f"-{text[1:-1]}"

                try:
                    return float(text)
                except ValueError:
                    return 0.0

            return 0.0

        def sort_key(item):
            row_num, _row_cells = item

            bu_value = target_ws.cell(row=row_num, column=col_to_index("BU")).value
            g_value = target_ws.cell(row=row_num, column=col_to_index("G")).value

            try:
                bu_key = int(bu_value)
            except (TypeError, ValueError):
                bu_key = 9999

            g_key = parse_amount(g_value)

            return (bu_key, -g_key)

        rows_snapshot = []

        for row in range(start_row, last_row + 1):
            row_cells = []
            for col in range(1, max_col + 1):
                cell = target_ws.cell(row=row, column=col)
                row_cells.append(
                    {
                        "value": cell.value,
                        "style": copy(cell._style) if cell.has_style else None,
                        "number_format": cell.number_format,
                        "font": copy(cell.font),
                        "fill": copy(cell.fill),
                        "border": copy(cell.border),
                        "alignment": copy(cell.alignment),
                        "protection": copy(cell.protection),
                        "origin": cell.coordinate,
                    }
                )
            rows_snapshot.append((row, row_cells))

        rows_snapshot.sort(key=sort_key)

        for new_row, (old_row, row_cells) in enumerate(rows_snapshot, start=start_row):
            source_height = target_ws.row_dimensions[old_row].height
            if source_height is not None:
                target_ws.row_dimensions[new_row].height = source_height

            for col, cell_data in enumerate(row_cells, start=1):
                target_cell = target_ws.cell(row=new_row, column=col)

                value = cell_data["value"]

                if isinstance(value, str) and value.startswith("="):
                    try:
                        value = Translator(
                            value,
                            origin=cell_data["origin"],
                        ).translate_formula(target_cell.coordinate)
                    except Exception:
                        pass

                target_cell.value = value

                if cell_data["style"] is not None:
                    target_cell._style = copy(cell_data["style"])

                target_cell.number_format = cell_data["number_format"]
                target_cell.font = copy(cell_data["font"])
                target_cell.fill = copy(cell_data["fill"])
                target_cell.border = copy(cell_data["border"])
                target_cell.alignment = copy(cell_data["alignment"])
                target_cell.protection = copy(cell_data["protection"])

    @staticmethod
    def _refresh_auto_filter(target_ws) -> None:
        if not target_ws.auto_filter:
            return

        last_row = find_last_data_row(target_ws, DATA_START_ROW)
        last_col = get_column_letter(target_ws.max_column)
        target_ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col}{last_row}"

    @staticmethod
    def _enable_recalc_on_open(workbook) -> None:
        calculation = getattr(workbook, "calculation", None)
        if calculation is None:
            return
        if hasattr(calculation, "fullCalcOnLoad"):
            calculation.fullCalcOnLoad = True
        if hasattr(calculation, "forceFullCalc"):
            calculation.forceFullCalc = True

    @staticmethod
    def _save_output_workbook(workbook, output_path: Path) -> None:
        try:
            workbook.save(output_path)
        except Exception as exc:  # noqa: BLE001
            raise ProcessingError(f"出力ファイルの保存に失敗しました。\n{output_path}\n{exc}") from exc

    @staticmethod
    def _report_progress(
        progress_callback: Callable[[int, str], None] | None,
        percent: int,
        message: str,
    ) -> None:
        if progress_callback is None:
            return

        try:
            progress_callback(percent, message)
        except Exception:
            pass

    @staticmethod
    def _sanitize_formula_references(workbook) -> None:
        """
        [1]ALL!E:E のような疑似外部参照を、内部参照 ALL!E:E に戻す。
        重くなりすぎないよう、実際に値を持っているセルだけを見る。
        """
        bracketed_ref_pattern = re.compile(r"\[(\d+)\](?=[^!]+!)")

        for ws in workbook.worksheets:
            cells = getattr(ws, "_cells", None)
            if not cells:
                continue

            for cell in cells.values():
                value = cell.value

                if not isinstance(value, str):
                    continue
                if not value.startswith("="):
                    continue
                if "[" not in value or "]" not in value:
                    continue

                normalized = bracketed_ref_pattern.sub("", value)

                if normalized != value:
                    cell.value = normalized