from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from zipfile import ZIP_STORED, ZipFile

from openpyxl.utils import get_column_letter
from openpyxl.writer.excel import ExcelWriter

from core.constants import DATA_START_ROW, HEADER_ROW
from core.mapping import ConditionalCopyRule, CopyRule, DEFAULT_MAPPING, MappingConfig
from core.validator import validate_inputs
from utils.excel_helpers import (
    col_to_index,
    copy_row_style_and_formulas,
    find_last_data_row,
    is_blank,
    iter_non_blank_rows,
    next_sequential_number,
)


class ProcessingError(Exception):
    pass


@dataclass(frozen=True)
class ProcessingResult:
    output_path: Path
    source_row_count: int
    appended_row_count: int


class WorkbookProcessor:
    def __init__(self, mapping: MappingConfig | None = None) -> None:
        self.mapping = mapping or DEFAULT_MAPPING

    def process(
        self,
        monthly_path: str | Path,
        master_path: str | Path,
        output_path: str | Path,
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

        append_row = find_last_data_row(target_ws, DATA_START_ROW) + 1
        appended_count = 0
        progress_step = max(1, total_rows // 100)

        for source_row in source_rows:
            template_row = append_row - 1 if append_row > DATA_START_ROW else DATA_START_ROW
            copy_row_style_and_formulas(target_ws, template_row, append_row)

            self._apply_copy_rules(source_ws, target_ws, source_row, append_row)
            self._apply_conditional_copy_rules(source_ws, target_ws, source_row, append_row)
            self._apply_value_map_copy_rules(source_ws, target_ws, source_row, append_row)
            self._apply_auto_number_rules(target_ws, append_row)
            self._apply_fixed_value_rules(target_ws, append_row)

            append_row += 1
            appended_count += 1

            if appended_count % progress_step == 0 or appended_count == total_rows:
                percent = int((appended_count / total_rows) * 90)
                self._report_progress(
                    progress_callback,
                    percent,
                    f"データ追加中… {appended_count}/{total_rows}",
                )

        self._report_progress(progress_callback, 93, "フィルタを更新しています…")
        self._refresh_auto_filter(target_ws)

        self._enable_recalc_on_open(validated.master_workbook)

        self._report_progress(progress_callback, 95, "数式参照を調整しています…")
        self._sanitize_formula_references(validated.master_workbook)

        self._report_progress(progress_callback, 97, "保存しています…（大きいファイルのため少し時間がかかることがあります）")
        self._save_output_workbook(validated.master_workbook, output_path)

        self._report_progress(progress_callback, 100, "完了しました。")

        return ProcessingResult(
            output_path=output_path,
            source_row_count=total_rows,
            appended_row_count=appended_count,
        )

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
        temp_path = output_path.with_name(f"~{output_path.stem}_tmp.xlsx")

        try:
            if temp_path.exists():
                temp_path.unlink()

            with ZipFile(temp_path, mode="w", compression=ZIP_STORED, allowZip64=True) as archive:
                writer = ExcelWriter(workbook, archive)
                writer.save()

            if output_path.exists():
                output_path.unlink()

            temp_path.replace(output_path)

        except Exception as exc:  # noqa: BLE001
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except OSError:
                pass

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