from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk

from core.constants import APP_TITLE
from core.file_namer import build_output_filename, build_output_path, build_renamed_output_path
from core.processor import ProcessingError, WorkbookProcessor
from core.validator import ValidationError
from gui.dialogs import SaveConflictAction, ask_save_conflict, show_error, show_info


class MainWindow(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x340")
        self.resizable(False, False)

        self.monthly_path_var = tk.StringVar()
        self.master_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.output_file_var = tk.StringVar(value=build_output_filename())
        self.status_var = tk.StringVar(value="ファイルを選択してください。")

        self.last_output_path: Path | None = None
        self.processor = WorkbookProcessor()

        self._worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._is_processing = False

        self._build_ui()

    def _build_ui(self) -> None:
        padding = {"padx": 10, "pady": 8}

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(container, text="月次ファイル").grid(row=0, column=0, sticky="w", **padding)
        self.monthly_entry = ttk.Entry(container, textvariable=self.monthly_path_var, width=70)
        self.monthly_entry.grid(row=0, column=1, sticky="ew", **padding)
        self.monthly_browse_button = ttk.Button(container, text="参照", command=self._choose_monthly_file)
        self.monthly_browse_button.grid(row=0, column=2, **padding)

        ttk.Label(container, text="マスターファイル").grid(row=1, column=0, sticky="w", **padding)
        self.master_entry = ttk.Entry(container, textvariable=self.master_path_var, width=70)
        self.master_entry.grid(row=1, column=1, sticky="ew", **padding)
        self.master_browse_button = ttk.Button(container, text="参照", command=self._choose_master_file)
        self.master_browse_button.grid(row=1, column=2, **padding)

        ttk.Label(container, text="出力先フォルダ").grid(row=2, column=0, sticky="w", **padding)
        self.output_dir_entry = ttk.Entry(container, textvariable=self.output_dir_var, width=70)
        self.output_dir_entry.grid(row=2, column=1, sticky="ew", **padding)
        self.output_browse_button = ttk.Button(container, text="参照", command=self._choose_output_dir)
        self.output_browse_button.grid(row=2, column=2, **padding)

        ttk.Label(container, text="出力ファイル名").grid(row=3, column=0, sticky="w", **padding)
        self.output_file_entry = ttk.Entry(container, textvariable=self.output_file_var, width=70)
        self.output_file_entry.grid(row=3, column=1, sticky="ew", **padding)

        ttk.Separator(container, orient="horizontal").grid(row=4, column=0, columnspan=3, sticky="ew", pady=12)

        button_frame = ttk.Frame(container)
        button_frame.grid(row=5, column=0, columnspan=3, sticky="w", padx=10, pady=8)

        self.run_button = ttk.Button(button_frame, text="実行", command=self._run)
        self.run_button.pack(side="left")

        self.open_output_button = ttk.Button(
            button_frame,
            text="出力ファイルを開く",
            command=self._open_output_file,
            state="disabled",
        )
        self.open_output_button.pack(side="left", padx=8)

        self.progress = ttk.Progressbar(
            container,
            mode="determinate",
            length=240,
            maximum=100,
        )
        self.progress.grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=(4, 4))
        self.progress["value"] = 0

        self.status_label = ttk.Label(
            container,
            textvariable=self.status_var,
            foreground="#444444",
            width=80,
            anchor="w",
        )
        self.status_label.grid(row=7, column=0, columnspan=3, sticky="w", padx=10, pady=8)

        container.columnconfigure(1, weight=1)

    def _choose_monthly_file(self) -> None:
        path = filedialog.askopenfilename(
            title="月次ファイルを選択",
            filetypes=[("Excel ファイル", "*.xlsx")],
        )
        if path:
            self.monthly_path_var.set(path)
            if not self.output_dir_var.get():
                self.output_dir_var.set(str(Path(path).parent))

    def _choose_master_file(self) -> None:
        path = filedialog.askopenfilename(
            title="マスターファイルを選択",
            filetypes=[("Excel ファイル", "*.xlsx")],
        )
        if path:
            self.master_path_var.set(path)
            if not self.output_dir_var.get():
                self.output_dir_var.set(str(Path(path).parent))

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.output_dir_var.set(path)

    def _run(self) -> None:
        if self._is_processing:
            return

        try:
            monthly_path = self.monthly_path_var.get().strip()
            master_path = self.master_path_var.get().strip()
            output_dir = self.output_dir_var.get().strip()
            output_file = self.output_file_var.get().strip() or build_output_filename()

            if not monthly_path or not master_path or not output_dir:
                raise ValidationError("月次ファイル、マスターファイル、出力先フォルダをすべて指定してください。")

            output_path = build_output_path(output_dir, output_file)
            output_path = self._resolve_output_path(output_path)
            if output_path is None:
                self.status_var.set("保存をキャンセルしました。")
                return

            self._set_processing_state(True)
            self.progress["value"] = 0
            self.status_var.set("処理を開始しています… (0%)")
            self.update_idletasks()

            worker = threading.Thread(
                target=self._worker_process,
                args=(monthly_path, master_path, output_path),
                daemon=True,
            )
            worker.start()
            self.after(100, self._poll_worker_queue)

        except (ValidationError, ProcessingError) as exc:
            self._set_processing_state(False)
            self.progress["value"] = 0
            self.status_var.set("エラーが発生しました。")
            show_error(self, str(exc))
        except Exception as exc:  # noqa: BLE001
            self._set_processing_state(False)
            self.progress["value"] = 0
            self.status_var.set("想定外のエラーが発生しました。")
            show_error(self, f"想定外のエラーが発生しました。\n{exc}")

    def _worker_process(self, monthly_path: str, master_path: str, output_path: Path) -> None:
        try:
            result = self.processor.process(
                monthly_path=monthly_path,
                master_path=master_path,
                output_path=output_path,
                progress_callback=lambda percent, message: self._worker_queue.put(
                    ("progress", (percent, message))
                ),
            )
            self._worker_queue.put(("success", result))
        except Exception as exc:  # noqa: BLE001
            self._worker_queue.put(("error", exc))

    def _poll_worker_queue(self) -> None:
        processed_any = False

        while True:
            try:
                kind, payload = self._worker_queue.get_nowait()
            except queue.Empty:
                break

            processed_any = True

            if kind == "progress":
                percent, message = payload
                self._on_progress(percent, message)
            elif kind == "success":
                self._on_process_success(payload)
                return
            else:
                self._on_process_error(payload)
                return

        if self._is_processing:
            self.after(50 if processed_any else 100, self._poll_worker_queue)

    def _on_progress(self, percent: int, message: str) -> None:
        self.progress["value"] = max(0, min(100, percent))
        self.status_var.set(f"{message} ({percent}%)")

    def _on_process_success(self, result) -> None:
        self._set_processing_state(False)
        self.last_output_path = result.output_path
        self.open_output_button.config(state="normal")
        self.progress["value"] = 100
        self.status_var.set(
            f"完了: 更新 {result.updated_row_count} 件 / 追加 {result.appended_row_count} 件"
        )

        show_info(
            self,
            (
                "処理が完了しました。\n\n"
                f"読み込み件数: {result.source_row_count}\n"
                f"更新件数: {result.updated_row_count}\n"
                f"追加件数: {result.appended_row_count}\n"
                f"出力先: {result.output_path}"
            ),
        )

    def _on_process_error(self, exc: Exception) -> None:
        self._set_processing_state(False)
        self.progress["value"] = 0
        self.status_var.set("エラーが発生しました。")
        show_error(self, str(exc))

    def _set_processing_state(self, is_processing: bool) -> None:
        self._is_processing = is_processing

        widget_state = "disabled" if is_processing else "normal"

        self.monthly_browse_button.config(state=widget_state)
        self.master_browse_button.config(state=widget_state)
        self.output_browse_button.config(state=widget_state)
        self.run_button.config(state=widget_state)

        self.monthly_entry.config(state=widget_state)
        self.master_entry.config(state=widget_state)
        self.output_dir_entry.config(state=widget_state)
        self.output_file_entry.config(state=widget_state)

        if is_processing:
            self.open_output_button.config(state="disabled")
        else:
            if self.last_output_path is not None:
                self.open_output_button.config(state="normal")

    def _resolve_output_path(self, output_path: Path) -> Path | None:
        if not output_path.exists():
            return output_path

        action = ask_save_conflict(self, output_path)
        if action == SaveConflictAction.OVERWRITE:
            return output_path
        if action == SaveConflictAction.RENAME:
            return build_renamed_output_path(output_path)
        return None

    def _open_output_file(self) -> None:
        if not self.last_output_path:
            return

        path = self.last_output_path
        if not path.exists():
            show_error(self, f"出力ファイルが見つかりません。\n{path}")
            return

        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as exc:  # noqa: BLE001
            show_error(self, f"出力ファイルを開けませんでした。\n{exc}")