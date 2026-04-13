from __future__ import annotations

from enum import Enum
from pathlib import Path
import tkinter as tk
from tkinter import messagebox


class SaveConflictAction(str, Enum):
    OVERWRITE = "overwrite"
    RENAME = "rename"
    CANCEL = "cancel"



def show_error(parent: tk.Misc, message: str) -> None:
    messagebox.showerror("エラー", message, parent=parent)



def show_info(parent: tk.Misc, message: str) -> None:
    messagebox.showinfo("完了", message, parent=parent)



def ask_save_conflict(parent: tk.Misc, path: str | Path) -> SaveConflictAction:
    answer = messagebox.askyesnocancel(
        "保存先ファイルの確認",
        (
            f"同名ファイルが既に存在します。\n\n{Path(path)}\n\n"
            "はい: 上書き\n"
            "いいえ: 別名で保存\n"
            "キャンセル: 保存を中止"
        ),
        parent=parent,
    )
    if answer is True:
        return SaveConflictAction.OVERWRITE
    if answer is False:
        return SaveConflictAction.RENAME
    return SaveConflictAction.CANCEL
