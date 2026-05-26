"""嵌入式包裝 — 讓 DWG → PDF 轉檔工具 跑在 Launcher 的分頁裡。

實作 create_frame(parent) -> ttk.Frame，由 Launcher 動態載入。
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk

_TOOL_ROOT = Path(__file__).parent

# 確保子套件（cad2pdf）能被 import — 把工具根目錄加進 sys.path
if str(_TOOL_ROOT) not in sys.path:
    sys.path.insert(0, str(_TOOL_ROOT))


def _load_tool():
    """用 importlib 從絕對路徑載入 gui.py，給唯一模組名避免衝突。"""
    spec = importlib.util.spec_from_file_location(
        "_dwg_to_pdf_tool", _TOOL_ROOT / "gui.py"
    )
    mod = importlib.util.module_from_spec(spec)
    sys.modules["_dwg_to_pdf_tool"] = mod
    spec.loader.exec_module(mod)
    return mod


_tool = _load_tool()
_App = _tool.App


class _EmbeddedApp(_App):
    """把 App 嵌進任意 Tkinter widget（不需要 tk.Tk、不調整視窗屬性）。"""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)


def create_frame(parent: tk.Widget) -> ttk.Frame:
    frame = ttk.Frame(parent)
    _EmbeddedApp(frame)
    return frame
