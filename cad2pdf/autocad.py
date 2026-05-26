"""AutoCAD COM backend — 透過 pywin32 驅動 AutoCAD 原生 PLOT 指令。

需求：
    - Windows
    - 安裝 AutoCAD 完整版（LT 版不支援 COM 自動化）
    - pip install pywin32

對外公開：
    is_autocad_available()                          檢查環境是否可用
    acad_convert_dwg(dwg_path, pdf_path)            單檔轉換
    acad_convert_folder(folder, mode, ...)          批次轉換
"""

from __future__ import annotations

import os
import tempfile
import time
from enum import Enum
from pathlib import Path
from typing import Callable

def _ensure_pywin32_dll_path() -> None:
    """讓 launcher 用 pip install --target 安裝的 pywin32 能找到 DLL。

    pywin32 的 pywintypes / pythoncom 是動態載入 pywin32_system32/*.dll，
    正常 install 流程會跑 post-install 把 DLL 註冊到 Python 根目錄；
    但 --target 安裝跳過 post-install，DLL 還在 pywin32_system32/ 沒移動，
    Python 找不到就會炸 ImportError: No module named 'pywintypes'。

    對策：在 import 前用 os.add_dll_directory() 把 pywin32_system32/
    加進 Windows 的 DLL 搜尋路徑。
    """
    import os
    from pathlib import Path

    # cad2pdf 的上一層 = 工具根目錄（含 pywin32_system32/ 與其他 site-packages）
    tool_root = Path(__file__).resolve().parent.parent
    syspath = tool_root / "pywin32_system32"
    if syspath.is_dir() and hasattr(os, "add_dll_directory"):
        try:
            os.add_dll_directory(str(syspath))
        except (OSError, ValueError):
            pass
    # 同時加進 PATH,給較舊的 C extension 用
    if syspath.is_dir():
        os.environ["PATH"] = str(syspath) + os.pathsep + os.environ.get("PATH", "")


_ensure_pywin32_dll_path()

try:
    import win32com.client
    import pythoncom
    _PYWIN32_OK = True
    _PYWIN32_ERR = None
except ImportError as e:
    _PYWIN32_OK = False
    _PYWIN32_ERR = e

from pypdf import PdfWriter


class AutoCadNotAvailableError(RuntimeError):
    """無法連到 AutoCAD（未安裝、或安裝的是 LT 版）。"""


class ConvertMode(str, Enum):
    SEPARATE = "separate"
    MERGED = "merged"


ProgressCb = Callable[[int, int, str], None]


# AutoCAD 內建的 PDF plotter 名稱
_PDF_PLOTTER = "DWG To PDF.pc3"

# 預設紙張 / 樣式（可被覆寫）
_DEFAULT_PAPER = "ISO_full_bleed_A3_(420.00_x_297.00_MM)"
_DEFAULT_STYLE = "monochrome.ctb"


def _com_retry(callable_, retries: int = 8, delay: float = 0.5):
    """COM 呼叫的重試包裝 — AutoCAD 忙碌時會丟 RPC_E_CALL_REJECTED
    (-2147418111 / 0x80010001) 或 RPC_E_SERVERCALL_RETRYLATER
    (-2147417846 / 0x8001010A)，等一下再試通常會通。
    """
    last_exc = None
    for _ in range(retries):
        try:
            return callable_()
        except pythoncom.com_error as e:  # type: ignore[union-attr]
            last_exc = e
            hr = e.args[0] if e.args else 0
            if hr in (-2147418111, -2147417846):  # 忙線可重試
                time.sleep(delay)
                continue
            raise
    if last_exc:
        raise last_exc


def _dispatch_autocad():
    """取得 AutoCAD Application 物件 — 優先用 EnsureDispatch (early binding)，
    失敗時 fallback 到 Dispatch (late binding)。

    早期繫結能正確識別 Document / Layout 等介面，避免 dynamic dispatch 下
    Documents.Open() 回傳的物件型別不明、後續 .ActiveLayout 取屬性失敗。
    """
    try:
        return win32com.client.gencache.EnsureDispatch("AutoCAD.Application")
    except Exception:
        # gen_py cache 損壞 / 沒寫入權限時 fallback
        return win32com.client.Dispatch("AutoCAD.Application")


def is_autocad_available() -> tuple[bool, str]:
    """檢查 AutoCAD COM 是否可用。

    回傳：(是否可用, 說明訊息)
    """
    if not _PYWIN32_OK:
        return False, f"未安裝 pywin32 套件：{_PYWIN32_ERR}"

    try:
        pythoncom.CoInitialize()
        try:
            acad = _dispatch_autocad()
            version = acad.Version
            return True, f"已偵測到 AutoCAD（版本 {version}）"
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        return False, (
            f"找不到 AutoCAD（{e}）。\n"
            "請確認已安裝 AutoCAD 完整版（LT 版不支援 COM 自動化）。"
        )


class _AutoCadSession:
    """AutoCAD 應用程式 session 管理 — context manager。

    開啟時連線 / 啟動 AutoCAD，關閉時恢復視窗顯示狀態但不關閉 AutoCAD
    （避免每檔都重啟，大幅加速批次處理）。
    """

    def __init__(self, visible: bool = False):
        self.visible = visible
        self.acad = None
        self._previous_visible = None

    def __enter__(self):
        if not _PYWIN32_OK:
            raise AutoCadNotAvailableError(
                f"未安裝 pywin32 套件：{_PYWIN32_ERR}"
            )
        pythoncom.CoInitialize()
        try:
            self.acad = _dispatch_autocad()
        except Exception as e:
            pythoncom.CoUninitialize()
            raise AutoCadNotAvailableError(
                f"無法啟動 AutoCAD：{e}"
            ) from e

        try:
            self._previous_visible = self.acad.Visible
            self.acad.Visible = self.visible
        except Exception:
            # 設定 Visible 失敗不致命,繼續走
            self._previous_visible = None
        return self

    def __exit__(self, exc_type, exc, tb):
        # cleanup 絕不能拋例外蓋過呼叫端的真正錯誤
        try:
            if self.acad is not None and self._previous_visible is not None:
                # 用 retry 包裝,AutoCAD 可能還在處理上次動作而拒絕呼叫
                try:
                    _com_retry(
                        lambda: setattr(self.acad, "Visible", self._previous_visible)
                    )
                except Exception:
                    pass
        finally:
            self.acad = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def plot_dwg(
        self,
        dwg_path: Path,
        pdf_path: Path,
        paper: str = _DEFAULT_PAPER,
        style: str = _DEFAULT_STYLE,
    ) -> None:
        """打開 DWG 並輸出成 PDF。"""
        dwg_path = Path(dwg_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        # 開啟 DWG。Documents.Open() 在 late-binding 下回傳值有時不可靠
        # (拿到的物件 .ActiveLayout 會炸 AttributeError: Open.ActiveLayout),
        # 安全網:若 open_result 取不到屬性就改從 ActiveDocument 拿。
        # 用 retry 包裝,避免 AutoCAD 載入中拒絕 COM 呼叫。
        open_result = _com_retry(
            lambda: self.acad.Documents.Open(str(dwg_path), True)  # read-only
        )
        doc = open_result if (open_result is not None and hasattr(open_result, "ActiveLayout")) \
              else self.acad.ActiveDocument

        try:
            try:
                doc.SetVariable("BACKGROUNDPLOT", 0)
            except Exception:
                pass  # 部分版本不允許設定，忽略

            layout = _com_retry(lambda: doc.ActiveLayout)

            # 設定 plotter 與紙張
            try:
                layout.RefreshPlotDeviceInfo()
            except Exception:
                pass

            try:
                layout.ConfigName = _PDF_PLOTTER
            except Exception as e:
                raise RuntimeError(
                    f"AutoCAD 找不到 plotter '{_PDF_PLOTTER}'，"
                    f"請確認 AutoCAD 安裝完整：{e}"
                ) from e

            try:
                layout.CanonicalMediaName = paper
            except Exception:
                # 紙張名稱因版本不同可能略有差異，失敗就用預設
                pass

            try:
                layout.StyleSheet = style
            except Exception:
                pass

            # Plot 範圍 = Extents（圖面範圍），縮放 = Fit
            try:
                layout.PlotType = 1          # acExtents
                layout.StandardScale = 0      # acScaleToFit
                layout.CenterPlot = True
                layout.PlotWithLineweights = True
            except Exception:
                pass

            # 輸出 — Plot 也用 retry,大檔渲染中 COM 可能短暫拒絕
            plot = doc.Plot
            ok = _com_retry(lambda: plot.PlotToFile(str(pdf_path)))
            if not ok:
                raise RuntimeError(
                    f"AutoCAD PlotToFile 回傳失敗：{dwg_path}"
                )

            # 等檔案落地（AutoCAD 有時非同步）
            for _ in range(30):
                if pdf_path.is_file() and pdf_path.stat().st_size > 0:
                    break
                time.sleep(0.2)

            if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
                raise RuntimeError(f"PDF 沒有產生或為空檔：{pdf_path}")

        finally:
            try:
                _com_retry(lambda: doc.Close(False))  # 不存檔
            except Exception:
                pass


def acad_convert_dwg(
    dwg_path: str | Path,
    pdf_path: str | Path | None = None,
    visible: bool = False,
) -> Path:
    """單一 DWG → PDF（AutoCAD COM 版本）。"""
    dwg_path = Path(dwg_path).resolve()
    if pdf_path is None:
        pdf_path = dwg_path.with_suffix(".pdf")
    else:
        pdf_path = Path(pdf_path).resolve()

    with _AutoCadSession(visible=visible) as session:
        session.plot_dwg(dwg_path, pdf_path)

    return pdf_path


def _iter_dwg_files(folder: Path, recursive: bool) -> list[Path]:
    if recursive:
        files = [p for p in folder.rglob("*")
                 if p.is_file() and p.suffix.lower() == ".dwg"]
    else:
        files = [p for p in folder.iterdir()
                 if p.is_file() and p.suffix.lower() == ".dwg"]
    files.sort(key=lambda p: str(p).lower())
    return files


def acad_convert_folder(
    folder: str | Path,
    mode: ConvertMode = ConvertMode.SEPARATE,
    output_dir: str | Path | None = None,
    merged_name: str | None = None,
    visible: bool = False,
    progress: ProgressCb | None = None,
    recursive: bool = False,
) -> list[Path]:
    """批次轉換資料夾（AutoCAD COM 版本）。

    參數同 cad2pdf.convert_folder，但無 oda_exe，多了 visible。
    """
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(folder)

    out_dir = Path(output_dir).resolve() if output_dir else folder
    out_dir.mkdir(parents=True, exist_ok=True)

    dwgs = _iter_dwg_files(folder, recursive)
    if not dwgs:
        return []

    total = len(dwgs)
    produced: list[Path] = []

    with _AutoCadSession(visible=visible) as session:
        if mode == ConvertMode.SEPARATE:
            for i, dwg in enumerate(dwgs, 1):
                if progress:
                    progress(i - 1, total, dwg.name)
                rel = dwg.relative_to(folder) if recursive else Path(dwg.name)
                pdf = out_dir / rel.with_suffix(".pdf")
                session.plot_dwg(dwg, pdf)
                produced.append(pdf)
                if progress:
                    progress(i, total, dwg.name)

        elif mode == ConvertMode.MERGED:
            merged_name = merged_name or folder.name
            merged_pdf = out_dir / f"{merged_name}.pdf"
            writer = PdfWriter()
            with tempfile.TemporaryDirectory(prefix="acad_merge_") as tmp:
                tmp_dir = Path(tmp)
                for i, dwg in enumerate(dwgs, 1):
                    if progress:
                        progress(i - 1, total, dwg.name)
                    page_pdf = tmp_dir / f"{i:04d}_{dwg.stem}.pdf"
                    session.plot_dwg(dwg, page_pdf)
                    writer.append(str(page_pdf))
                    if progress:
                        progress(i, total, dwg.name)
                with open(merged_pdf, "wb") as f:
                    writer.write(f)
            produced.append(merged_pdf)

        else:
            raise ValueError(f"未知 mode: {mode}")

    return produced
