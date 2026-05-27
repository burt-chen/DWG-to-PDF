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
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Callable


# ───────── 診斷 log:寫到 %TEMP%\cad2pdf_acad_diag.log ─────────
# 中文亂碼這類問題在 BIMer 端遠端 debug 不易,把 AutoCAD 內部實際狀態
# (FONTMAP / PDFSHX / TextStyle.fontFile 等) 寫到 log 檔給使用者貼回來。
_DIAG_LOG = Path(tempfile.gettempdir()) / "cad2pdf_acad_diag.log"


def _diag(msg: str) -> None:
    """寫一行診斷訊息到 log 檔。失敗安靜放過。"""
    try:
        with open(_DIAG_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().isoformat(timespec='seconds')}] {msg}\n")
    except Exception:
        pass


def _diag_reset() -> None:
    """每次 GUI 點開始轉檔時呼叫,把 log 清空,只留本次紀錄。"""
    try:
        _DIAG_LOG.unlink(missing_ok=True)
    except Exception:
        pass


def get_diag_log_path() -> Path:
    return _DIAG_LOG

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


def _build_cjk_fontmap_file() -> Path | None:
    """建立一個 AutoCAD fontmap (.fmp) 檔,把常見西文/SHX 字型映射到中文 ttf。

    AutoCAD 的 FONTMAP 系統變數指向這個檔,plot / regen 時會用內含的對應表
    替換字型。這是「全域」的字型替代,比改 doc.TextStyles 更可靠 ——
    即使 STYLE.fontFile 改不動,plot 出來時 AutoCAD 仍會跑 fontmap 替換。

    挑系統實際有的中文字型當目標。回傳 fmp 檔路徑;系統無中文字型回 None。
    """
    import os
    import tempfile

    # Windows 字型目錄,優先嘗試的中文 ttf
    fonts_dir = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    candidates = [
        "msjh.ttc", "msjh.ttf",      # 微軟正黑體
        "msyh.ttc", "msyh.ttf",      # 微軟雅黑
        "mingliu.ttc", "pmingliu.ttf",  # 細明體
        "NotoSansTC-VF.ttf", "NotoSansHK-VF.ttf", "NotoSansSC-VF.ttf",
        "simhei.ttf", "simsun.ttc",
    ]
    target = next(
        (c for c in candidates if (fonts_dir / c).is_file()),
        None,
    )
    if target is None:
        return None

    # AutoCAD .fmp 格式: 一行一條 "原字型;替換字型"
    # 不分 SHX / TTF,常見會走 fallback 的字型都列出
    mappings = [
        # 西文 TTF (DWG STYLE 內最常見的 misconfig)
        "arial;" + target,
        "arial.ttf;" + target,
        "ARIAL;" + target,
        "ARIAL.TTF;" + target,
        "tahoma;" + target,
        "tahoma.ttf;" + target,
        "verdana;" + target,
        "calibri;" + target,
        # AutoCAD 預設西文 SHX (Standard style 預設用 txt)
        "txt;" + target,
        "txt.shx;" + target,
        "simplex;" + target,
        "simplex.shx;" + target,
        "romans;" + target,
        "romans.shx;" + target,
        # 中文 SHX (找不到 SHX 解析時 fallback)
        "chineset;" + target,
        "chineset.shx;" + target,
        "chinesetbig;" + target,
        "bigfont;" + target,
        "bigfont.shx;" + target,
        "hztxt;" + target,
        "hztxt.shx;" + target,
        "hzdx;" + target,
        "gbcbig;" + target,
        "gbcbig.shx;" + target,
        "extfont;" + target,
        "extfont2;" + target,
    ]
    fmp_dir = Path(tempfile.gettempdir()) / "cad2pdf_acad"
    fmp_dir.mkdir(parents=True, exist_ok=True)
    fmp_path = fmp_dir / "cjk_fontmap.fmp"
    fmp_path.write_text("\n".join(mappings) + "\n", encoding="ascii")
    return fmp_path


# 在 module load 時就準備好 fontmap 檔(便宜操作,失敗不致命)
try:
    _CJK_FONTMAP_FILE = _build_cjk_fontmap_file()
except Exception:
    _CJK_FONTMAP_FILE = None


class _AutoCadSession:
    """AutoCAD 應用程式 session 管理 — context manager。

    開啟時連線 / 啟動 AutoCAD，關閉時恢復視窗顯示狀態但不關閉 AutoCAD
    （避免每檔都重啟，大幅加速批次處理）。
    """

    def __init__(self, visible: bool = False):
        self.visible = visible
        self.acad = None
        self._previous_visible = None
        self._previous_fontmap = None

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

        _diag(f"=== session __enter__ ===")
        try:
            _diag(f"AutoCAD.Version = {self.acad.Version}")
        except Exception as e:
            _diag(f"acad.Version 讀取失敗: {e}")
        _diag(f"_CJK_FONTMAP_FILE = {_CJK_FONTMAP_FILE}")
        # 註:SetVariable/GetVariable 在 IAcadDocument 上,不是 IAcadApplication。
        # FONTMAP / PDFSHX 等變數的設定移到 plot_dwg 內 Open doc 之後執行。

        return self

    def __exit__(self, exc_type, exc, tb):
        # cleanup 絕不能拋例外蓋過呼叫端的真正錯誤
        try:
            if self.acad is not None and self._previous_visible is not None:
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

        _diag(f"=== plot_dwg: {dwg_path.name} ===")

        # 開啟 DWG (read-write,Close(False) 不存檔)
        open_result = _com_retry(
            lambda: self.acad.Documents.Open(str(dwg_path), False)
        )
        # 等 AutoCAD 載入完成,避免後續 COM call 立刻 RPC_E_CALL_REJECTED
        time.sleep(1.0)

        # 取 doc:open_result 屬性 access 用 _com_retry 包,然後若不可用
        # 退到 ActiveDocument。不用 hasattr (它會觸發 getattr,失敗無法 retry)。
        def _get_layout():
            return open_result.ActiveLayout if open_result is not None else None

        doc = None
        try:
            _com_retry(_get_layout)
            doc = open_result
        except Exception as e:
            _diag(f"open_result.ActiveLayout 失敗,改用 ActiveDocument: {e}")
            try:
                doc = _com_retry(lambda: self.acad.ActiveDocument)
            except Exception as e2:
                _diag(f"取 ActiveDocument 也失敗: {e2}")
                raise

        # 透過 doc 設變數 (SetVariable 在 IAcadDocument 上,不是 IAcadApplication)
        try:
            doc.SetVariable("PDFSHX", 1)
            _diag(f"doc.SetVariable(PDFSHX, 1) → actual={doc.GetVariable('PDFSHX')}")
        except Exception as e:
            _diag(f"doc.SetVariable(PDFSHX) 失敗: {e}")

        if _CJK_FONTMAP_FILE is not None:
            try:
                doc.SetVariable("FONTMAP", str(_CJK_FONTMAP_FILE))
                _diag(f"doc.SetVariable(FONTMAP) → actual={doc.GetVariable('FONTMAP')!r}")
            except Exception as e:
                _diag(f"doc.SetVariable(FONTMAP) 失敗: {e}")

        # TextStyle.fontFile 改用「絕對路徑」設定,因為 AutoCAD 的字型搜尋
        # 路徑不一定包含 C:\Windows\Fonts,只給檔名會拋 OLE「檔案錯誤」。
        cjk_full_path = None
        if _CJK_FONTMAP_FILE is not None:
            # 從 fmp 內第一條取出目標字型檔名
            try:
                first = _CJK_FONTMAP_FILE.read_text(encoding="ascii").splitlines()[0]
                _, target_name = first.split(";", 1)
                full = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / target_name
                if full.is_file():
                    cjk_full_path = str(full)
            except Exception as e:
                _diag(f"解析中文字型絕對路徑失敗: {e}")
        _diag(f"cjk_full_path for fontFile = {cjk_full_path!r}")

        _modified = 0
        _ts_total = 0
        if cjk_full_path:
            try:
                for ts in doc.TextStyles:
                    _ts_total += 1
                    try:
                        name = ts.Name
                    except Exception:
                        name = "?"
                    try:
                        before = ts.fontFile
                    except Exception:
                        before = "?"
                    set_ok = False
                    # AutoCAD COM:屬性叫 fontFile (從 log 確認大寫 FontFile 不存在)
                    try:
                        ts.fontFile = cjk_full_path
                        set_ok = True
                    except Exception as e:
                        _diag(f"  setattr(fontFile=絕對路徑) on {name!r} 失敗: {e}")
                        # 再退一步試純檔名
                        try:
                            ts.fontFile = Path(cjk_full_path).name
                            set_ok = True
                        except Exception as e2:
                            _diag(f"  setattr(fontFile=檔名) on {name!r} 也失敗: {e2}")
                    try:
                        after = ts.fontFile
                    except Exception:
                        after = "?"
                    _diag(f"  style {name!r}: before={before!r} after={after!r} ok={set_ok}")
                    if set_ok:
                        _modified += 1
                    try:
                        ts.bigFontFile = ""
                    except Exception:
                        pass
            except Exception as e:
                _diag(f"遍歷 TextStyles 例外: {e}")
        _diag(f"TextStyle modify result: total={_ts_total} modified={_modified}")

        # 強制 regen 讓 AutoCAD 重新讀取 style 字型設定
        try:
            doc.Regen(1)  # acAllViewports
            _diag("doc.Regen(acAllViewports) done")
        except Exception as e:
            _diag(f"doc.Regen 失敗: {e}")
            # fallback: SendCommand
            try:
                doc.SendCommand("_REGENALL\n")
                _diag("SendCommand(_REGENALL) sent")
            except Exception as e2:
                _diag(f"SendCommand(_REGENALL) 失敗: {e2}")

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

            # 預先刪掉舊 PDF,避免存在時跳互動 dialog
            try:
                if pdf_path.exists():
                    pdf_path.unlink()
            except Exception:
                pass

            # 輸出 — 用 PlotToFile + DWG-to-PDF.pc3
            # (doc.Export("PDF") 在 AutoCAD 2024 拋「無效的引數」,先不走那條)
            plot = doc.Plot
            ok = _com_retry(lambda: plot.PlotToFile(str(pdf_path)))
            _diag(f"PlotToFile 回傳 ok={ok}")
            if not ok:
                raise RuntimeError(f"AutoCAD PlotToFile 回傳失敗:{dwg_path}")

            # 等檔案落地（AutoCAD 有時非同步）
            for _ in range(60):
                if pdf_path.is_file() and pdf_path.stat().st_size > 0:
                    break
                time.sleep(0.3)

            if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
                raise RuntimeError(f"PDF 沒有產生或為空檔：{pdf_path}")
            _diag(f"PDF 落地 size={pdf_path.stat().st_size}")

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

    # 每次批次開始時清空診斷 log
    _diag_reset()
    _diag(f"acad_convert_folder start: folder={folder} mode={mode} files={len(dwgs)}")

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
