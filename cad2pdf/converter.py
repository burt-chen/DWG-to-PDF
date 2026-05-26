"""DWG / DXF → PDF 核心轉檔邏輯。

對外公開：
    convert_dwg(dwg_path, pdf_path)         單檔轉換
    convert_folder(folder, mode, ...)       批次轉換（單檔或合併）
"""

from __future__ import annotations

import tempfile
from enum import Enum
from pathlib import Path
from typing import Callable, Iterable

import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from ezdxf.addons.drawing.config import Configuration
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from pypdf import PdfWriter

from .oda import dwg_to_dxf, find_oda_converter


# 中文字型 fallback 順序 — 工程圖常用的繁/簡中文 TrueType
_CJK_FONTS = [
    "Microsoft JhengHei",  # Win 預設繁中
    "Microsoft YaHei",     # Win 預設簡中
    "MingLiU",             # 細明體
    "PMingLiU",
    "SimHei",
    "SimSun",
    "Noto Sans CJK TC",
    "Noto Sans CJK SC",
    "Arial Unicode MS",
]


def _setup_matplotlib_cjk() -> None:
    """讓 matplotlib 的 PDF 輸出能正確顯示中文。

    1. 把中文字型加進 sans-serif fallback 清單最前面
    2. pdf.fonttype = 42 → 用 TrueType,字型 subset 嵌入 PDF,確保 CJK glyph
       不會掉(預設 type3 對 CJK 支援差,常出現方塊)
    """
    current = list(matplotlib.rcParams.get("font.sans-serif", []))
    # 把 CJK 字型插到最前面,但保留原本 fallback
    merged = _CJK_FONTS + [f for f in current if f not in _CJK_FONTS]
    matplotlib.rcParams["font.family"] = "sans-serif"
    matplotlib.rcParams["font.sans-serif"] = merged
    matplotlib.rcParams["axes.unicode_minus"] = False
    matplotlib.rcParams["pdf.fonttype"] = 42  # TrueType, subset embedded


def _setup_ezdxf_cjk_mapping() -> None:
    """讓 ezdxf 把工程圖常用 SHX 中文字型對應到系統 TrueType。

    AutoCAD/Tekla 出的 DWG 文字 style 常指向 SHX 字型(例如 chineset.shx、
    bigfont.shx),ezdxf 找不到 SHX 解析器時會 fallback 到不含 CJK 的字型,
    導致中文渲染成方塊。手動把這些常見名字映射到系統的中文 TrueType。
    """
    target = None
    try:
        from ezdxf.fonts import fonts as ezdxf_fonts
        # 確保系統字型已掃描(預設應該已跑)
        try:
            ezdxf_fonts.build_system_font_cache()
        except Exception:
            pass

        # 找一個系統實際有的中文字型當對應目標
        for name in _CJK_FONTS:
            try:
                face = ezdxf_fonts.find_font_face(name)
            except Exception:
                face = None
            if face and getattr(face, "family", None):
                target = face.family
                break

        if target is None:
            return

        # 常見 SHX / 字型風格名 → 中文 TrueType
        shx_names = [
            "chineset.shx", "chinesetbig.shx", "bigfont.shx",
            "hzdx.shx", "hztxt.shx", "hzfs.shx", "gbcbig.shx",
            "extfont.shx", "extfont2.shx",
            "@MingLiU", "MingLiU", "PMingLiU", "標楷體", "DFKai-SB",
        ]
        # ezdxf 1.4 用 font_face_mapping(若 API 存在)
        mapping = getattr(ezdxf_fonts, "font_face_mapping", None)
        if isinstance(mapping, dict):
            for name in shx_names:
                mapping.setdefault(name.lower(), target)
    except Exception:
        # 字型 fallback 是錦上添花,失敗就讓 ezdxf 用自己的 default
        pass


# import 時就跑一次,確保下游使用前環境已就緒
_setup_matplotlib_cjk()
_setup_ezdxf_cjk_mapping()


class ConvertMode(str, Enum):
    """批次轉檔模式。"""
    SEPARATE = "separate"   # 每個 DWG 對應一個 PDF
    MERGED = "merged"       # 整個資料夾合併成單一 PDF


ProgressCb = Callable[[int, int, str], None]
"""進度回呼：(目前完成數, 總數, 目前處理檔名)"""


def _render_dxf_to_pdf(dxf_path: Path, pdf_path: Path) -> None:
    """用 ezdxf + matplotlib 把 DXF 渲染成單頁 PDF。"""
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()

    fig = plt.figure(figsize=(16.5, 11.7))  # A3 橫式 (英吋)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()

    ctx = RenderContext(doc)
    config = Configuration(background_policy=None)
    backend = MatplotlibBackend(ax)
    Frontend(ctx, backend, config=config).draw_layout(msp, finalize=True)

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(str(pdf_path), dpi=300, bbox_inches="tight", pad_inches=0.1)
    plt.close(fig)


def convert_dwg(
    dwg_path: str | Path,
    pdf_path: str | Path | None = None,
    oda_exe: str | Path | None = None,
) -> Path:
    """單一 DWG → PDF。

    參數：
        dwg_path: 來源 DWG 檔路徑
        pdf_path: 輸出 PDF 路徑，省略時放在 DWG 同資料夾、同檔名 .pdf
        oda_exe: ODAFileConverter.exe 路徑，省略時自動偵測

    回傳：產生的 PDF 路徑
    """
    dwg_path = Path(dwg_path).resolve()
    if not dwg_path.is_file():
        raise FileNotFoundError(dwg_path)

    if pdf_path is None:
        pdf_path = dwg_path.with_suffix(".pdf")
    else:
        pdf_path = Path(pdf_path).resolve()

    exe = Path(oda_exe) if oda_exe else find_oda_converter()

    with tempfile.TemporaryDirectory(prefix="cad2pdf_dxf_") as tmp:
        dxf = dwg_to_dxf(dwg_path, out_dir=Path(tmp), oda_exe=exe)
        _render_dxf_to_pdf(dxf, pdf_path)

    return pdf_path


def _iter_dwg_files(folder: Path) -> list[Path]:
    """資料夾內的 DWG 檔（不遞迴），依檔名排序。"""
    files = [p for p in folder.iterdir()
             if p.is_file() and p.suffix.lower() == ".dwg"]
    files.sort(key=lambda p: p.name.lower())
    return files


def convert_folder(
    folder: str | Path,
    mode: ConvertMode = ConvertMode.SEPARATE,
    output_dir: str | Path | None = None,
    merged_name: str | None = None,
    oda_exe: str | Path | None = None,
    progress: ProgressCb | None = None,
    recursive: bool = False,
) -> list[Path]:
    """批次轉換資料夾內的 DWG 檔。

    參數：
        folder: 來源資料夾
        mode: SEPARATE = 每檔一份 PDF；MERGED = 合併成單一 PDF
        output_dir: 輸出資料夾，省略時 = folder
        merged_name: MERGED 模式下的輸出檔名（不含 .pdf），
                     省略時用 folder 名稱
        oda_exe: ODAFileConverter.exe 路徑
        progress: 進度回呼
        recursive: 是否遞迴子資料夾

    回傳：產生的 PDF 路徑清單
    """
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(folder)

    out_dir = Path(output_dir).resolve() if output_dir else folder
    out_dir.mkdir(parents=True, exist_ok=True)

    if recursive:
        dwgs = sorted(
            (p for p in folder.rglob("*") if p.is_file() and p.suffix.lower() == ".dwg"),
            key=lambda p: str(p).lower(),
        )
    else:
        dwgs = _iter_dwg_files(folder)

    if not dwgs:
        return []

    exe = Path(oda_exe) if oda_exe else find_oda_converter()
    total = len(dwgs)
    produced: list[Path] = []

    if mode == ConvertMode.SEPARATE:
        for i, dwg in enumerate(dwgs, 1):
            if progress:
                progress(i - 1, total, dwg.name)
            # 維持子資料夾結構（recursive 模式時）
            rel = dwg.relative_to(folder) if recursive else Path(dwg.name)
            pdf = out_dir / rel.with_suffix(".pdf")
            convert_dwg(dwg, pdf, oda_exe=exe)
            produced.append(pdf)
            if progress:
                progress(i, total, dwg.name)

    elif mode == ConvertMode.MERGED:
        merged_name = merged_name or folder.name
        merged_pdf = out_dir / f"{merged_name}.pdf"
        writer = PdfWriter()
        with tempfile.TemporaryDirectory(prefix="cad2pdf_merge_") as tmp:
            tmp_dir = Path(tmp)
            for i, dwg in enumerate(dwgs, 1):
                if progress:
                    progress(i - 1, total, dwg.name)
                page_pdf = tmp_dir / f"{i:04d}_{dwg.stem}.pdf"
                convert_dwg(dwg, page_pdf, oda_exe=exe)
                writer.append(str(page_pdf))
                if progress:
                    progress(i, total, dwg.name)
            with open(merged_pdf, "wb") as f:
                writer.write(f)
        produced.append(merged_pdf)

    else:
        raise ValueError(f"未知 mode: {mode}")

    return produced
