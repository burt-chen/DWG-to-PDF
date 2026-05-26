"""ODA File Converter 的偵測與呼叫。

ODA File Converter 是 Open Design Alliance 提供的免費 DWG/DXF 轉檔工具，
本模組負責找到它的安裝位置、並用它把 DWG 批次轉成 DXF。

下載：https://www.opendesign.com/guestfiles/oda_file_converter
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path


class OdaNotFoundError(RuntimeError):
    """找不到 ODA File Converter 時拋出。"""


# Windows 常見安裝路徑（依版本可能略有不同，以 glob 比對）
_WINDOWS_INSTALL_GLOBS = [
    r"C:\Program Files\ODA\ODAFileConverter*\ODAFileConverter.exe",
    r"C:\Program Files (x86)\ODA\ODAFileConverter*\ODAFileConverter.exe",
]

# 專案內 bundled installer 預設位置（相對 repo 根目錄）
_BUNDLED_INSTALLER_DIR = Path(__file__).resolve().parent.parent / "installer"


def find_oda_converter(explicit_path: str | None = None) -> Path:
    """尋找 ODAFileConverter.exe，找不到就拋出 OdaNotFoundError。

    搜尋順序：
        1. 函式參數 explicit_path
        2. 環境變數 ODA_CONVERTER
        3. PATH 上的 ODAFileConverter / ODAFileConverter.exe
        4. Windows 常見安裝路徑
    """
    candidates: list[str] = []

    if explicit_path:
        candidates.append(explicit_path)

    env_path = os.environ.get("ODA_CONVERTER")
    if env_path:
        candidates.append(env_path)

    which = shutil.which("ODAFileConverter") or shutil.which("ODAFileConverter.exe")
    if which:
        candidates.append(which)

    for pattern in _WINDOWS_INSTALL_GLOBS:
        # glob 在 Path 上比較不直覺，這邊用 Path.glob 從 anchor 開始
        root = Path(pattern).anchor
        rel = Path(pattern).relative_to(root)
        for match in Path(root).glob(str(rel)):
            candidates.append(str(match))

    for c in candidates:
        p = Path(c)
        if p.is_file():
            return p

    raise OdaNotFoundError(
        "找不到 ODA File Converter。請至 "
        "https://www.opendesign.com/guestfiles/oda_file_converter 下載安裝，"
        "或設定環境變數 ODA_CONVERTER 指向 ODAFileConverter.exe。"
    )


def find_bundled_installer(installer_dir: Path | None = None) -> Path | None:
    """尋找專案內 installer/ 資料夾下的 ODA installer。

    支援 .exe 與 .msi 兩種格式。回傳第一個符合的檔案路徑；找不到回傳 None。
    """
    d = Path(installer_dir) if installer_dir else _BUNDLED_INSTALLER_DIR
    if not d.is_dir():
        return None
    matches = sorted(d.glob("ODAFileConverter*.exe")) + \
              sorted(d.glob("ODAFileConverter*.msi"))
    return matches[0] if matches else None


def run_installer(installer_path: Path, wait: bool = True) -> int:
    """執行 ODA installer（支援 .exe / .msi）。

    參數：
        installer_path: installer 檔案路徑
        wait: 是否等 installer 結束才回傳

    回傳：0（成功）

    注意：installer 需要 UAC 提權，使用者會看到 Windows 權限提示視窗。
    .msi 會透過 msiexec /i 執行。
    """
    installer_path = Path(installer_path).resolve()
    if not installer_path.is_file():
        raise FileNotFoundError(installer_path)

    import ctypes
    SW_SHOWNORMAL = 1
    suffix = installer_path.suffix.lower()

    if suffix == ".msi":
        # MSI 必須透過 msiexec 執行
        exe = "msiexec.exe"
        args = f'/i "{installer_path}"'
    elif suffix == ".exe":
        exe = str(installer_path)
        args = None
    else:
        raise ValueError(f"不支援的 installer 格式：{suffix}（僅支援 .exe / .msi）")

    # "runas" verb = 要求以管理員身份執行
    rc = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", exe, args, str(installer_path.parent), SW_SHOWNORMAL,
    )
    if rc <= 32:
        raise RuntimeError(f"ShellExecuteW 失敗 (rc={rc})，使用者可能按了取消")

    if not wait:
        return 0

    # ShellExecuteW 不會給我們 process handle，所以用 polling 等 installer 結束
    # 偵測方式：等 ODAFileConverter.exe 出現在常見安裝路徑（最多等 10 分鐘）
    import time
    start = time.time()
    timeout = 600
    while time.time() - start < timeout:
        time.sleep(2)
        try:
            find_oda_converter()
            return 0
        except OdaNotFoundError:
            continue

    raise TimeoutError("等待安裝完成超時（10 分鐘）")


def dwg_to_dxf(
    dwg_path: Path,
    out_dir: Path | None = None,
    oda_exe: Path | None = None,
    dxf_version: str = "ACAD2018",
) -> Path:
    """將單一 DWG 檔轉成 DXF，回傳產生的 DXF 路徑。

    ODA File Converter 是「資料夾轉資料夾」的工具，沒有單檔模式，
    所以這裡在暫存資料夾建一個只放單檔的 input 子資料夾、再呼叫。
    """
    dwg_path = Path(dwg_path).resolve()
    if not dwg_path.is_file():
        raise FileNotFoundError(dwg_path)

    exe = oda_exe or find_oda_converter()
    out_dir = Path(out_dir).resolve() if out_dir else dwg_path.parent

    with tempfile.TemporaryDirectory(prefix="cad2pdf_") as tmp:
        tmp_path = Path(tmp)
        in_dir = tmp_path / "in"
        out_tmp = tmp_path / "out"
        in_dir.mkdir()
        out_tmp.mkdir()

        # 用 hardlink 避免複製大檔；失敗就 fallback 到複製
        link_target = in_dir / dwg_path.name
        try:
            os.link(dwg_path, link_target)
        except OSError:
            shutil.copy2(dwg_path, link_target)

        # ODA File Converter 參數順序（位置式）：
        #   InputFolder OutputFolder OutputVer OutputFormat Recurse Audit [Filter]
        # OutputFormat: ACAD2018 DXF = "DXF"; OutputVer 與 OutputFormat 一起決定版本
        cmd = [
            str(exe),
            str(in_dir),
            str(out_tmp),
            dxf_version,
            "DXF",
            "0",   # Recurse: 0 = 不遞迴
            "1",   # Audit: 1 = 自動修復
            "*.DWG",
        ]
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=300
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"ODA File Converter 執行失敗 (exit {result.returncode}):\n"
                f"stdout: {result.stdout}\nstderr: {result.stderr}"
            )

        produced = out_tmp / (dwg_path.stem + ".dxf")
        if not produced.is_file():
            # 有些版本會大小寫不一致，再 fallback 掃描
            matches = list(out_tmp.glob("*.dxf")) + list(out_tmp.glob("*.DXF"))
            if not matches:
                raise RuntimeError(
                    f"ODA 沒有產出 DXF 檔。輸出資料夾內容：{list(out_tmp.iterdir())}"
                )
            produced = matches[0]

        out_dir.mkdir(parents=True, exist_ok=True)
        final = out_dir / (dwg_path.stem + ".dxf")
        shutil.move(str(produced), final)
        return final
