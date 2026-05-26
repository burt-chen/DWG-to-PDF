"""DWG → PDF 轉檔工具 — tkinter GUI（兩個 tab 分頁：ODA / AutoCAD）。

使用：
    python gui.py
或：
    python main.py
"""

from __future__ import annotations

import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import (
    Tk, StringVar, IntVar, filedialog, messagebox, END, DISABLED, NORMAL,
)
from tkinter import ttk

from cad2pdf import (
    ConvertMode,
    OdaNotFoundError,
    convert_folder,
    find_oda_converter,
    find_bundled_installer,
    run_installer,
    acad_convert_folder,
    is_autocad_available,
)


APP_TITLE = "DWG → PDF 轉檔工具"


class ConverterPanel(ttk.Frame):
    """通用轉檔面板 — 兩個 tab 共用同一份 UI 結構，差別只在 backend 函式。

    backend_run(folder, mode, output_dir, merged_name, recursive, progress, **extras)
        → list[Path]
    extras_builder(panel) → dict
        把 backend 專屬欄位（例如 ODA 路徑）打包成 kwargs 傳給 backend_run。
    backend_check(panel) → tuple[bool, str]
        執行前的環境檢查（找不到 ODA / 找不到 AutoCAD 都會在這擋下）。
    """

    def __init__(
        self,
        master,
        backend_run,
        backend_name: str,
        extras_section_builder=None,
        extras_kwargs_builder=None,
        backend_check=None,
    ):
        super().__init__(master)
        self.backend_run = backend_run
        self.backend_name = backend_name
        self.extras_section_builder = extras_section_builder
        self.extras_kwargs_builder = extras_kwargs_builder or (lambda p: {})
        self.backend_check = backend_check

        self.src_var = StringVar()
        self.out_var = StringVar()
        self.mode_var = StringVar(value=ConvertMode.SEPARATE.value)
        self.recursive_var = IntVar(value=0)
        self.merged_name_var = StringVar()

        self._busy = False
        self._pending_log: list[str] = []
        self.log = None
        self._build()

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 6}

        # 來源
        row = ttk.LabelFrame(self, text="來源 DWG 資料夾")
        row.pack(fill="x", **pad)
        ttk.Entry(row, textvariable=self.src_var).pack(
            side="left", fill="x", expand=True, padx=6, pady=6
        )
        ttk.Button(row, text="瀏覽…", command=self._pick_src).pack(
            side="left", padx=6, pady=6
        )

        # 輸出
        row = ttk.LabelFrame(self, text="輸出 PDF 資料夾（留空 = 同來源）")
        row.pack(fill="x", **pad)
        ttk.Entry(row, textvariable=self.out_var).pack(
            side="left", fill="x", expand=True, padx=6, pady=6
        )
        ttk.Button(row, text="瀏覽…", command=self._pick_out).pack(
            side="left", padx=6, pady=6
        )

        # 模式
        row = ttk.LabelFrame(self, text="輸出模式")
        row.pack(fill="x", **pad)
        ttk.Radiobutton(
            row, text="每個 DWG 一個 PDF",
            variable=self.mode_var, value=ConvertMode.SEPARATE.value,
            command=self._refresh_merged_state,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(
            row, text="合併成單一 PDF（多頁）",
            variable=self.mode_var, value=ConvertMode.MERGED.value,
            command=self._refresh_merged_state,
        ).pack(anchor="w", padx=10, pady=2)
        sub = ttk.Frame(row)
        sub.pack(fill="x", padx=10, pady=4)
        ttk.Label(sub, text="合併檔名：").pack(side="left")
        self.merged_entry = ttk.Entry(sub, textvariable=self.merged_name_var)
        self.merged_entry.pack(side="left", fill="x", expand=True)
        ttk.Label(sub, text=".pdf").pack(side="left")
        ttk.Checkbutton(
            row, text="包含子資料夾", variable=self.recursive_var,
        ).pack(anchor="w", padx=10, pady=4)

        # backend 專屬欄位
        if self.extras_section_builder:
            self.extras_section_builder(self)

        # 操作 + 進度
        bar = ttk.Frame(self)
        bar.pack(fill="x", **pad)
        self.run_btn = ttk.Button(bar, text=f"開始轉檔（{self.backend_name}）",
                                   command=self._start)
        self.run_btn.pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        # log（用 Text 控件以支援自動換行）
        row = ttk.LabelFrame(self, text="訊息")
        row.pack(fill="both", expand=True, **pad)
        log_frame = ttk.Frame(row)
        log_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.log = tk.Text(log_frame, height=6, wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical",
                                   command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # flush 在 log 建立前累積的訊息
        if self._pending_log:
            self.log.configure(state="normal")
            for msg in self._pending_log:
                self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self._pending_log.clear()

        self._refresh_merged_state()

    # ---------- handlers ----------
    def _pick_src(self) -> None:
        d = filedialog.askdirectory(title="選擇 DWG 資料夾")
        if d:
            self.src_var.set(d)
            if not self.merged_name_var.get():
                self.merged_name_var.set(Path(d).name)

    def _pick_out(self) -> None:
        d = filedialog.askdirectory(title="選擇輸出 PDF 資料夾")
        if d:
            self.out_var.set(d)

    def _refresh_merged_state(self) -> None:
        if self.mode_var.get() == ConvertMode.MERGED.value:
            self.merged_entry.configure(state=NORMAL)
        else:
            self.merged_entry.configure(state=DISABLED)

    def log_msg(self, msg: str) -> None:
        if self.log is None:
            # log 控件尚未建立 → 先暫存，等 _build 完成後 flush
            self._pending_log.append(msg)
            return
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _start(self) -> None:
        if self._busy:
            return

        # 環境檢查
        if self.backend_check:
            ok, msg = self.backend_check(self)
            if not ok:
                messagebox.showerror(APP_TITLE, msg)
                self.log_msg(f"✗ {msg}")
                return

        src = self.src_var.get().strip()
        if not src or not Path(src).is_dir():
            messagebox.showwarning(APP_TITLE, "請選擇有效的來源資料夾")
            return

        out = self.out_var.get().strip() or None
        mode = ConvertMode(self.mode_var.get())
        merged_name = self.merged_name_var.get().strip() or None
        recursive = bool(self.recursive_var.get())
        extras = self.extras_kwargs_builder(self)

        self._busy = True
        self.run_btn.configure(state=DISABLED, text="轉檔中…")
        self.progress.configure(value=0, maximum=100)
        self.log_msg(f"=== 開始轉檔（{self.backend_name} / {mode.value}）===")

        t = threading.Thread(
            target=self._run_job,
            args=(src, out, mode, merged_name, recursive, extras),
            daemon=True,
        )
        t.start()

    def _run_job(self, src, out, mode, merged_name, recursive, extras) -> None:
        try:
            def on_progress(done: int, total: int, name: str) -> None:
                pct = int(done / total * 100) if total else 0
                self.after(0, self._update_progress, pct, done, total, name)

            produced = self.backend_run(
                folder=src,
                mode=mode,
                output_dir=out,
                merged_name=merged_name,
                progress=on_progress,
                recursive=recursive,
                **extras,
            )
            self.after(0, self._on_done, produced, None)
        except Exception as e:
            tb = traceback.format_exc()
            self.after(0, self._on_done, None, (e, tb))

    def _update_progress(self, pct: int, done: int, total: int, name: str) -> None:
        self.progress.configure(value=pct)
        self.log_msg(f"[{done}/{total}] {name}")

    def _on_done(self, produced, error) -> None:
        self._busy = False
        self.run_btn.configure(state=NORMAL, text=f"開始轉檔（{self.backend_name}）")
        if error:
            exc, tb = error
            self.log_msg(f"✗ 失敗：{exc}")
            messagebox.showerror(APP_TITLE, f"轉檔失敗：\n{exc}\n\n{tb}")
            return
        self.progress.configure(value=100)
        if not produced:
            self.log_msg("⚠ 沒有找到任何 DWG 檔")
            messagebox.showwarning(APP_TITLE, "來源資料夾內沒有 DWG 檔")
            return
        self.log_msg(f"✓ 完成，共產出 {len(produced)} 個 PDF")
        for p in produced:
            self.log_msg(f"  → {p}")
        messagebox.showinfo(APP_TITLE, f"完成！共產出 {len(produced)} 個 PDF。")


# ---------- 各 tab 的 backend 設定 ----------

def _build_oda_extras(panel: ConverterPanel) -> None:
    panel.oda_var = StringVar()

    # 路徑列
    row = ttk.LabelFrame(panel, text="ODA File Converter（自動偵測）")
    row.pack(fill="x", padx=10, pady=6)
    ttk.Entry(row, textvariable=panel.oda_var).pack(
        side="left", fill="x", expand=True, padx=6, pady=6
    )
    def pick():
        f = filedialog.askopenfilename(
            title="選擇 ODAFileConverter.exe",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")],
        )
        if f:
            panel.oda_var.set(f)
    ttk.Button(row, text="瀏覽…", command=pick).pack(side="left", padx=6, pady=6)

    # 安裝按鈕列（只在偵測不到 ODA、但 installer/ 內有 .exe 時顯示）
    panel.install_frame = ttk.Frame(panel)
    panel.install_label = ttk.Label(
        panel.install_frame, text="", foreground="#b85c00"
    )
    panel.install_label.pack(side="left", padx=10, pady=4)
    panel.install_btn = ttk.Button(
        panel.install_frame, text="執行安裝程式",
        command=lambda: _run_oda_installer(panel),
    )
    panel.install_btn.pack(side="right", padx=10, pady=4)

    # 重新偵測按鈕
    panel.redetect_btn = ttk.Button(
        panel.install_frame, text="重新偵測",
        command=lambda: _redetect_oda(panel),
    )
    panel.redetect_btn.pack(side="right", padx=2, pady=4)

    _redetect_oda(panel)


def _redetect_oda(panel: ConverterPanel) -> None:
    """重新偵測 ODA；依結果顯示／隱藏安裝按鈕列。"""
    try:
        p = find_oda_converter()
        panel.oda_var.set(str(p))
        panel.log_msg(f"✓ 已偵測到 ODA File Converter：{p}")
        panel.install_frame.pack_forget()
    except OdaNotFoundError:
        installer = find_bundled_installer()
        if installer:
            panel.install_label.configure(
                text=f"⚠ 尚未安裝 ODA，已找到安裝程式：{installer.name}"
            )
            panel.install_btn.configure(state=NORMAL)
            panel.install_frame.pack(fill="x", padx=10, pady=4)
            panel.log_msg(
                f"⚠ 尚未安裝 ODA File Converter。"
                f"已找到 installer：{installer}"
            )
        else:
            panel.install_label.configure(
                text="⚠ 尚未安裝 ODA，且 installer/ 資料夾內找不到安裝檔"
            )
            panel.install_btn.configure(state=DISABLED)
            panel.install_frame.pack(fill="x", padx=10, pady=4)
            panel.log_msg(
                "⚠ 尚未安裝 ODA File Converter，"
                "且 installer/ 資料夾內沒有 ODAFileConverter*.exe"
            )


def _run_oda_installer(panel: ConverterPanel) -> None:
    """執行 bundled installer。在背景 thread 等它完成，完成後重新偵測。"""
    installer = find_bundled_installer()
    if not installer:
        messagebox.showerror(
            APP_TITLE, "找不到 installer。請把 ODAFileConverter*.exe 放在 installer/ 資料夾"
        )
        return

    if not messagebox.askyesno(
        APP_TITLE,
        f"即將執行：\n{installer}\n\n"
        "Windows 會跳出權限提示（UAC），按「是」允許。\n"
        "請在 installer 內完成安裝後，本程式會自動偵測。",
    ):
        return

    panel.install_btn.configure(state=DISABLED, text="安裝中…")
    panel.log_msg(f"→ 啟動 installer：{installer}")

    def worker():
        try:
            run_installer(installer, wait=True)
            panel.after(0, lambda: (
                panel.log_msg("✓ ODA 安裝完成，重新偵測…"),
                _redetect_oda(panel),
                panel.install_btn.configure(state=NORMAL, text="執行安裝程式"),
            ))
        except Exception as e:
            panel.after(0, lambda: (
                panel.log_msg(f"✗ 安裝失敗或被取消：{e}"),
                panel.install_btn.configure(state=NORMAL, text="執行安裝程式"),
                messagebox.showwarning(APP_TITLE, f"安裝未完成：\n{e}"),
            ))

    threading.Thread(target=worker, daemon=True).start()


def _oda_extras_kwargs(panel: ConverterPanel) -> dict:
    return {"oda_exe": panel.oda_var.get().strip() or None}


def _oda_check(panel: ConverterPanel) -> tuple[bool, str]:
    path = panel.oda_var.get().strip()
    if path and Path(path).is_file():
        return True, ""
    try:
        find_oda_converter()
        return True, ""
    except OdaNotFoundError as e:
        return False, str(e)


def _build_acad_extras(panel: ConverterPanel) -> None:
    panel.visible_var = IntVar(value=0)
    row = ttk.LabelFrame(panel, text="AutoCAD 設定")
    row.pack(fill="x", padx=10, pady=6)
    ttk.Checkbutton(
        row, text="轉檔時顯示 AutoCAD 視窗（除錯用，平時不需勾）",
        variable=panel.visible_var,
    ).pack(anchor="w", padx=10, pady=4)

    # 偵測 AutoCAD
    ok, msg = is_autocad_available()
    panel.log_msg(("✓ " if ok else "⚠ ") + msg)


def _acad_extras_kwargs(panel: ConverterPanel) -> dict:
    return {"visible": bool(panel.visible_var.get())}


def _acad_check(panel: ConverterPanel) -> tuple[bool, str]:
    ok, msg = is_autocad_available()
    if not ok:
        return False, msg
    return True, ""


# ---------- 主視窗 ----------

class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title(APP_TITLE)
        root.geometry("680x640")
        root.minsize(600, 560)

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # Tab 1: ODA
        self.oda_panel = ConverterPanel(
            nb,
            backend_run=convert_folder,
            backend_name="ODA File Converter",
            extras_section_builder=_build_oda_extras,
            extras_kwargs_builder=_oda_extras_kwargs,
            backend_check=_oda_check,
        )
        nb.add(self.oda_panel, text="ODA File Converter（免費）")

        # Tab 2: AutoCAD
        self.acad_panel = ConverterPanel(
            nb,
            backend_run=acad_convert_folder,
            backend_name="AutoCAD COM",
            extras_section_builder=_build_acad_extras,
            extras_kwargs_builder=_acad_extras_kwargs,
            backend_check=_acad_check,
        )
        nb.add(self.acad_panel, text="AutoCAD COM（需安裝 AutoCAD）")


def main() -> None:
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
