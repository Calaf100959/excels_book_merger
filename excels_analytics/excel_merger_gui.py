from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def is_excel_file(path: Path) -> bool:
    if path.is_dir():
        return False
    if path.name.startswith("~$"):
        return False
    return path.suffix.lower() in EXCEL_EXTS


def safe_default_filename(folder: Path) -> str:
    ts = time.strftime("%Y%m%d_%H%M%S")
    return f"merged_{folder.name}_{ts}.xlsx"


def sanitize_sheet_name(name: str) -> str:
    # Excel sheet name constraints: max 31 chars, cannot contain : \ / ? * [ ]
    illegal = {":", "\\", "/", "?", "*", "[", "]"}
    cleaned = "".join("_" if ch in illegal else ch for ch in name).strip()
    return cleaned or "Sheet"


def truncate_sheet_name(name: str, max_len: int = 31) -> str:
    return name[:max_len]


def make_unique_sheet_name(existing_names: set[str], desired: str) -> str:
    base = truncate_sheet_name(sanitize_sheet_name(desired), 31)
    if base not in existing_names:
        return base

    for n in range(2, 10000):
        suffix = f"_{n}"
        max_base = 31 - len(suffix)
        candidate = truncate_sheet_name(base, max_base) + suffix
        if candidate not in existing_names:
            return candidate

    raise RuntimeError("Failed to generate a unique sheet name.")


def fileformat_for_path(path: str) -> int:
    # Excel constants (avoid importing constants to keep pywin32 dynamic)
    ext = Path(path).suffix.lower()
    if ext == ".xlsx":
        return 51  # xlOpenXMLWorkbook
    if ext == ".xlsm":
        return 52  # xlOpenXMLWorkbookMacroEnabled
    if ext == ".xlsb":
        return 50  # xlExcel12
    if ext == ".xls":
        return 56  # xlExcel8
    return 51


@dataclass(frozen=True)
class SaveRequest:
    suggested_name: str


class ExcelMergeWorker(threading.Thread):
    def __init__(
        self,
        folder: Path,
        files: list[Path],
        ui_queue: "queue.Queue[tuple[str, object]]",
        save_queue: "queue.Queue[Optional[str]]",
        cancel_event: threading.Event,
    ) -> None:
        super().__init__(daemon=True)
        self.folder = folder
        self.files = files
        self.ui_queue = ui_queue
        self.save_queue = save_queue
        self.cancel_event = cancel_event

    def log(self, msg: str) -> None:
        self.ui_queue.put(("log", msg))

    def progress(self, current: int, total: int, filename: str) -> None:
        self.ui_queue.put(("progress", (current, total, filename)))

    def request_save_path(self, suggested_name: str) -> Optional[str]:
        self.ui_queue.put(("request_save", SaveRequest(suggested_name=suggested_name)))
        return self.save_queue.get()

    def run(self) -> None:
        try:
            self._run_impl()
        except Exception as e:  # noqa: BLE001
            self.log(f"[ERROR] {e}")
            self.ui_queue.put(("error", e))
            self.ui_queue.put(("done", "error"))

    def _run_impl(self) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception:
            self._run_impl_powershell()
            return

        pythoncom.CoInitialize()
        excel = None
        dest_wb = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            try:
                excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            try:
                excel.Calculation = -4135  # xlCalculationManual
            except Exception:
                pass

            dest_wb = excel.Workbooks.Add()
            initial_sheets = []
            try:
                for i in range(1, dest_wb.Worksheets.Count + 1):
                    initial_sheets.append(dest_wb.Worksheets(i))
            except Exception:
                initial_sheets = []

            total_files = len(self.files)
            self.log(f"対象ファイル数: {total_files}")
            copied_any = False

            for idx, src_path in enumerate(self.files, start=1):
                if self.cancel_event.is_set():
                    self.log("キャンセルされました。後処理中...")
                    break

                self.progress(idx, total_files, src_path.name)
                self.log(f"開く: {src_path.name}")

                src_wb = None
                try:
                    src_wb = excel.Workbooks.Open(
                        str(src_path),
                        ReadOnly=True,
                        UpdateLinks=0,
                        AddToMru=False,
                    )

                    for sidx in range(1, src_wb.Worksheets.Count + 1):
                        if self.cancel_event.is_set():
                            break

                        src_ws = src_wb.Worksheets(sidx)
                        desired = str(src_ws.Name)

                        existing = set()
                        for j in range(1, dest_wb.Worksheets.Count + 1):
                            existing.add(str(dest_wb.Worksheets(j).Name))

                        dest_ws_count = dest_wb.Worksheets.Count
                        src_ws.Copy(After=dest_wb.Worksheets(dest_ws_count))
                        copied = excel.ActiveSheet

                        unique = make_unique_sheet_name(existing_names=existing, desired=desired)
                        try:
                            copied.Name = unique
                        except Exception:
                            # If rename fails (rare), try with a sanitized fallback.
                            fallback = make_unique_sheet_name(existing_names=existing, desired=f"{desired}_copy")
                            copied.Name = fallback

                        copied_any = True
                        self.log(f"  コピー: {desired} -> {copied.Name}")

                except Exception as e:  # noqa: BLE001
                    self.log(f"[WARN] {src_path.name} を処理できません: {e}")
                finally:
                    try:
                        if src_wb is not None:
                            src_wb.Close(SaveChanges=False)
                    except Exception:
                        pass

            if self.cancel_event.is_set():
                if dest_wb is not None:
                    dest_wb.Close(SaveChanges=False)
                excel.Quit()
                self.ui_queue.put(("done", "cancelled"))
                return

            # Remove initial blank sheets after at least one copy.
            if copied_any and initial_sheets:
                try:
                    dest_wb.Worksheets(1).Activate()
                except Exception:
                    pass
                for sh in initial_sheets:
                    try:
                        if dest_wb.Worksheets.Count <= 1:
                            break
                        sh.Delete()
                    except Exception:
                        pass

            self.log("統合完了。保存先を選択してください。")
            suggested = safe_default_filename(self.folder)

            while True:
                save_path = self.request_save_path(suggested_name=suggested)
                if save_path is None:
                    self.log("保存がキャンセルされました。保存せずに終了します。")
                    dest_wb.Close(SaveChanges=False)
                    excel.Quit()
                    self.ui_queue.put(("done", "nosave"))
                    return

                fmt = fileformat_for_path(save_path)
                try:
                    dest_wb.SaveAs(save_path, FileFormat=fmt)
                    self.log(f"保存しました: {save_path}")
                    break
                except Exception as e:  # noqa: BLE001
                    self.log(f"[ERROR] 保存に失敗しました: {e}")
                    # Ask UI again by looping.

            dest_wb.Close(SaveChanges=False)
            excel.Quit()
            self.ui_queue.put(("done", "saved"))

        finally:
            try:
                if dest_wb is not None:
                    dest_wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            try:
                import pythoncom  # type: ignore

                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _run_impl_powershell(self) -> None:
        script_path = app_base_dir() / "merge_excel_sheets.ps1"
        if not script_path.exists():
            raise RuntimeError(f"必要なスクリプトが見つかりません: {script_path}")

        local_app_data = os.environ.get("LOCALAPPDATA")
        base_dir = Path(local_app_data) if local_app_data else Path.home()
        tmp_dir = base_dir / "ExcelMerger" / "tmp"
        tmp_dir.mkdir(parents=True, exist_ok=True)

        file_list_path = tmp_dir / f"filelist_{int(time.time())}.txt"
        cancel_flag_path = tmp_dir / f"cancel_{int(time.time())}.flag"

        file_list_path.write_text(
            "\n".join(str(p) for p in self.files),
            encoding="utf-8",
        )

        total_files = len(self.files)
        self.log(f"(PowerShell) 対象ファイル数: {total_files}")

        cmd = [
            "powershell",
            "-NoProfile",
            "-NoLogo",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            "-FileListPath",
            str(file_list_path),
            "-CancelFlagPath",
            str(cancel_flag_path),
            "-SuggestedName",
            safe_default_filename(self.folder),
        ]

        proc = subprocess.Popen(  # noqa: S603
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        try:
            assert proc.stdout is not None
            assert proc.stdin is not None

            while True:
                if self.cancel_event.is_set() and not cancel_flag_path.exists():
                    try:
                        cancel_flag_path.write_text("1", encoding="utf-8")
                    except Exception:
                        pass

                line = proc.stdout.readline()
                if line == "" and proc.poll() is not None:
                    break

                line = line.strip()
                if not line:
                    continue

                if line.startswith("LOG|"):
                    self.log(line[4:])
                    continue

                if line.startswith("PROGRESS|"):
                    parts = line.split("|", 3)
                    if len(parts) == 4:
                        try:
                            current = int(parts[1])
                            total = int(parts[2])
                        except ValueError:
                            continue
                        filename = parts[3]
                        self.progress(current, total, filename)
                    continue

                if line.startswith("REQUEST_SAVE|"):
                    suggested = line.split("|", 1)[1] if "|" in line else safe_default_filename(self.folder)
                    save_path = self.request_save_path(suggested_name=suggested)
                    proc.stdin.write((save_path or "") + "\n")
                    proc.stdin.flush()
                    continue

                self.log(line)

            rc = proc.wait()
            if self.cancel_event.is_set():
                self.ui_queue.put(("done", "cancelled"))
                return

            if rc == 0:
                self.ui_queue.put(("done", "saved"))
            elif rc == 3:
                self.ui_queue.put(("done", "nosave"))
            elif rc == 2:
                self.ui_queue.put(("done", "cancelled"))
            else:
                raise RuntimeError(f"PowerShell 統合処理が失敗しました (exit={rc})")

        finally:
            try:
                if proc.poll() is None:
                    proc.terminate()
            except Exception:
                pass
            try:
                if file_list_path.exists():
                    file_list_path.unlink()
            except Exception:
                pass
            try:
                if cancel_flag_path.exists():
                    cancel_flag_path.unlink()
            except Exception:
                pass


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 統合ツール（シートコピー）")
        self.geometry("820x520")

        self.ui_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.save_queue: "queue.Queue[Optional[str]]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: Optional[ExcelMergeWorker] = None

        self.folder_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="フォルダを選択してください。")
        self.selected_only_var = tk.BooleanVar(value=False)
        self.item_to_path: dict[str, Path] = {}

        self._build_ui()
        self.after(100, self._poll_ui_queue)

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="対象フォルダ:").pack(side=tk.LEFT)
        entry = ttk.Entry(top, textvariable=self.folder_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        ttk.Button(top, text="参照...", command=self._browse_folder).pack(side=tk.LEFT)
        ttk.Button(top, text="一覧更新", command=self._refresh_list).pack(side=tk.LEFT, padx=(8, 0))

        mid = ttk.Frame(self, padding=(10, 0, 10, 10))
        mid.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(
            mid,
            columns=("name", "size"),
            show="headings",
            height=12,
            selectmode="extended",
        )
        self.tree.heading("name", text="ファイル名")
        self.tree.heading("size", text="サイズ")
        self.tree.column("name", width=600, anchor=tk.W)
        self.tree.column("size", width=120, anchor=tk.E)

        vsb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        mid.grid_rowconfigure(0, weight=1)
        mid.grid_columnconfigure(0, weight=1)

        ctl = ttk.Frame(self, padding=(10, 0, 10, 10))
        ctl.pack(fill=tk.X)

        self.start_btn = ttk.Button(ctl, text="統合開始", command=self._start_merge, state=tk.DISABLED)
        self.cancel_btn = ttk.Button(ctl, text="キャンセル", command=self._cancel_merge, state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT)
        self.cancel_btn.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Checkbutton(
            ctl,
            text="選択したファイルのみ",
            variable=self.selected_only_var,
        ).pack(side=tk.LEFT, padx=(12, 0))

        self.pbar = ttk.Progressbar(ctl, mode="determinate")
        self.pbar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(12, 0))

        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.pack(fill=tk.BOTH, expand=False)

        ttk.Label(bottom, textvariable=self.status_var).pack(anchor=tk.W)

        self.log_text = tk.Text(bottom, height=8, wrap="word")
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.log_text.configure(state=tk.DISABLED)

    def _browse_folder(self) -> None:
        folder = filedialog.askdirectory(title="対象フォルダを選択")
        if folder:
            self.folder_var.set(folder)
            self._refresh_list()

    def _refresh_list(self) -> None:
        folder_str = self.folder_var.get().strip()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.item_to_path.clear()

        if not folder_str:
            self.status_var.set("フォルダを選択してください。")
            self.start_btn.configure(state=tk.DISABLED)
            return

        folder = Path(folder_str)
        if not folder.exists():
            self.status_var.set("フォルダが存在しません。")
            self.start_btn.configure(state=tk.DISABLED)
            return

        files = sorted([p for p in folder.iterdir() if is_excel_file(p)], key=lambda p: p.name.lower())
        for p in files:
            size = f"{p.stat().st_size:,}"
            iid = self.tree.insert("", tk.END, values=(p.name, size))
            self.item_to_path[str(iid)] = p

        self.status_var.set(f"見つかった Excel ファイル: {len(files)} 件")
        self.start_btn.configure(state=tk.NORMAL if files else tk.DISABLED)

    def _start_merge(self) -> None:
        if self.worker is not None:
            return

        folder = Path(self.folder_var.get().strip())
        all_files = sorted([p for p in folder.iterdir() if is_excel_file(p)], key=lambda p: p.name.lower())
        if self.selected_only_var.get():
            selected_iids = list(self.tree.selection())
            files = [self.item_to_path[iid] for iid in selected_iids if iid in self.item_to_path]
            files = sorted(files, key=lambda p: p.name.lower())
        else:
            files = all_files

        if not files:
            if self.selected_only_var.get():
                messagebox.showinfo("情報", "ファイルが選択されていません。一覧から対象を選択してください。")
            else:
                messagebox.showinfo("情報", "統合対象の Excel ファイルが見つかりません。")
            return

        self._log("=== 統合開始 ===")
        self.cancel_event.clear()
        self.worker = ExcelMergeWorker(
            folder=folder,
            files=files,
            ui_queue=self.ui_queue,
            save_queue=self.save_queue,
            cancel_event=self.cancel_event,
        )
        self._set_busy(True, total=len(files))
        self.worker.start()

    def _cancel_merge(self) -> None:
        if self.worker is None:
            return
        self.cancel_event.set()
        self._log("キャンセル要求を送信しました。")

    def _set_busy(self, busy: bool, total: int = 0) -> None:
        self.start_btn.configure(state=tk.DISABLED if busy else tk.NORMAL)
        self.cancel_btn.configure(state=tk.NORMAL if busy else tk.DISABLED)
        self.pbar.configure(maximum=max(total, 1), value=0)

    def _log(self, msg: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _poll_ui_queue(self) -> None:
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "progress":
                    current, total, filename = payload  # type: ignore[misc]
                    self.pbar.configure(maximum=max(int(total), 1), value=int(current))
                    self.status_var.set(f"{current}/{total}: {filename}")
                elif kind == "request_save":
                    req = payload  # type: ignore[assignment]
                    self._handle_save_request(req)
                elif kind == "done":
                    self._handle_done(str(payload))
                elif kind == "error":
                    # keep going; done should come after cleanup
                    pass
        except queue.Empty:
            pass
        finally:
            self.after(120, self._poll_ui_queue)

    def _handle_save_request(self, req: SaveRequest) -> None:
        folder = Path(self.folder_var.get().strip())
        initial = req.suggested_name

        while True:
            path = filedialog.asksaveasfilename(
                title="名前を付けて保存",
                initialdir=str(folder),
                initialfile=initial,
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel Workbook (*.xlsx)", "*.xlsx"),
                    ("Excel Macro-Enabled Workbook (*.xlsm)", "*.xlsm"),
                    ("Excel 97-2003 Workbook (*.xls)", "*.xls"),
                    ("Excel Binary Workbook (*.xlsb)", "*.xlsb"),
                ],
            )
            if path:
                normalized = os.path.normpath(path)
                if Path(normalized).suffix == "":
                    normalized += ".xlsx"
                parent = Path(normalized).parent
                if not parent.exists():
                    messagebox.showerror("保存エラー", f"保存先フォルダが存在しません:\n{parent}")
                    continue
                self.save_queue.put(normalized)
                return

            retry = messagebox.askyesno("保存キャンセル", "保存がキャンセルされました。再度保存先を選択しますか？")
            if not retry:
                # Signal abort to worker.
                self.save_queue.put(None)
                return

    def _handle_done(self, status: str) -> None:
        self.worker = None
        self._set_busy(False)
        self._refresh_list()

        if status == "saved":
            self.status_var.set("完了: 保存しました。")
            self._log("=== 完了（保存済み） ===")
        elif status == "cancelled":
            self.status_var.set("中断: キャンセルされました。")
            self._log("=== 中断（キャンセル） ===")
        elif status == "nosave":
            self.status_var.set("完了: 保存せず終了しました。")
            self._log("=== 完了（未保存） ===")
        elif status == "error":
            self.status_var.set("エラー: ログを確認してください。")
            self._log("=== エラー ===")
            messagebox.showerror("エラー", "処理中にエラーが発生しました。ログを確認してください。")
        else:
            self.status_var.set("完了")
            self._log("=== 完了 ===")


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
