"""勤務表画像・PDF 解析ツールの Windows GUI エントリ。"""

from __future__ import annotations

import json
import os
import re
import sys
import threading
import ctypes
from ctypes import wintypes
from datetime import date
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

_CODE_DIR = Path(__file__).resolve().parent
_ROOT = _CODE_DIR.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from kintai_core import (
    DEFAULT_PARALLEL_ANALYSIS_CHATS,
    EXCEL_SUFFIXES,
    LEGACY_COMPANY_COL,
    LEGACY_COMPANY_NAME_COL,
    LEGACY_COMPANY_READ_LONG_COL,
    LEGACY_FILE_NAME_COL,
    LEGACY_EMPLOYEE_NO_COL,
    LEGACY_MONTH_COL,
    LEGACY_PERSON_COL,
    LEGACY_TARGET_FILE_NAME_COL,
    LEGACY_YEAR_COL,
    PARALLEL_WORKERS_MAX,
    LEGACY_USER_JUDGMENT_COL,
    LEGACY_BILLING_UPDATE_HOURS_COL,
    SUMMARY_BILLING_UPDATE_RESULT_COL,
    SUMMARY_BILLING_UPDATE_HOURS_COL,
    SUMMARY_BILLING_UPDATE_TRANSPORT_COL,
    SUMMARY_COMPANY_COL,
    SUMMARY_EMPLOYEE_NO_COL,
    SUMMARY_FINAL_JUDGMENT_COL,
    SUMMARY_MONTH_COL,
    SUMMARY_PERSON_COL,
    SUMMARY_ROW_NO_COL,
    SUMMARY_YEAR_COL,
    TARGET_ASSISTANT_NAME,
    TARGET_FILE_NAME_COL,
    _decimal_for_table_display,
    _normalize_month_value,
    _normalize_year_value,
    _work_hours_string_to_decimal,
    auto_judgment_symbol,
    create_client,
    is_manual_user_judgment,
    normalize_judgment_symbol,
    populate_billing_update_columns,
    row_display_values,
    run_analysis,
    summary_header_cells,
    update_billing_engineer_ts_sheet,
    _row_billing_update_hours_decimal,
)
from newtonx_adk.exceptions import APIError

_WIN32 = sys.platform == "win32"
_STILL_ACTIVE = 259
_YM_DIR_RE = re.compile(r"^(\d{4})年(\d{1,2})月$")


def _today_year_month() -> tuple[int, int]:
    """本日の年・月（データフォルダ名 YYYY年M月 と同じ暦年）。"""
    today = date.today()
    return today.year, today.month


def _parse_year_month_dir_name(name: str) -> tuple[int, int] | None:
    m = _YM_DIR_RE.fullmatch((name or "").strip())
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def _parse_data_dir_path(
    path: Path,
) -> tuple[Path | None, int | None, int | None, str]:
    """パスから Data ルート・年・月・年月下のサブフォルダ（支社名等）を推定。"""
    p = path.resolve()
    parts = p.parts
    for i, part in enumerate(parts):
        ym = _parse_year_month_dir_name(part)
        if ym is None:
            continue
        year, month = ym
        root = Path(*parts[:i]) if i > 0 else None
        ym_path = Path(*parts[: i + 1])
        try:
            rel = p.relative_to(ym_path)
            branch = "" if str(rel) == "." else str(rel)
        except ValueError:
            branch = p.name
        return root, year, month, branch
    return None, None, None, ""


def _guess_data_root() -> Path | None:
    for cand in (_ROOT.parent / "Data", _ROOT / "Data"):
        if cand.is_dir():
            return cand.resolve()
    return None


class _ShellExecuteInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("fMask", wintypes.ULONG),
        ("hwnd", wintypes.HWND),
        ("lpVerb", wintypes.LPCWSTR),
        ("lpFile", wintypes.LPCWSTR),
        ("lpParameters", wintypes.LPCWSTR),
        ("lpDirectory", wintypes.LPCWSTR),
        ("nShow", ctypes.c_int),
        ("hInstApp", wintypes.HMODULE),
        ("lpIDList", ctypes.c_void_p),
        ("lpClass", wintypes.LPCWSTR),
        ("hkeyClass", wintypes.HKEY),
        ("dwHotKey", wintypes.DWORD),
        ("hMonitor", wintypes.HANDLE),
        ("hProcess", wintypes.HANDLE),
    ]


def _win_shell_open_file(path: Path) -> int | None:
    """既定アプリでファイルを開き、取得できればプロセスハンドルを返す。"""
    if not _WIN32:
        os.startfile(os.fspath(path))
        return None
    sei = _ShellExecuteInfo()
    sei.cbSize = ctypes.sizeof(_ShellExecuteInfo)
    sei.fMask = 0x00000040  # SEE_MASK_NOCLOSEPROCESS
    sei.lpVerb = "open"
    sei.lpFile = os.fspath(path.resolve())
    sei.nShow = 1  # SW_SHOWNORMAL
    if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(sei)):
        raise OSError(ctypes.GetLastError(), "ShellExecuteExW", str(path))
    return int(sei.hProcess or 0) or None


def _win_terminate_process_handle(handle: int | None) -> None:
    if not handle:
        return
    kernel32 = ctypes.windll.kernel32
    exit_code = wintypes.DWORD()
    if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
        if exit_code.value == _STILL_ACTIVE:
            kernel32.TerminateProcess(handle, 0)
    kernel32.CloseHandle(handle)


def _try_close_excel_workbook(path: Path) -> bool:
    """実行中の Excel から同一ファイルのブックだけを閉じる（pywin32 がある場合）。"""
    try:
        import win32com.client  # type: ignore[import-untyped]
    except ImportError:
        return False
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return False
    target = str(path.resolve()).lower()
    closed = False
    try:
        for i in range(int(excel.Workbooks.Count), 0, -1):
            wb = excel.Workbooks(i)
            try:
                full = str(Path(wb.FullName).resolve()).lower()
            except Exception:
                continue
            if full == target:
                wb.Close(SaveChanges=0)
                closed = True
    except Exception:
        return closed
    return closed


class KintaiApp(tk.Frame):
    TARGET_FILE_NAME_COL = TARGET_FILE_NAME_COL
    LEGACY_TARGET_FILE_NAME_COL = LEGACY_TARGET_FILE_NAME_COL
    LEGACY_FILE_NAME_COL = LEGACY_FILE_NAME_COL
    FINAL_JUDGMENT_COL = SUMMARY_FINAL_JUDGMENT_COL
    LEGACY_USER_JUDGMENT_COL = LEGACY_USER_JUDGMENT_COL
    BILLING_UPDATE_RESULT_COL = SUMMARY_BILLING_UPDATE_RESULT_COL
    AUTO_JUDGMENT_COL = "自動判断"
    YEAR_COL = SUMMARY_YEAR_COL
    MONTH_COL = SUMMARY_MONTH_COL
    EMPLOYEE_NO_COL = SUMMARY_EMPLOYEE_NO_COL
    ROW_NO_COL = SUMMARY_ROW_NO_COL
    BILLING_UPDATE_HOURS_COL = SUMMARY_BILLING_UPDATE_HOURS_COL
    BILLING_UPDATE_TRANSPORT_COL = SUMMARY_BILLING_UPDATE_TRANSPORT_COL
    PERSON_COL = SUMMARY_PERSON_COL
    TOTAL_HOURS_DECIMAL_COL = "合計勤務時間（10進）"
    TOTAL_HOURS_RAW_COL = "合計勤務時間（読取）"
    TRANSPORT_EXPENSE_COL = "交通費合計（読取）"
    MATCH_COMPANY_COL = "会社名比較"
    LEGACY_MATCH_COMPANY_COL = "会社名比較（ファイル名✖文書）"
    COMPANY_COL = SUMMARY_COMPANY_COL
    LEGACY_COMPANY_COL = LEGACY_COMPANY_COL
    LEGACY_COMPANY_NAME_COL = LEGACY_COMPANY_NAME_COL
    LEGACY_COMPANY_READ_LONG_COL = LEGACY_COMPANY_READ_LONG_COL
    LEGACY_YEAR_COL = LEGACY_YEAR_COL
    LEGACY_MONTH_COL = LEGACY_MONTH_COL
    LEGACY_PERSON_COL = LEGACY_PERSON_COL
    LEGACY_EMPLOYEE_NO_COL = LEGACY_EMPLOYEE_NO_COL
    # 「エラー再解析」対象: 最終判断が「〇」以外の行
    _ERROR_REANALYSIS_OK_VALUES = ("〇",)
    # 記号列（〇/△/✖ 等）の列幅（px）
    _SYMBOL_COLUMN_WIDTHS: dict[str, int] = {
        ROW_NO_COL: 44,
        "アップロード": 72,
        "対象シート有無": 88,
        YEAR_COL: 56,
        MONTH_COL: 56,
        FINAL_JUDGMENT_COL: 72,
        BILLING_UPDATE_RESULT_COL: 120,
        AUTO_JUDGMENT_COL: 72,
        MATCH_COMPANY_COL: 72,
        "押印有無": 72,
    }
    _TAG_REANALYSIS_ACTIVE = "reanalysis_active"

    def __init__(
        self,
        master: tk.Tk,
        *,
        client,
        assistants: list[dict],
    ) -> None:
        super().__init__(master)
        self._root = master
        self._root.title("勤務表解析")
        self._root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._client = client
        self._assistants = list(assistants)
        assistant_names = [
            str(a.get("name") or "").strip()
            for a in self._assistants
            if str(a.get("name") or "").strip()
        ]
        default_name = (
            TARGET_ASSISTANT_NAME
            if TARGET_ASSISTANT_NAME in assistant_names
            else (assistant_names[0] if assistant_names else "")
        )
        self._assistant_var = tk.StringVar(value=default_name)
        self._workers_var = tk.StringVar(value=str(DEFAULT_PARALLEL_ANALYSIS_CHATS))
        self._data_dir: Path | None = None
        self._billing_file_path: Path | None = None
        self._data_root: Path | None = _guess_data_root()
        self._data_branch: str = ""
        default_year, _ = _today_year_month()
        self._year_var = tk.StringVar(value=str(default_year))
        self._month_var = tk.StringVar(value="")
        self._busy = False
        self._item_paths: dict[str, str] = {}
        self._preview_row_iid: str = ""
        self._preview_file_path: Path | None = None
        self._preview_process_handle: int | None = None
        self._cancel_event: threading.Event | None = None

        self._loaded_rows: list[dict[str, str]] = []
        self._loaded_json_path: Path | None = None
        self._last_saved_snapshot: str = ""
        self._chain_error_reanalysis_after_new = False
        self._sort_column: str | None = None
        self._sort_reverse: bool = False
        self._tree_column_headings: tuple[str, ...] = ()

        self._build_ui()

    def _tree_heading_font(self) -> tkfont.Font:
        spec = ttk.Style().lookup("Treeview.Heading", "font")
        if spec:
            return tkfont.Font(root=self._root, font=spec)
        return tkfont.nametofont("TkDefaultFont")

    def _column_width_for_heading(
        self, heading: str, *, font: tkfont.Font | None = None
    ) -> int:
        """見出し文字列の描画幅に合わせた列幅（px）を返す。"""
        if heading in self._SYMBOL_COLUMN_WIDTHS:
            return self._SYMBOL_COLUMN_WIDTHS[heading]
        f = font or self._tree_heading_font()
        padding_px = 20
        min_px = 48
        return max(f.measure(heading) + padding_px, min_px)

    def _window_width_for_columns(
        self, col_px: list[int], *, y_scroll: ttk.Scrollbar
    ) -> int:
        """全列が欠けずに見えるよう、ウィンドウ幅（px）を算出する。"""
        self.update_idletasks()
        scrollbar_w = y_scroll.winfo_reqwidth()
        if scrollbar_w <= 1:
            scrollbar_w = 18
        grid_pad_x = 16  # grid_frame padding (8, 0, 8, 8)
        chrome_x = 16
        w = sum(col_px) + scrollbar_w + grid_pad_x + chrome_x
        return min(w, self._root.winfo_screenwidth())

    def _parallel_workers_value(self) -> int:
        try:
            nw = int(str(self._workers_var.get()).strip())
        except (TypeError, ValueError):
            nw = DEFAULT_PARALLEL_ANALYSIS_CHATS
        return max(1, min(nw, PARALLEL_WORKERS_MAX))

    def _assistant_uid_for_name(self, name: str) -> str:
        selected_name = (name or "").strip()
        if not selected_name:
            return ""
        for a in self._assistants:
            if str(a.get("name") or "").strip() != selected_name:
                continue
            raw = a.get("uid")
            if raw is None or str(raw).strip() == "":
                raw = a.get("uuid")
            return str(raw).strip() if raw is not None else ""
        return ""

    def _require_assistant_uid(self) -> str | None:
        name = (self._assistant_var.get() or "").strip()
        if not name:
            messagebox.showwarning(
                "アシスタント未選択",
                "使用するアシスタントを選択してください。",
                parent=self._root,
            )
            return None
        uid = self._assistant_uid_for_name(name)
        if not uid:
            messagebox.showerror(
                "エラー",
                f"アシスタント「{name}」の ID を取得できません。",
                parent=self._root,
            )
            return None
        return uid

    def _build_ui(self) -> None:
        cfg = ttk.Frame(self, padding=(8, 8, 8, 0))
        cfg.pack(fill=tk.X)
        ttk.Label(cfg, text="アシスタント:").grid(row=0, column=0, sticky="w")
        assistant_names = [
            str(a.get("name") or "").strip()
            for a in self._assistants
            if str(a.get("name") or "").strip()
        ]
        self._assistant_combo = ttk.Combobox(
            cfg,
            textvariable=self._assistant_var,
            values=assistant_names,
            state="readonly",
            width=36,
        )
        self._assistant_combo.grid(row=0, column=1, sticky="w", padx=(4, 16))
        ttk.Label(cfg, text="並列数:").grid(row=0, column=2, sticky="w")
        tk.Spinbox(
            cfg,
            from_=1,
            to=PARALLEL_WORKERS_MAX,
            textvariable=self._workers_var,
            width=6,
            justify="center",
        ).grid(row=0, column=3, sticky="w", padx=(4, 8))
        ttk.Label(
            cfg,
            text="（1～4推奨、少ないほうが安定）",
        ).grid(row=0, column=4, sticky="w")

        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        self._folder_var = tk.StringVar(value="（未選択）")
        ttk.Label(top, text="データフォルダ:").grid(row=0, column=0, sticky="w")
        ttk.Button(top, text="参照…", command=self._browse_folder).grid(
            row=0, column=1, sticky="w", padx=(4, 4)
        )
        ttk.Label(top, textvariable=self._folder_var, anchor="w").grid(
            row=0, column=2, sticky="ew", padx=(0, 0)
        )
        top.columnconfigure(2, weight=1)

        self._billing_file_var = tk.StringVar(value="（未選択）")
        ttk.Label(top, text="請求用ファイル:").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Button(
            top, text="参照…", command=self._browse_billing_file
        ).grid(row=1, column=1, sticky="w", padx=(4, 4), pady=(6, 0))
        ttk.Label(top, textvariable=self._billing_file_var, anchor="w").grid(
            row=1, column=2, sticky="ew", pady=(6, 0)
        )

        ym_row = ttk.Frame(top)
        ym_row.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(ym_row, text="年度:").pack(side=tk.LEFT)
        self._year_combo = ttk.Combobox(
            ym_row,
            textvariable=self._year_var,
            width=8,
            state="readonly",
        )
        self._year_combo.pack(side=tk.LEFT, padx=(4, 12))
        self._year_combo.bind("<<ComboboxSelected>>", self._on_year_month_changed)

        ttk.Label(ym_row, text="月:").pack(side=tk.LEFT)
        self._month_combo = ttk.Combobox(
            ym_row,
            textvariable=self._month_var,
            width=4,
            state="readonly",
        )
        self._month_combo.pack(side=tk.LEFT, padx=(4, 0))
        self._month_combo.bind("<<ComboboxSelected>>", self._on_year_month_changed)

        self._refresh_year_month_combos()

        ctrl = ttk.Frame(self, padding=(8, 0, 8, 8))
        ctrl.pack(fill=tk.X)

        self._new_btn = ttk.Button(
            ctrl, text="新規解析", command=self._start_new_analysis, state=tk.DISABLED
        )
        self._new_btn.grid(row=0, column=0, sticky="w")

        self._new_plus_error_btn = ttk.Button(
            ctrl,
            text="新規解析＋エラー再解析",
            command=self._start_new_analysis_then_error_reanalysis,
            state=tk.DISABLED,
        )
        self._new_plus_error_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self._cont_btn = ttk.Button(
            ctrl, text="継続解析", command=self._start_continue_analysis, state=tk.DISABLED
        )
        self._cont_btn.grid(row=0, column=2, sticky="w", padx=(8, 0))

        self._selected_reanalysis_btn = ttk.Button(
            ctrl,
            text="選択行解析",
            command=self._start_selected_rows_reanalysis,
            state=tk.DISABLED,
        )
        self._selected_reanalysis_btn.grid(row=0, column=3, sticky="w", padx=(8, 0))

        self._error_reanalysis_btn = ttk.Button(
            ctrl,
            text="エラー再解析",
            command=self._start_error_reanalysis,
            state=tk.DISABLED,
        )
        self._error_reanalysis_btn.grid(row=0, column=4, sticky="w", padx=(8, 0))

        self._billing_prepare_btn = ttk.Button(
            ctrl,
            text="請求データ作成",
            command=self._create_billing_data,
            state=tk.DISABLED,
        )
        self._billing_prepare_btn.grid(row=0, column=5, sticky="w", padx=(8, 0))

        self._billing_update_btn = ttk.Button(
            ctrl,
            text="請求ファイル更新",
            command=self._update_billing_file,
            state=tk.DISABLED,
        )
        self._billing_update_btn.grid(row=0, column=6, sticky="w", padx=(8, 0))

        self._save_btn = ttk.Button(
            ctrl, text="保存", command=self._save_json, state=tk.DISABLED
        )
        self._save_btn.grid(row=0, column=7, sticky="w", padx=(8, 0))

        self._load_btn = ttk.Button(
            ctrl, text="読み込み", command=self._load_json, state=tk.NORMAL
        )
        self._load_btn.grid(row=0, column=8, sticky="w", padx=(8, 0))

        self._cancel_btn = ttk.Button(
            ctrl, text="中断", command=self._cancel_analysis, state=tk.DISABLED
        )
        self._cancel_btn.grid(row=0, column=9, sticky="w", padx=(16, 0))

        self._progress_var = tk.StringVar(value="")
        self._status_var = tk.StringVar(value="準備完了")
        # 「実行済 100 / 対象 120」など3桁になっても欠けないよう、表示幅を広げる
        # ttk.Label の width は“文字数”ベースなので、minsize と合わせて余裕を持たせる。
        # （環境によってフォントが少し太く、26文字だと末尾が欠けるケースがあったため更に増やす）
        ttk.Label(ctrl, textvariable=self._progress_var, width=30).grid(
            row=0, column=10, sticky="w", padx=(16, 0)
        )
        # 進捗表示（実行済/対象）は桁数により伸びるため、最低幅を確保して欠けを防ぐ
        ctrl.columnconfigure(10, minsize=240)

        self._status_label = ttk.Label(
            ctrl,
            textvariable=self._status_var,
            anchor="w",
            width=70,
            wraplength=900,
            justify="left",
        )
        self._status_label.grid(row=0, column=11, sticky="ew", padx=(12, 0))
        ctrl.columnconfigure(11, weight=1)

        grid_frame = ttk.Frame(self, padding=(8, 0, 8, 8))
        grid_frame.pack(fill=tk.BOTH, expand=True)

        headings = summary_header_cells()
        self._tree_column_headings = headings
        y_scroll = ttk.Scrollbar(grid_frame)
        x_scroll = ttk.Scrollbar(grid_frame, orient=tk.HORIZONTAL)

        heading_font = self._tree_heading_font()
        col_px = [
            self._column_width_for_heading(h, font=heading_font) for h in headings
        ]

        self._tree = ttk.Treeview(
            grid_frame,
            columns=list(headings),
            show="headings",
            selectmode="extended",
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set,
        )

        # 再解析中の行を反転表示（背景/文字色）
        # OSテーマにより見え方が変わるため、強めのコントラストにする。
        self._tree.tag_configure(self._TAG_REANALYSIS_ACTIVE, background="#1f2937", foreground="#ffffff")
        y_scroll.configure(command=self._tree.yview)
        x_scroll.configure(command=self._tree.xview)

        for w_px, h in zip(col_px, headings, strict=True):
            self._tree.column(
                h, width=w_px, minwidth=48, stretch=tk.NO, anchor="w"
            )
            self._tree.heading(
                h,
                text=h,
                anchor="w",
                command=lambda col=h: self._sort_grid_by_column(col),
            )
        self._update_column_heading_labels()

        self._tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        grid_frame.rowconfigure(0, weight=1)
        grid_frame.columnconfigure(0, weight=1)

        self._tree.bind("<Button-3>", self._on_tree_right_click)
        self._tree.bind("<Button-1>", self._on_tree_left_click, add="+")
        self._tree.bind("<Double-1>", self._on_row_double_click)
        self._tree.bind("<<TreeviewSelect>>", self._on_tree_selection_changed)

        initial_w = self._window_width_for_columns(col_px, y_scroll=y_scroll)
        min_w = min(960, initial_w)
        self._root.minsize(min_w, 520)
        self._root.geometry(f"{initial_w}x680")
        self._sync_folder_display()
        self._sync_billing_file_display()
        self._refresh_reanalysis_buttons_state()

    def _sync_folder_display(self) -> None:
        """データフォルダ欄を _data_dir の状態に合わせる（未設定時は未選択）。"""
        if self._data_dir is not None:
            self._folder_var.set(str(self._data_dir.resolve()))
        else:
            self._folder_var.set("（未選択）")

    def _sync_billing_file_display(self) -> None:
        """請求用ファイル欄を _billing_file_path の状態に合わせる。"""
        if self._billing_file_path is not None and self._billing_file_path.is_file():
            self._billing_file_var.set(self._billing_file_path.name)
        else:
            self._billing_file_var.set("（未選択）")

    def _discover_years(self) -> list[str]:
        today_y, _ = _today_year_month()
        years: set[int] = {today_y, today_y - 1, today_y + 1}
        if self._data_root is not None and self._data_root.is_dir():
            for child in self._data_root.iterdir():
                if not child.is_dir():
                    continue
                ym = _parse_year_month_dir_name(child.name)
                if ym is not None:
                    years.add(ym[0])
        return [str(y) for y in sorted(years, reverse=True)]

    def _discover_months(self, year: int) -> list[str]:
        months: set[int] = set(range(1, 13))
        if self._data_root is not None and self._data_root.is_dir():
            for child in self._data_root.iterdir():
                if not child.is_dir():
                    continue
                ym = _parse_year_month_dir_name(child.name)
                if ym is not None and ym[0] == year:
                    months.add(ym[1])
        return [str(m) for m in sorted(months)]

    def _refresh_year_month_combos(self) -> None:
        years = self._discover_years()
        if years:
            self._year_combo.configure(values=years)
            if self._year_var.get() not in years:
                self._year_var.set(years[0])
        try:
            year = int(self._year_var.get().strip())
        except ValueError:
            year = _today_year_month()[0]
        months = self._discover_months(year)
        if months:
            self._month_combo.configure(values=months)
            cur_m = self._month_var.get().strip()
            if cur_m and cur_m not in months:
                self._month_var.set(months[-1])

    def _expected_year_month(self) -> tuple[int | None, int | None]:
        """画面上の年度・月コンボの値（run_analysis の照合用）。"""
        try:
            return (
                int(self._year_var.get().strip()),
                int(self._month_var.get().strip()),
            )
        except ValueError:
            return None, None

    def _analysis_prerequisites_met(self) -> bool:
        """データフォルダと請求用ファイルの両方が選択済みか。"""
        return (
            self._data_dir is not None
            and self._data_dir.is_dir()
            and self._billing_file_path is not None
            and self._billing_file_path.is_file()
        )

    def _update_data_dir_dependent_buttons(self) -> None:
        can_analyze = self._analysis_prerequisites_met()
        if self._busy:
            return
        self._progress_var.set("")
        self._new_btn.configure(state=(tk.NORMAL if can_analyze else tk.DISABLED))
        self._new_plus_error_btn.configure(
            state=(tk.NORMAL if can_analyze else tk.DISABLED)
        )
        self._cont_btn.configure(
            state=(
                tk.NORMAL
                if (can_analyze and self._loaded_rows)
                else tk.DISABLED
            )
        )
        self._save_btn.configure(
            state=(tk.NORMAL if self._tree.get_children() else tk.DISABLED)
        )
        self._cancel_btn.configure(state=tk.DISABLED)
        self._refresh_reanalysis_buttons_state()

    def _refresh_billing_buttons_state(self) -> None:
        if self._busy:
            self._billing_prepare_btn.configure(state=tk.DISABLED)
            self._billing_update_btn.configure(state=tk.DISABLED)
            return
        has_rows = bool(self._tree.get_children())
        self._billing_prepare_btn.configure(
            state=(tk.NORMAL if has_rows else tk.DISABLED)
        )
        enabled = (
            has_rows
            and self._billing_file_path is not None
            and self._billing_file_path.is_file()
        )
        self._billing_update_btn.configure(state=(tk.NORMAL if enabled else tk.DISABLED))

    def _refresh_billing_update_button_state(self) -> None:
        self._refresh_billing_buttons_state()

    def _on_tree_selection_changed(self, _event: tk.Event | None = None) -> None:
        self._refresh_selected_rows_reanalysis_button_state()

    def _on_year_month_changed(self, _event: tk.Event | None = None) -> None:
        self._refresh_year_month_combos()

    def _set_data_path(
        self, path: Path, *, sync_year_month_from_path: bool = False
    ) -> None:
        """解析対象フォルダを設定する。sync_year_month_from_path 時のみパスから年月コンボを更新。"""
        root, year, month, branch = _parse_data_dir_path(path)
        if root is not None:
            self._data_root = root
        if sync_year_month_from_path:
            if year is not None:
                self._year_var.set(str(year))
            if month is not None:
                self._month_var.set(str(month))
        self._data_branch = branch
        self._data_dir = path.resolve()
        self._sync_folder_display()
        self._refresh_year_month_combos()
        self._update_data_dir_dependent_buttons()

    def _browse_folder(self) -> None:
        self._prepare_native_dialog()
        initial = self._data_dir or self._data_root
        d = filedialog.askdirectory(
            title="データを読み込むフォルダを選択",
            parent=self._root,
            initialdir=str(initial) if initial else None,
        )
        if not d:
            return
        self._set_data_path(Path(d), sync_year_month_from_path=False)

    def _browse_billing_file(self) -> None:
        self._prepare_native_dialog()
        initial = self._billing_file_path or self._data_dir or self._data_root
        initialdir: str | None = None
        initialfile: str | None = None
        if initial is not None:
            if initial.is_file():
                initialdir = str(initial.parent)
                initialfile = initial.name
            elif initial.is_dir():
                initialdir = str(initial)
        fp = filedialog.askopenfilename(
            title="請求用ファイルを選択",
            filetypes=[
                ("Excel", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("すべて", "*.*"),
            ],
            parent=self._root,
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if not fp:
            return
        path = Path(fp)
        if not path.is_file():
            messagebox.showerror(
                "請求用ファイル",
                f"ファイルが見つかりません:\n{path}",
                parent=self._root,
            )
            return
        self._billing_file_path = path.resolve()
        self._sync_billing_file_display()
        self._update_data_dir_dependent_buttons()

    def _final_judgment_symbol_from_row(self, row: dict[str, str]) -> str:
        return normalize_judgment_symbol(
            (
                row.get(self.FINAL_JUDGMENT_COL)
                or row.get(self.LEGACY_USER_JUDGMENT_COL)
                or row.get("user_judgment_company")
                or ""
            ).strip()
        )

    def _set_billing_update_result_cell(self, rid: str, symbol: str) -> None:
        try:
            ci = list(self._tree["columns"]).index(self.BILLING_UPDATE_RESULT_COL)
        except ValueError:
            return
        vals = list(self._tree.item(rid, "values") or [])
        while len(vals) <= ci:
            vals.append("")
        vals[ci] = symbol
        self._tree.item(rid, values=tuple(vals))

    def _create_billing_data(self) -> None:
        if self._busy:
            return
        if not self._tree.get_children():
            messagebox.showinfo(
                "請求データ作成",
                "グリッドに行がありません。",
                parent=self._root,
            )
            return
        if not messagebox.askokcancel(
            "請求データ作成",
            "社員番号ごとに更新用合計勤務時間（10進）・更新用交通費合計を作成します。\n"
            "同一社員番号が複数ある場合は、No順の先頭行に合算値を設定します。",
            parent=self._root,
        ):
            return

        ordered: list[tuple[str, dict[str, str]]] = []
        for iid in self._tree.get_children():
            ordered.append((iid, self._row_dict_to_core(self._current_row_dict_from_iid(iid))))

        cores = [core for _, core in ordered]
        group_count = populate_billing_update_columns(cores)
        for (iid, _), core in zip(ordered, cores, strict=True):
            self._replace_row_with_result(iid, core)

        self._loaded_rows = self._current_grid_rows()
        self._status_var.set(
            f"請求データ作成完了: {group_count} グループに更新用列を設定しました"
        )
        messagebox.showinfo(
            "請求データ作成",
            f"更新用列を設定しました（{group_count} グループ）。",
            parent=self._root,
        )

    def _update_billing_file(self) -> None:
        if self._busy:
            return
        if self._billing_file_path is None or not self._billing_file_path.is_file():
            messagebox.showwarning(
                "請求ファイル更新",
                "請求用ファイルを選択してください。",
                parent=self._root,
            )
            return
        targets: list[tuple[str, dict[str, str]]] = []
        for iid in self._tree.get_children():
            ui_row = self._current_row_dict_from_iid(iid)
            if self._final_judgment_symbol_from_row(ui_row) != "〇":
                continue
            core_row = self._row_dict_to_core(ui_row)
            if not _row_billing_update_hours_decimal(core_row):
                continue
            targets.append((iid, core_row))
        if not targets:
            messagebox.showinfo(
                "請求ファイル更新",
                "最終判断が「〇」かつ更新用合計勤務時間（10進）が設定された行がありません。\n"
                "先に「請求データ作成」を実行してください。",
                parent=self._root,
            )
            return
        if not messagebox.askokcancel(
            "請求ファイル更新",
            "最終判断が「〇」のレコードについて、更新用合計勤務時間（10進）・"
            "更新用交通費合計を請求用ファイルへ反映します。",
            parent=self._root,
        ):
            return
        try:
            results = update_billing_engineer_ts_sheet(
                self._billing_file_path,
                [core_row for _, core_row in targets],
            )
        except OSError as e:
            messagebox.showerror(
                "請求ファイル更新",
                f"請求用ファイルを保存できませんでした。\n\n{e}",
                parent=self._root,
            )
            return
        except Exception as e:
            messagebox.showerror(
                "請求ファイル更新",
                f"請求用ファイルの更新に失敗しました。\n\n{e}",
                parent=self._root,
            )
            return
        ok_count = 0
        for (iid, _), symbol in zip(targets, results, strict=True):
            self._set_billing_update_result_cell(iid, symbol)
            if symbol == "〇":
                ok_count += 1
        fail_count = len(targets) - ok_count
        self._loaded_rows = self._current_grid_rows()
        self._status_var.set(
            f"請求ファイル更新完了: 成功 {ok_count} 件 / 失敗 {fail_count} 件 "
            f"（対象 {len(targets)} 件）"
        )
        messagebox.showinfo(
            "請求ファイル更新",
            (
                f"請求用ファイル:\n{self._billing_file_path.name}\n\n"
                f"成功: {ok_count} 件\n失敗: {fail_count} 件"
            ),
            parent=self._root,
        )

    def _should_ignore_status_log(self, message: str) -> bool:
        """run_analysis(on_log=...) 経由で流れてくるログのうち、
        GUIステータス欄に出すとノイズになりやすいものを除外する。

        例: チャット削除失敗（後片付け）など。
        """
        m = (message or "").strip()
        if not m:
            return False

        # 途中経過の割合はGUI側で組み立てるため、素のログは出さない
        if m.startswith("会社名比較 〇率(途中経過):"):
            return True

        # 後片付けのチャット削除失敗は、解析結果そのものには影響しないためステータスには出さない
        # （文言揺れ: 「チャットの削除に失敗」「チャット削除に失敗しました」「Failed to delete chat」など）
        low = m.lower()
        if ("チャット" in m or "chat" in low) and ("削除" in m or "delete" in low) and (
            "失敗" in m or "failed" in low
        ):
            return True
        if "削除に失敗" in m:
            return True
        return False

    def _is_error_reanalysis_target_row(self, row: dict[str, str]) -> bool:
        symbol = normalize_judgment_symbol(
            (
                row.get(self.FINAL_JUDGMENT_COL)
                or row.get(self.LEGACY_USER_JUDGMENT_COL)
                or row.get("user_judgment_company")
                or ""
            ).strip()
        )
        # 空欄や △/✖ 等も「〇以外」として対象にする
        return symbol not in self._ERROR_REANALYSIS_OK_VALUES

    def _eligible_error_reanalysis_iids(self) -> list[str]:
        iids: list[str] = []
        for iid in self._tree.get_children():
            r = self._current_row_dict_from_iid(iid)
            if self._is_error_reanalysis_target_row(r):
                iids.append(iid)
        return iids

    def _selected_tree_iids(self) -> list[str]:
        return list(self._tree.selection())

    def _refresh_selected_rows_reanalysis_button_state(self) -> None:
        if self._busy or not self._analysis_prerequisites_met():
            self._selected_reanalysis_btn.configure(state=tk.DISABLED)
            return
        enabled = bool(self._selected_tree_iids())
        self._selected_reanalysis_btn.configure(state=(tk.NORMAL if enabled else tk.DISABLED))

    def _refresh_error_reanalysis_button_state(self) -> None:
        if self._busy or not self._analysis_prerequisites_met():
            self._error_reanalysis_btn.configure(state=tk.DISABLED)
            return
        # 対象行が1件でもあれば有効
        enabled = bool(self._eligible_error_reanalysis_iids())
        self._error_reanalysis_btn.configure(state=(tk.NORMAL if enabled else tk.DISABLED))

    def _refresh_reanalysis_buttons_state(self) -> None:
        self._refresh_error_reanalysis_button_state()
        self._refresh_selected_rows_reanalysis_button_state()
        self._refresh_billing_update_button_state()

    def _clear_grid(self) -> None:
        self._close_row_file()
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._item_paths.clear()
        self._sort_column = None
        self._sort_reverse = False
        self._update_column_heading_labels()

    def _grid_sort_empty_marker(self, value: str) -> bool:
        t = (value or "").strip()
        return not t or t in ("（なし）", "不明", "（不明）", "（対象レコードなし）")

    def _grid_sort_numeric_columns(self) -> frozenset[str]:
        return frozenset(
            {
                self.ROW_NO_COL,
                self.YEAR_COL,
                self.MONTH_COL,
                self.TOTAL_HOURS_DECIMAL_COL,
                self.BILLING_UPDATE_HOURS_COL,
                self.BILLING_UPDATE_TRANSPORT_COL,
                self.TRANSPORT_EXPENSE_COL,
            }
        )

    def _grid_sort_key(self, value: str, column: str) -> tuple[int, float | str]:
        if self._grid_sort_empty_marker(value):
            return (2, "")
        t = (value or "").strip()
        if column in self._grid_sort_numeric_columns():
            cleaned = t.replace(",", "").replace("，", "").replace("円", "")
            num_match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
            if num_match:
                try:
                    return (0, float(num_match.group()))
                except ValueError:
                    pass
        return (1, t.casefold())

    def _update_column_heading_labels(self) -> None:
        for h in self._tree_column_headings:
            text = h
            if h == self._sort_column:
                text += " ▲" if not self._sort_reverse else " ▼"
            try:
                self._tree.heading(h, text=text)
            except tk.TclError:
                pass

    def _sort_grid_by_column(self, column: str) -> None:
        """列見出しクリックで昇順/降順をトグルし、行全体を並べ替える。"""
        if column not in self._tree_column_headings:
            return
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False

        cols = list(self._tree["columns"])
        try:
            col_idx = cols.index(column)
        except ValueError:
            return

        children = list(self._tree.get_children())
        if len(children) <= 1:
            self._update_column_heading_labels()
            return

        def row_key(iid: str) -> tuple[int, float | str]:
            values = self._tree.item(iid, "values") or ()
            cell = values[col_idx] if col_idx < len(values) else ""
            return self._grid_sort_key(str(cell), column)

        for index, iid in enumerate(
            sorted(children, key=row_key, reverse=self._sort_reverse)
        ):
            self._tree.move(iid, "", index)

        self._refresh_row_numbers()
        self._update_column_heading_labels()

    def _resolve_file_path_for_row(self, rid: str) -> Path | None:
        """行に対応する画像/PDF/Excel の絶対パスを返す。"""
        path_str = self._item_paths.get(rid, "").strip()
        if path_str:
            p = Path(path_str)
            if p.is_file():
                return p.resolve()

        if self._data_dir is None:
            return None
        row = self._current_row_dict_from_iid(rid)
        fn = self._file_name_from_row(row)
        if not fn:
            return None
        p = (self._data_dir / fn).resolve()
        if p.is_file():
            self._item_paths[rid] = str(p)
            return p
        return None

    def _close_row_file(self) -> None:
        """前に開いたプレビューファイルを閉じる（可能な範囲）。"""
        path = self._preview_file_path
        if path is not None and path.suffix.lower() in EXCEL_SUFFIXES:
            _try_close_excel_workbook(path)

        _win_terminate_process_handle(self._preview_process_handle)
        self._preview_process_handle = None
        self._preview_file_path = None
        self._preview_row_iid = ""

    def _open_row_file(self, rid: str, *, force: bool = False) -> None:
        """行に対応するファイルを既定アプリで開く。"""
        if not force and rid == self._preview_row_iid:
            return
        self._close_row_file()
        path = self._resolve_file_path_for_row(rid)
        if path is None:
            return
        try:
            handle = _win_shell_open_file(path)
        except OSError as e:
            messagebox.showerror("起動できませんでした", str(e))
            return
        self._preview_row_iid = rid
        self._preview_file_path = path
        self._preview_process_handle = handle

    def _column_index_at_event(self, event: tk.Event) -> int:
        if self._tree.identify_region(event.x, event.y) != "cell":
            return -1
        col_w = self._tree.identify_column(event.x)
        try:
            return int(col_w.replace("#", "")) - 1
        except ValueError:
            return -1

    def _on_tree_left_click(self, event: tk.Event) -> None:
        """対象ファイル名列の左クリックでファイルを開く（別行なら前のファイルを閉じる）。"""
        rid = self._tree.identify_row(event.y)
        if not rid:
            return
        cols = list(self._tree["columns"])
        ci = self._column_index_at_event(event)
        if not (0 <= ci < len(cols) and cols[ci] == self.TARGET_FILE_NAME_COL):
            return
        self._tree.selection_set(rid)
        self._open_row_file(rid)

    def _prepare_native_dialog(self) -> None:
        """Windows/Tk でネイティブダイアログが背面化・ハング見えしないよう状態を整える。"""
        try:
            grabbed = self._root.grab_current()
            if grabbed is not None:
                grabbed.grab_release()
        except tk.TclError:
            pass
        try:
            self._root.deiconify()
            self._root.lift()
            self._root.focus_force()
            self._root.update_idletasks()
        except tk.TclError:
            pass

    def _file_name_from_row(self, row: dict[str, str]) -> str:
        return (
            row.get(self.TARGET_FILE_NAME_COL)
            or row.get(self.LEGACY_TARGET_FILE_NAME_COL)
            or row.get(self.LEGACY_FILE_NAME_COL)
            or row.get("file_name")
            or ""
        ).strip()

    def _row_dict_to_core(self, row: dict[str, str]) -> dict[str, str]:
        """Treeview/JSON 行を kintai_core 互換 dict に変換する。"""
        out = dict(row)
        pairs = (
            (self.TARGET_FILE_NAME_COL, "file_name"),
            ("アップロード", "upload_ok"),
            ("対象シート有無", "target_sheet_exists"),
            (self.FINAL_JUDGMENT_COL, "user_judgment_company"),
            (self.BILLING_UPDATE_RESULT_COL, "billing_file_update_result"),
            (self.YEAR_COL, "year"),
            (self.MONTH_COL, "month"),
            (self.COMPANY_COL, "name_company_1"),
            (self.PERSON_COL, "name_person_from_doc"),
            (self.EMPLOYEE_NO_COL, "employee_no"),
            (self.BILLING_UPDATE_HOURS_COL, "billing_update_hours_decimal"),
            (LEGACY_BILLING_UPDATE_HOURS_COL, "billing_update_hours_decimal"),
            (self.TOTAL_HOURS_DECIMAL_COL, "total_hours_decimal"),
            (self.TOTAL_HOURS_RAW_COL, "total_hours_raw"),
            (self.BILLING_UPDATE_TRANSPORT_COL, "billing_update_transport"),
            (self.TRANSPORT_EXPENSE_COL, "transport_expense_raw"),
            (self.MATCH_COMPANY_COL, "match_company"),
            ("押印有無", "seal_in_doc"),
        )
        for ui_key, core_key in pairs:
            if ui_key in row:
                val = str(row.get(ui_key) or "").strip()
            elif core_key == "file_name":
                val = self._file_name_from_row(row)
            elif core_key == "year":
                if core_key in row:
                    val = str(row.get(core_key) or "").strip()
                else:
                    val = str(
                        row.get(self.YEAR_COL)
                        or row.get(self.LEGACY_YEAR_COL)
                        or self._year_var.get()
                        or ""
                    ).strip()
            elif core_key == "month":
                if core_key in row:
                    val = str(row.get(core_key) or "").strip()
                else:
                    val = str(
                        row.get(self.MONTH_COL)
                        or row.get(self.LEGACY_MONTH_COL)
                        or self._month_var.get()
                        or ""
                    ).strip()
            elif core_key == "match_company":
                val = str(
                    row.get(core_key)
                    or row.get(self.LEGACY_MATCH_COMPANY_COL)
                    or ""
                ).strip()
            elif core_key == "name_company_1":
                val = str(
                    row.get(self.COMPANY_COL)
                    or row.get(self.LEGACY_COMPANY_READ_LONG_COL)
                    or row.get(self.LEGACY_COMPANY_NAME_COL)
                    or row.get(self.LEGACY_COMPANY_COL)
                    or row.get(core_key)
                    or ""
                ).strip()
            elif core_key == "name_person_from_doc":
                val = str(
                    row.get(self.PERSON_COL)
                    or row.get(self.LEGACY_PERSON_COL)
                    or row.get(core_key)
                    or ""
                ).strip()
            elif core_key == "employee_no":
                val = str(
                    row.get(self.EMPLOYEE_NO_COL)
                    or row.get(self.LEGACY_EMPLOYEE_NO_COL)
                    or row.get(core_key)
                    or ""
                ).strip()
            elif core_key == "user_judgment_company":
                val = str(
                    row.get(self.FINAL_JUDGMENT_COL)
                    or row.get(self.LEGACY_USER_JUDGMENT_COL)
                    or row.get(core_key)
                    or ""
                ).strip()
                val = normalize_judgment_symbol(val) if val else ""
            elif core_key == "billing_file_update_result":
                val = str(
                    row.get(self.BILLING_UPDATE_RESULT_COL)
                    or row.get("請求量ファイル更新結果")
                    or row.get(core_key)
                    or ""
                ).strip()
            else:
                val = str(row.get(core_key) or "").strip()
            if core_key == "user_judgment_company":
                out[core_key] = normalize_judgment_symbol(val) if val else ""
            else:
                out[core_key] = val
        ey, em = self._expected_year_month()
        if ey is not None:
            out["expected_year"] = str(ey)
        if em is not None:
            out["expected_month"] = str(em)
        return out

    def _grid_values_from_row(
        self, row: dict[str, str], *, sync_user_judgment_to_auto: bool = False
    ) -> tuple[str, ...]:
        """内部row形式・UI保存形式のどちらでも Treeview 表示値へ変換する。"""
        return row_display_values(
            self._row_dict_to_core(row),
            sync_user_judgment_to_auto=sync_user_judgment_to_auto,
        )

    def _refresh_row_numbers(self) -> None:
        """No 列に 1 からの連番を付与する。"""
        cols = list(self._tree["columns"])
        try:
            no_ci = cols.index(self.ROW_NO_COL)
        except ValueError:
            return
        for seq, iid in enumerate(self._tree.get_children(), start=1):
            vals = list(self._tree.item(iid, "values") or [])
            while len(vals) <= no_ci:
                vals.append("")
            vals[no_ci] = str(seq)
            self._tree.item(iid, values=tuple(vals))

    def _user_judgment_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.FINAL_JUDGMENT_COL)

    def _employee_no_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.EMPLOYEE_NO_COL)

    def _total_hours_decimal_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.TOTAL_HOURS_DECIMAL_COL)

    def _total_hours_raw_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.TOTAL_HOURS_RAW_COL)

    def _recalculate_auto_judgment_for_row(
        self, rid: str, row: dict[str, str] | None = None
    ) -> None:
        """列編集後に自動判断を再計算し、ユーザ判断も同じ結果で上書きする。"""
        core = self._row_dict_to_core(
            row if row is not None else self._current_row_dict_from_iid(rid)
        )
        aj = auto_judgment_symbol(core)
        core["auto_judgment"] = aj
        core["user_judgment_company"] = aj
        core[self.FINAL_JUDGMENT_COL] = aj
        self._replace_row_with_result(
            rid, core, sync_user_judgment_to_auto=True
        )

    def _on_tree_right_click(self, event: tk.Event) -> None:
        if self._busy:
            return
        rid = self._tree.identify_row(event.y)
        if not rid:
            return
        if self._tree.identify_region(event.x, event.y) != "cell":
            return
        col_w = self._tree.identify_column(event.x)
        try:
            ci = int(col_w.replace("#", "")) - 1
        except ValueError:
            return
        cols = list(self._tree["columns"])
        if not (0 <= ci < len(cols)):
            return

        self._tree.selection_set(rid)
        menu = tk.Menu(self, tearoff=0)

        # 最終判断列: 右クリックで 〇/△/✖ を選択
        if cols[ci] == self.FINAL_JUDGMENT_COL:
            opts = (
                ("〇", "〇"),
                ("△", "△"),
                ("✖", "✖"),
            )
            for label, inner in opts:
                menu.add_command(
                    label=label,
                    command=lambda v=inner: self._set_user_judgment_cell(rid, v),
                )
            menu.add_separator()

        # 社員番号列: 右クリックで入力（空欄許可）
        if cols[ci] == self.EMPLOYEE_NO_COL:
            menu.add_command(
                label="社員番号を編集",
                command=lambda: self._prompt_edit_employee_no(rid),
            )
            menu.add_separator()

        # 合計勤務時間（読取）列: 右クリックで入力し、10進列も更新
        if cols[ci] == self.TOTAL_HOURS_RAW_COL:
            menu.add_command(
                label="合計勤務時間（読取）を編集",
                command=lambda: self._prompt_edit_total_hours_raw(rid),
            )
            menu.add_separator()

        if cols[ci] == self.YEAR_COL:
            menu.add_command(
                label="年を編集",
                command=lambda: self._prompt_edit_year(rid),
            )
            menu.add_separator()

        if cols[ci] == self.MONTH_COL:
            menu.add_command(
                label="月を編集",
                command=lambda: self._prompt_edit_month(rid),
            )
            menu.add_separator()

        menu.add_command(label="再解析", command=lambda: self._start_row_reanalysis(rid))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _current_row_dict_from_iid(self, rid: str) -> dict[str, str]:
        cols = list(self._tree["columns"])
        values = list(self._tree.item(rid, "values") or [])
        row = {cols[i]: (values[i] if i < len(values) else "") for i in range(len(cols))}
        row["resolved_path"] = self._item_paths.get(rid, "")
        return row

    def _replace_row_with_result(
        self,
        rid: str,
        row: dict[str, str],
        *,
        sync_user_judgment_to_auto: bool = False,
    ) -> None:
        self._tree.item(
            rid,
            values=self._grid_values_from_row(
                row, sync_user_judgment_to_auto=sync_user_judgment_to_auto
            ),
        )
        self._item_paths[rid] = row.get("resolved_path", "")
        self._refresh_row_numbers()

    def _set_row_reanalysis_highlight(self, rid: str, active: bool) -> None:
        """Treeview の指定行に、再解析中ハイライト（反転）を付与/解除する。"""
        try:
            cur = tuple(self._tree.item(rid, "tags") or ())
        except tk.TclError:
            return
        tag = self._TAG_REANALYSIS_ACTIVE
        if active:
            if tag not in cur:
                self._tree.item(rid, tags=cur + (tag,))
        else:
            if tag in cur:
                self._tree.item(rid, tags=tuple(t for t in cur if t != tag))

    def _set_rows_reanalysis_highlight(self, rids: list[str], active: bool) -> None:
        for rid in rids:
            self._set_row_reanalysis_highlight(rid, active)

    def _start_row_reanalysis(self, rid: str) -> None:
        if self._busy:
            return
        if not self._analysis_prerequisites_met():
            messagebox.showwarning(
                "再解析",
                "再解析を実行するには、データフォルダと請求用ファイルの両方を選択してください。",
                parent=self._root,
            )
            return

        current_row = self._current_row_dict_from_iid(rid)

        # 選択行の「再解析」は、会社名の内容に関わらず実行可能とする。
        # （不明/（存在しない）のみ一括で再解析したい場合は「エラー再解析」ボタンを使用。）

        file_name = self._file_name_from_row(current_row)
        if not file_name:
            messagebox.showinfo("再解析", "選択行のファイル名を取得できません。")
            return

        td = self._data_dir.resolve()
        client = self._client
        aid = self._require_assistant_uid()
        if not aid:
            return

        self._busy = True
        self._cancel_event = threading.Event()
        self._new_btn.configure(state=tk.DISABLED)
        self._new_plus_error_btn.configure(state=tk.DISABLED)
        self._cont_btn.configure(state=tk.DISABLED)
        self._selected_reanalysis_btn.configure(state=tk.DISABLED)
        self._save_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._error_reanalysis_btn.configure(state=tk.DISABLED)
        self._billing_prepare_btn.configure(state=tk.DISABLED)
        self._billing_update_btn.configure(state=tk.DISABLED)
        self._progress_var.set("再解析中 0 / 1")
        self._status_var.set(f"再解析しています… {file_name}")

        # 対象行を反転表示
        self._set_row_reanalysis_highlight(rid, True)

        def log_line(message: str) -> None:
            def apply_log(m: str = message) -> None:
                if self._should_ignore_status_log(m):
                    return
                self._status_var.set(m[:800])

            self.after(0, apply_log)

        def on_progress(done: int, total: int) -> None:
            self.after(0, lambda d=done, t=total: self._progress_var.set(f"再解析中 {d} / {t}"))

        def worker() -> None:
            result_rows: list[dict[str, str]] | None = None
            err: BaseException | None = None
            try:
                exp_y, exp_m = self._expected_year_month()
                result_rows = run_analysis(
                    client,
                    aid,
                    td,
                    save_md_path=Path.cwd() / "解析結果.md",
                    on_log=log_line,
                    emit_progress_md_rows=False,
                    on_file_progress=on_progress,
                    cancel_event=self._cancel_event,
                    target_file_names={file_name},
                    parallel_chats=1,
                    expected_year=exp_y,
                    expected_month=exp_m,
                )
            except BaseException as e:
                err = e

            def finish() -> None:
                self._busy = False
                self._cancel_btn.configure(state=tk.DISABLED)
                cancelled = self._cancel_event is not None and self._cancel_event.is_set()
                self._cancel_event = None

                # 反転表示を解除
                self._set_row_reanalysis_highlight(rid, False)

                self._update_data_dir_dependent_buttons()

                if err is not None:
                    messagebox.showerror("再解析エラー", str(err))
                    self._status_var.set(f"再解析エラー: {file_name}")
                    return
                if cancelled:
                    self._status_var.set(f"再解析を中断しました: {file_name}")
                    return
                if not result_rows:
                    messagebox.showwarning("再解析", f"再解析結果を取得できませんでした。\n{file_name}")
                    self._status_var.set(f"再解析結果なし: {file_name}")
                    return

                new_row = result_rows[0]
                # 再解析時は自動判断でユーザ判断を上書きする
                new_row["user_judgment_company"] = auto_judgment_symbol(new_row)
                self._replace_row_with_result(rid, new_row)
                self._loaded_rows = self._current_grid_rows()
                self._progress_var.set("再解析完了 1 / 1")
                ratio_text = self._company_match_ratio_text(self._loaded_rows)
                self._status_var.set(f"再解析完了: {file_name} / {ratio_text}")
                self._refresh_reanalysis_buttons_state()
                messagebox.showinfo("再解析完了", f"選択行の再解析が完了しました。\n{file_name}")

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _iid_by_file_from_iids(self, iids: list[str]) -> dict[str, str]:
        """Treeview 行 ID のリストから file_name -> iid の対応を作る。"""
        iid_by_file: dict[str, str] = {}
        for iid in iids:
            row = self._current_row_dict_from_iid(iid)
            file_name = self._file_name_from_row(row)
            if file_name:
                iid_by_file[file_name] = iid
        return iid_by_file

    def _run_targeted_reanalysis(
        self,
        iid_by_file: dict[str, str],
        *,
        label: str,
        complete_message: str,
    ) -> None:
        """指定行（ファイル名）のみ再解析する。"""
        if self._busy or not iid_by_file:
            return
        if not self._analysis_prerequisites_met():
            messagebox.showwarning(
                label,
                f"{label}を実行するには、データフォルダと請求用ファイルの両方を選択してください。",
                parent=self._root,
            )
            return

        file_names = set(iid_by_file.keys())
        total = len(file_names)
        target_iids = list(iid_by_file.values())

        td = self._data_dir.resolve()
        client = self._client
        aid = self._require_assistant_uid()
        if not aid:
            return
        parallel = self._parallel_workers_value()

        self._busy = True
        self._cancel_event = threading.Event()
        self._new_btn.configure(state=tk.DISABLED)
        self._new_plus_error_btn.configure(state=tk.DISABLED)
        self._cont_btn.configure(state=tk.DISABLED)
        self._selected_reanalysis_btn.configure(state=tk.DISABLED)
        self._save_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._error_reanalysis_btn.configure(state=tk.DISABLED)
        self._billing_prepare_btn.configure(state=tk.DISABLED)
        self._billing_update_btn.configure(state=tk.DISABLED)
        self._progress_var.set(f"{label}中 0 / {total}")
        self._status_var.set(f"{label}しています… {total} 件（並列 {parallel}）")

        self._set_rows_reanalysis_highlight(target_iids, False)
        active_rids: set[str] = set()

        def log_line(message: str) -> None:
            def apply_log(m: str = message) -> None:
                if self._should_ignore_status_log(m):
                    return
                self._status_var.set(m[:800])

            self.after(0, apply_log)

        def on_progress(done: int, t: int) -> None:
            self.after(
                0, lambda d=done, tt=t: self._progress_var.set(f"{label}中 {d} / {tt}")
            )

        def on_file_started(file_name: str) -> None:
            fn = (file_name or "").strip()
            if not fn:
                return
            rid = iid_by_file.get(fn) or ""
            if not rid:
                return

            def apply_started() -> None:
                active_rids.add(rid)
                self._set_row_reanalysis_highlight(rid, True)
                try:
                    self._tree.see(rid)
                except tk.TclError:
                    pass

            self.after(0, apply_started)

        def on_row_completed(row: dict[str, str]) -> None:
            fn = (row.get("file_name") or "").strip()
            if not fn:
                return
            row["user_judgment_company"] = auto_judgment_symbol(row)
            rid = iid_by_file.get(fn)
            if not rid:
                return

            def apply_row() -> None:
                self._replace_row_with_result(
                    rid, row, sync_user_judgment_to_auto=True
                )
                self._set_row_reanalysis_highlight(rid, False)
                active_rids.discard(rid)
                current_rows = self._current_grid_rows()
                ratio_text = self._company_match_ratio_text(
                    current_rows, total_target_count=len(current_rows)
                )
                self._status_var.set(f"{label}中… / {ratio_text}")

            self.after(0, apply_row)

        def worker() -> None:
            err: BaseException | None = None
            try:
                exp_y, exp_m = self._expected_year_month()
                run_analysis(
                    client,
                    aid,
                    td,
                    save_md_path=Path.cwd() / "解析結果.md",
                    on_log=log_line,
                    emit_progress_md_rows=False,
                    on_file_started=on_file_started,
                    on_file_progress=on_progress,
                    on_row_completed=on_row_completed,
                    cancel_event=self._cancel_event,
                    target_file_names=file_names,
                    parallel_chats=parallel,
                    expected_year=exp_y,
                    expected_month=exp_m,
                )
            except BaseException as e:
                err = e

            def finish() -> None:
                self._busy = False
                self._cancel_btn.configure(state=tk.DISABLED)
                cancelled = self._cancel_event is not None and self._cancel_event.is_set()
                self._cancel_event = None

                for ar in list(active_rids):
                    self._set_row_reanalysis_highlight(ar, False)
                active_rids.clear()
                self._set_rows_reanalysis_highlight(target_iids, False)

                self._update_data_dir_dependent_buttons()

                self._loaded_rows = self._current_grid_rows()
                self._progress_var.set(f"{label}完了 {total} / {total}")
                ratio_text = self._company_match_ratio_text(self._loaded_rows)

                if err is not None:
                    messagebox.showerror(f"{label}エラー", str(err))
                    self._status_var.set(f"{label}エラー（途中結果は保持） / {ratio_text}")
                    return
                if cancelled:
                    self._status_var.set(f"{label}を中断しました（途中結果は保持） / {ratio_text}")
                    return

                self._status_var.set(f"{label}完了: {total} 件 / {ratio_text}")
                messagebox.showinfo(f"{label}完了", complete_message)

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _start_selected_rows_reanalysis(self) -> None:
        """グリッドで選択した行だけを再解析する。"""
        if self._busy:
            return
        iids = self._selected_tree_iids()
        if not iids:
            return
        iid_by_file = self._iid_by_file_from_iids(iids)
        if not iid_by_file:
            messagebox.showwarning(
                "選択行解析", "選択行のファイル名を取得できませんでした。"
            )
            self._refresh_selected_rows_reanalysis_button_state()
            return
        total = len(iid_by_file)
        self._run_targeted_reanalysis(
            iid_by_file,
            label="選択行解析",
            complete_message=f"選択した行を再解析しました。\n対象: {total} 件",
        )

    def _start_error_reanalysis(self) -> None:
        """ユーザ判断が「〇」以外の行だけをまとめて再解析する。"""
        if self._busy:
            return

        target_iids = self._eligible_error_reanalysis_iids()
        if not target_iids:
            messagebox.showinfo(
                "エラー再解析", "対象行がありません（最終判断が『〇』以外の行）。"
            )
            self._refresh_reanalysis_buttons_state()
            return

        iid_by_file = self._iid_by_file_from_iids(target_iids)
        if not iid_by_file:
            messagebox.showwarning("エラー再解析", "対象行のファイル名を取得できませんでした。")
            self._refresh_reanalysis_buttons_state()
            return
        total = len(iid_by_file)
        self._run_targeted_reanalysis(
            iid_by_file,
            label="エラー再解析",
            complete_message=f"最終判断が〇以外の行を再解析しました。\n対象: {total} 件",
        )

    def _set_user_judgment_cell(self, rid: str, value: str) -> None:
        value = normalize_judgment_symbol(value)
        if value not in ("〇", "△", "✖"):
            return
        ci = self._user_judgment_column_index()
        vals = list(self._tree.item(rid, "values"))
        if ci >= len(vals):
            while len(vals) <= ci:
                vals.append("")
        vals[ci] = value
        self._tree.item(rid, values=tuple(vals))

    def _prompt_edit_employee_no(self, rid: str) -> None:
        """社員番号セルを右クリックから編集する（空欄/6桁数字/社員番号エラー等を許容）。"""
        try:
            ci = self._employee_no_column_index()
        except ValueError:
            return
        vals = list(self._tree.item(rid, "values") or [])
        cur = vals[ci] if ci < len(vals) else ""

        top = tk.Toplevel(self)
        top.title("社員番号を編集")
        top.transient(self)
        top.grab_set()

        ttk.Label(
            top,
            text=(
                "社員番号（7桁数字 または BP+5桁）を入力してください。\n"
                "- 空欄: 未設定\n"
                "- 7桁数字: 社員番号\n"
                "- BP+5桁: 社員番号\n"
                "- それ以外: 社員番号エラー として扱われます"
            ),
            justify="left",
        ).pack(fill=tk.X, padx=10, pady=(10, 6))

        var = tk.StringVar(value=str(cur))
        ent = ttk.Entry(top, textvariable=var, width=24)
        ent.pack(fill=tk.X, padx=10)
        ent.focus_set()
        ent.select_range(0, tk.END)

        btns = ttk.Frame(top)
        btns.pack(fill=tk.X, padx=10, pady=10)

        def normalize(v: str) -> str:
            t = (v or "").strip()
            if not t:
                return ""
            if t.isdigit() and len(t) == 7:
                return t
            if len(t) == 7 and t[:2].upper() == "BP" and t[2:].isdigit():
                return t[:2].upper() + t[2:]
            return "社員番号エラー"

        def on_ok() -> None:
            new_v = normalize(var.get())
            row = self._row_dict_to_core(self._current_row_dict_from_iid(rid))
            row["employee_no"] = new_v
            row[self.EMPLOYEE_NO_COL] = new_v
            self._recalculate_auto_judgment_for_row(rid, row)
            top.destroy()

        def on_cancel() -> None:
            top.destroy()

        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 8))

        top.bind("<Return>", lambda _e: on_ok())
        top.bind("<Escape>", lambda _e: on_cancel())

    def _prompt_edit_year(self, rid: str) -> None:
        """年列を右クリックから編集する（空欄可）。"""
        row = self._row_dict_to_core(self._current_row_dict_from_iid(rid))
        cur = (row.get("year") or "").strip()

        top = tk.Toplevel(self)
        top.title("年を編集")
        top.transient(self)
        top.grab_set()

        ttk.Label(
            top,
            text="年（4桁の西暦）を入力してください。\n空欄: 未設定",
            justify="left",
        ).pack(fill=tk.X, padx=10, pady=(10, 6))

        var = tk.StringVar(value=cur)
        ent = ttk.Entry(top, textvariable=var, width=12)
        ent.pack(fill=tk.X, padx=10)
        ent.focus_set()
        ent.select_range(0, tk.END)

        btns = ttk.Frame(top)
        btns.pack(fill=tk.X, padx=10, pady=10)

        def on_ok() -> None:
            raw = (var.get() or "").strip()
            if not raw:
                new_y = ""
            else:
                norm = _normalize_year_value(raw)
                if not norm:
                    messagebox.showwarning(
                        "入力エラー",
                        "年は4桁の西暦（例: 2026）で入力してください。",
                        parent=top,
                    )
                    return
                new_y = norm
            row["year"] = new_y
            row[self.YEAR_COL] = new_y
            self._recalculate_auto_judgment_for_row(rid, row)
            top.destroy()

        def on_cancel() -> None:
            top.destroy()

        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(
            side=tk.RIGHT, padx=(0, 8)
        )
        top.bind("<Return>", lambda _e: on_ok())
        top.bind("<Escape>", lambda _e: on_cancel())

    def _prompt_edit_month(self, rid: str) -> None:
        """月列を右クリックから編集する（空欄可）。"""
        row = self._row_dict_to_core(self._current_row_dict_from_iid(rid))
        cur = (row.get("month") or "").strip()

        top = tk.Toplevel(self)
        top.title("月を編集")
        top.transient(self)
        top.grab_set()

        ttk.Label(
            top,
            text="月（1〜12）を入力してください。\n空欄: 未設定",
            justify="left",
        ).pack(fill=tk.X, padx=10, pady=(10, 6))

        var = tk.StringVar(value=cur)
        ent = ttk.Entry(top, textvariable=var, width=8)
        ent.pack(fill=tk.X, padx=10)
        ent.focus_set()
        ent.select_range(0, tk.END)

        btns = ttk.Frame(top)
        btns.pack(fill=tk.X, padx=10, pady=10)

        def on_ok() -> None:
            raw = (var.get() or "").strip()
            if not raw:
                new_m = ""
            else:
                norm = _normalize_month_value(raw)
                if not norm:
                    messagebox.showwarning(
                        "入力エラー",
                        "月は1〜12の整数（例: 3 または 3月）で入力してください。",
                        parent=top,
                    )
                    return
                new_m = norm
            row["month"] = new_m
            row[self.MONTH_COL] = new_m
            self._recalculate_auto_judgment_for_row(rid, row)
            top.destroy()

        def on_cancel() -> None:
            top.destroy()

        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(
            side=tk.RIGHT, padx=(0, 8)
        )
        top.bind("<Return>", lambda _e: on_ok())
        top.bind("<Escape>", lambda _e: on_cancel())

    def _prompt_edit_total_hours_raw(self, rid: str) -> None:
        """合計勤務時間（読取）を右クリックから編集し、10進列も同時更新する。"""
        try:
            raw_ci = self._total_hours_raw_column_index()
            dec_ci = self._total_hours_decimal_column_index()
        except ValueError:
            return

        vals = list(self._tree.item(rid, "values") or [])
        cur = vals[raw_ci] if raw_ci < len(vals) else ""

        top = tk.Toplevel(self)
        top.title("合計勤務時間（読取）を編集")
        top.transient(self)
        top.grab_set()

        ttk.Label(
            top,
            text=(
                "合計勤務時間（読取）を入力してください。\n"
                "例: 8:20 / 8時間20分 / 8.20 / 101_10H / 86.17H\n"
                "入力後、同じ規則で10進数へ変換して右列へ反映します。"
            ),
            justify="left",
        ).pack(fill=tk.X, padx=10, pady=(10, 6))

        var = tk.StringVar(value=str(cur))
        ent = ttk.Entry(top, textvariable=var, width=28)
        ent.pack(fill=tk.X, padx=10)
        ent.focus_set()
        ent.select_range(0, tk.END)

        btns = ttk.Frame(top)
        btns.pack(fill=tk.X, padx=10, pady=10)

        def on_ok() -> None:
            new_raw = (var.get() or "").strip()
            dec_str = _work_hours_string_to_decimal(new_raw) if new_raw else ""
            new_dec = (
                _decimal_for_table_display(dec_str) if dec_str else "（なし）"
            )
            row = self._row_dict_to_core(self._current_row_dict_from_iid(rid))
            row["total_hours_raw"] = new_raw or "（なし）"
            row[self.TOTAL_HOURS_RAW_COL] = row["total_hours_raw"]
            row["total_hours_decimal"] = dec_str if dec_str else ""
            row[self.TOTAL_HOURS_DECIMAL_COL] = new_dec
            self._recalculate_auto_judgment_for_row(rid, row)
            top.destroy()

        def on_cancel() -> None:
            top.destroy()

        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 8))

        top.bind("<Return>", lambda _e: on_ok())
        top.bind("<Escape>", lambda _e: on_cancel())

    def _cancel_analysis(self) -> None:
        if not self._busy:
            return
        if self._cancel_event is not None:
            self._cancel_event.set()
        self._cancel_btn.configure(state=tk.DISABLED)
        self._status_var.set("中断しています…（現在の処理が終わり次第停止します）")

    def _current_grid_rows(self) -> list[dict[str, str]]:
        """現在のグリッド表示を dict 行の配列に戻す（保存用）。"""
        cols = list(self._tree["columns"])
        rows: list[dict[str, str]] = []
        for iid in self._tree.get_children():
            values = list(self._tree.item(iid, "values") or [])
            row = {cols[i]: (values[i] if i < len(values) else "") for i in range(len(cols))}
            row.pop(self.ROW_NO_COL, None)
            # 内部データ
            row["resolved_path"] = self._item_paths.get(iid, "")
            rows.append(row)
        return rows

    def _compute_snapshot(self, rows: list[dict[str, str]]) -> str:
        try:
            return json.dumps(rows, ensure_ascii=False, sort_keys=True)
        except Exception:
            return str(rows)

    def _company_match_ratio_text(
        self,
        rows: list[dict[str, str]],
        *,
        total_target_count: int | None = None,
        prefix: str = "会社名比較 〇率",
    ) -> str:
        """会社名比較の〇率文字列を返す。△は〇側に含め、分母は全レコード件数。"""
        ok_count = 0
        processed_count = 0
        for row in rows:
            if not isinstance(row, dict):
                continue
            symbol = (
                (row.get(self.MATCH_COMPANY_COL) or "").strip()
                or (row.get(self.LEGACY_MATCH_COMPANY_COL) or "").strip()
                or (row.get("match_company") or "").strip()
            )
            processed_count += 1
            if symbol in ("〇", "△"):
                ok_count += 1

        target_count = total_target_count if total_target_count is not None else processed_count
        if processed_count <= 0:
            return f"{prefix}: 対象データなし（実行済 0 / 対象 {target_count}）"

        ratio = (ok_count / processed_count) * 100
        return (
            f"{prefix}: {ratio:.1f}% "
            f"（〇扱い {ok_count}件 / 実行済 {processed_count}件 / 対象 {target_count}件）"
        )

    def _json_paths_payload(self) -> dict[str, str]:
        """保存用 JSON に含めるデータフォルダ・請求用ファイルのパス。"""
        folder = (
            str(self._data_dir.resolve())
            if self._data_dir is not None and self._data_dir.is_dir()
            else ""
        )
        billing = (
            str(self._billing_file_path.resolve())
            if self._billing_file_path is not None
            and self._billing_file_path.is_file()
            else ""
        )
        return {
            "data_dir": folder,
            "data_folder": folder,
            "billing_file_path": billing,
            "billing_file": billing,
        }

    def _restore_paths_from_json(self, data: dict) -> list[str]:
        """JSON からデータフォルダ・請求用ファイルを復元する。警告メッセージのリストを返す。"""
        warnings: list[str] = []

        dd = (data.get("data_dir") or data.get("data_folder") or "").strip()
        if dd:
            data_path = Path(dd)
            if data_path.is_dir():
                self._set_data_path(data_path, sync_year_month_from_path=True)
            else:
                self._data_dir = None
                self._data_branch = ""
                self._sync_folder_display()
                self._refresh_year_month_combos()
                warnings.append(f"データフォルダが見つかりません:\n{dd}")
        else:
            self._data_dir = None
            self._data_branch = ""
            self._sync_folder_display()
            self._refresh_year_month_combos()

        bf = (data.get("billing_file_path") or data.get("billing_file") or "").strip()
        if bf:
            billing_path = Path(bf)
            if billing_path.is_file():
                self._billing_file_path = billing_path.resolve()
            else:
                self._billing_file_path = None
                warnings.append(f"請求用ファイルが見つかりません:\n{bf}")
        else:
            self._billing_file_path = None

        self._sync_billing_file_display()
        return warnings

    def _write_json_file(self, path: Path, rows: list[dict[str, str]]) -> None:
        payload = {
            "version": 2,
            **self._json_paths_payload(),
            "rows": rows,
        }
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _suggest_timestamped_name(self) -> str:
        # 例: kintai_results_20260511_1159.json
        ts = self._now_ts()
        return f"kintai_results_{ts}.json"

    def _now_ts(self) -> str:
        # YYYYMMDD_HHMM
        import time

        return time.strftime("%Y%m%d_%H%M")

    def _save_json_via_dialog(self) -> bool:
        """保存ボタン相当の保存処理。保存成功時 True、キャンセル/未保存時 False。"""
        rows = self._current_grid_rows()
        if not rows:
            messagebox.showinfo("保存", "保存する行がありません。")
            return False
        self._prepare_native_dialog()
        initialfile = (
            os.path.basename(str(self._loaded_json_path))
            if self._loaded_json_path
            else self._suggest_timestamped_name()
        )
        fp = filedialog.asksaveasfilename(
            title="結果JSONを保存",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile=initialfile,
            parent=self._root,
        )
        if not fp:
            return False
        out = Path(fp)
        self._write_json_file(out, rows)
        self._loaded_json_path = out
        self._loaded_rows = rows
        self._last_saved_snapshot = self._compute_snapshot(rows)
        self._status_var.set(f"保存しました: {self._loaded_json_path}")
        self._update_title()
        return True

    def _save_json(self) -> None:
        if self._busy:
            return
        self._save_json_via_dialog()

    def _load_json(self) -> None:
        if self._busy:
            return
        self._prepare_native_dialog()
        fp = filedialog.askopenfilename(
            title="結果JSONを読み込み",
            filetypes=[("JSON", "*.json"), ("すべて", "*.*")],
            parent=self._root,
        )
        if not fp:
            return
        data = json.loads(Path(fp).read_text(encoding="utf-8"))
        rows = data.get("rows")
        if not isinstance(rows, list):
            raise ValueError("JSON形式が不正です: rows がありません")
        self._loaded_json_path = Path(fp)
        self._loaded_rows = [r for r in rows if isinstance(r, dict)]
        self._last_saved_snapshot = self._compute_snapshot(self._loaded_rows)

        path_warnings = self._restore_paths_from_json(data)

        self._rebuild_grid_from_rows(self._loaded_rows)
        if not self._busy:
            self._update_data_dir_dependent_buttons()
        ratio_text = self._company_match_ratio_text(self._loaded_rows)
        status = f"読み込みました: {self._loaded_json_path} / {ratio_text}"
        if self._analysis_prerequisites_met():
            status += "（データフォルダ・請求用ファイルを復元）"
        self._status_var.set(status)
        if path_warnings:
            messagebox.showwarning(
                "パス復元",
                "一部のパスを復元できませんでした。\n\n"
                + "\n\n".join(path_warnings),
                parent=self._root,
            )
        self._update_title()

    def _rebuild_grid_from_rows(self, rows: list[dict[str, str]]) -> None:
        self._clear_grid()
        for r in rows:
            vals = self._grid_values_from_row(r)
            iid = self._tree.insert("", tk.END, values=vals)
            self._item_paths[iid] = (r.get("resolved_path") or "").strip()
        self._refresh_row_numbers()

    def _update_title(self) -> None:
        base = "勤務表解析"
        if self._loaded_json_path is None:
            self._root.title(base)
        else:
            self._root.title(f"{base} - {self._loaded_json_path.name}")

    def _has_unsaved_changes(self) -> bool:
        rows = self._current_grid_rows()
        if not rows:
            return False
        cur = self._compute_snapshot(rows)
        return cur != (self._last_saved_snapshot or "")

    def _on_close(self) -> None:
        if self._busy:
            messagebox.showwarning(
                "処理中",
                "解析中は終了できません。中断するか、完了を待ってください。",
            )
            return
        if self._has_unsaved_changes():
            ok = messagebox.askyesno(
                "未保存データ",
                "未保存のグリッドデータがあります。終了前に保存しますか？\n（保存ファイル名は日時付きで自動生成されます）",
            )
            if ok:
                try:
                    if not self._save_json_via_dialog():
                        return
                except Exception as e:
                    messagebox.showerror("保存失敗", str(e))
                    return
        self._close_row_file()
        self._root.destroy()

    def _confirm_clear_grid_for_new_analysis(self) -> bool:
        """グリッドに行があるとき、新規解析でデータが消える旨を確認する。"""
        if not self._tree.get_children():
            return True
        return messagebox.askokcancel(
            "新規解析",
            "グリッドに表示されているデータは、新規解析を開始するとすべて消えます。\n\n"
            "続行しますか？",
            parent=self._root,
        )

    def _start_new_analysis(self, *, chain_error_reanalysis_after: bool = False) -> None:
        if self._busy:
            return
        if not self._confirm_clear_grid_for_new_analysis():
            return
        # 新規解析: グリッドをクリアして最初から
        self._chain_error_reanalysis_after_new = chain_error_reanalysis_after
        self._loaded_rows = []
        self._loaded_json_path = None
        self._last_saved_snapshot = ""
        self._update_title()
        self._start_analysis(base_rows=None)

    def _start_new_analysis_then_error_reanalysis(self) -> None:
        """新規解析の正常完了後に、エラー再解析を続けて実行する（中断時は行わない）。"""
        self._start_new_analysis(chain_error_reanalysis_after=True)

    def _start_continue_analysis(self) -> None:
        # 継続解析: 読み込み済み（または現在表示）の状態を起点
        base = self._loaded_rows or self._current_grid_rows()
        self._start_analysis(base_rows=base)

    def _start_analysis(self, *, base_rows: list[dict[str, str]] | None) -> None:
        if self._busy:
            return
        if not self._analysis_prerequisites_met():
            messagebox.showwarning(
                "解析",
                "解析を実行するには、データフォルダと請求用ファイルの両方を選択してください。",
                parent=self._root,
            )
            return
        aid = self._require_assistant_uid()
        if not aid:
            return
        parallel = self._parallel_workers_value()
        self._busy = True
        self._cancel_event = threading.Event()
        self._new_btn.configure(state=tk.DISABLED)
        self._new_plus_error_btn.configure(state=tk.DISABLED)
        self._cont_btn.configure(state=tk.DISABLED)
        self._selected_reanalysis_btn.configure(state=tk.DISABLED)
        self._save_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._error_reanalysis_btn.configure(state=tk.DISABLED)
        self._billing_prepare_btn.configure(state=tk.DISABLED)
        self._billing_update_btn.configure(state=tk.DISABLED)
        self._progress_var.set("")
        self._status_var.set("解析を準備しています…")

        # 継続解析のときは、まず既存行をグリッドに反映（ユーザ判断も含む）
        if base_rows is None:
            self._clear_grid()
        else:
            self._rebuild_grid_from_rows(base_rows)

        td = self._data_dir.resolve()
        client = self._client

        # 以降は worker スレッドで解析

        def log_line(message: str) -> None:
            def apply_log(m: str = message) -> None:
                if self._should_ignore_status_log(m):
                    return
                self._status_var.set(m[:800])

            self.after(0, apply_log)

        base_done = 0
        if base_rows:
            base_done = len([r for r in base_rows if isinstance(r, dict)])
            self._progress_var.set(f"実行済 {base_done} / 対象 ?")

        progress_state = {"target": base_done}

        def refresh_running_ratio_status() -> None:
            current_rows = self._current_grid_rows()
            target_count = max(progress_state["target"], len(current_rows))
            ratio_text = self._company_match_ratio_text(
                current_rows,
                total_target_count=target_count,
                prefix="会社名比較 〇率(途中経過)",
            )
            self._status_var.set(f"解析中… / {ratio_text}")

        def on_progress(done: int, total: int) -> None:
            # 継続解析時は読み込み済み件数を加算して表示
            def apply_progress(d: int = done, t: int = total, b: int = base_done) -> None:
                progress_state["target"] = b + t
                self._progress_var.set(f"実行済 {b + d} / 対象 {b + t}")
                refresh_running_ratio_status()

            self.after(0, apply_progress)

        def on_row(r: dict[str, str]) -> None:
            def append_row(row: dict[str, str]) -> None:
                vals = self._grid_values_from_row(row)
                iid = self._tree.insert("", tk.END, values=vals)
                self._item_paths[iid] = row.get("resolved_path", "")
                self._refresh_row_numbers()
                try:
                    self._tree.yview_moveto(1)
                except tk.TclError:
                    pass
                refresh_running_ratio_status()

            self.after(0, lambda rr=r: append_row(rr))

        def _as_core_row(r: dict[str, str]) -> dict[str, str]:
            """UI行(dict) -> kintai_core の row(dict) に寄せる"""
            # 継続解析時に既存行の情報を落とさないよう、UI行をできるだけ保持したまま
            # kintai_core が参照するキーへ寄せる。
            out: dict[str, str] = dict(r)
            out["file_name"] = self._file_name_from_row(r)
            out["resolved_path"] = (r.get("resolved_path") or "").strip()
            uj = (
                r.get(self.FINAL_JUDGMENT_COL)
                or r.get(self.LEGACY_USER_JUDGMENT_COL)
                or r.get("user_judgment_company")
                or ""
            ).strip()
            if uj:
                out["user_judgment_company"] = normalize_judgment_symbol(uj)
            ts = (r.get("対象シート有無") or r.get("target_sheet_exists") or "").strip()
            if ts:
                out["target_sheet_exists"] = ts
            return out

        def worker() -> None:
            rows_result: list[dict[str, str]] | None = None
            err: BaseException | None = None

            # 起点行（継続解析）: file_name -> row
            base_map: dict[str, dict[str, str]] = {}
            if base_rows:
                for br in base_rows:
                    cr = _as_core_row(br)
                    fn = (cr.get("file_name") or "").strip()
                    if fn:
                        base_map[fn] = cr

            def on_row_merge(new_row: dict[str, str]) -> None:
                # 手動変更済みのユーザ判断のみ引き継ぐ（未変更時は自動判断を採用）
                fn = (new_row.get("file_name") or "").strip()
                if fn and fn in base_map and is_manual_user_judgment(base_map[fn]):
                    uj = normalize_judgment_symbol(
                        (
                            base_map[fn].get("user_judgment_company")
                            or base_map[fn].get(self.FINAL_JUDGMENT_COL)
                            or base_map[fn].get(self.LEGACY_USER_JUDGMENT_COL)
                            or ""
                        ).strip()
                    )
                    if uj:
                        new_row["user_judgment_company"] = uj
                on_row(new_row)

            try:
                skip_names: set[str] | None = None
                if base_rows:
                    skip_names = set()
                    for br in base_rows:
                        if not isinstance(br, dict):
                            continue
                        fn = self._file_name_from_row(br)
                        if fn:
                            skip_names.add(fn)

                exp_y, exp_m = self._expected_year_month()
                rows_result = run_analysis(
                    client,
                    aid,
                    td,
                    save_md_path=Path.cwd() / "解析結果.md",
                    on_log=log_line,
                    emit_progress_md_rows=False,
                    on_file_progress=on_progress,
                    on_row_completed=on_row_merge,
                    cancel_event=self._cancel_event,
                    skip_file_names=skip_names,
                    parallel_chats=parallel,
                    expected_year=exp_y,
                    expected_month=exp_m,
                )
            except BaseException as e:
                err = e
                rows_result = None

            rows_final = rows_result or []

            # 継続解析: base を起点に、結果を追加（既存は skip_file_names でスキップされる想定）
            merged: dict[str, dict[str, str]] = {k: v for k, v in base_map.items()}
            for nr in rows_final:
                fn = (nr.get("file_name") or "").strip()
                if not fn:
                    continue
                if fn not in merged:
                    merged[fn] = nr

            rows_final = list(merged.values())

            def finish() -> None:
                # --- 共通: busy解除・中断ボタンは常に無効化 ---
                self._busy = False
                self._cancel_btn.configure(state=tk.DISABLED)
                cancelled = self._cancel_event is not None and self._cancel_event.is_set()
                self._cancel_event = None
                chain_error = self._chain_error_reanalysis_after_new
                self._chain_error_reanalysis_after_new = False

                # --- 途中結果の確定（エラー時も含む） ---
                # エラー発生時は worker で rows_result が None になりやすい。
                # しかし UI は on_row_completed で逐次 append 済みなので、そのグリッド内容を
                # 「保存可能な途中結果」として loaded_rows に確定する。
                if cancelled:
                    # 途中までappend済みのグリッドを保存対象にする
                    self._loaded_rows = self._current_grid_rows()
                elif err is not None:
                    # エラー時も同様に、グリッド上の途中結果を保持
                    self._loaded_rows = self._current_grid_rows()
                else:
                    # 正常完了時は merged 結果を確定し、グリッドを作り直す
                    self._loaded_rows = rows_final
                    self._rebuild_grid_from_rows(rows_final)

                # 保存スナップショットも更新（エラー時に保存→その後閉じる、の導線を作る）
                try:
                    self._last_saved_snapshot = self._compute_snapshot(self._loaded_rows)
                except Exception:
                    # snapshot 失敗は致命ではない
                    self._last_saved_snapshot = ""

                # --- UIボタン復帰 ---
                self._update_data_dir_dependent_buttons()
                self._load_btn.configure(state=tk.NORMAL)
                if self._loaded_rows:
                    self._save_btn.configure(state=tk.NORMAL)

                # --- ステータス表示 ---
                n_rows = len(self._loaded_rows or [])
                ratio_text = self._company_match_ratio_text(self._loaded_rows)
                if err is not None:
                    # 要件: 送信エラー等でも中断状態・保存/継続解析を有効にする
                    messagebox.showerror(
                        "解析エラー",
                        str(err),
                    )
                    self._status_var.set(
                        f"エラーで停止しました: {n_rows} 件を保持（JSON保存で続きから再開できます） / {ratio_text}"
                    )
                    # 進捗は固定せず、最後に見えていた値を維持（空にはしない）
                    self._refresh_reanalysis_buttons_state()
                    return

                if cancelled:
                    self._status_var.set(
                        f"中断しました: {n_rows} 件を保持（JSON保存で続きから再開できます） / {ratio_text}"
                    )
                    self._refresh_reanalysis_buttons_state()
                    return

                self._status_var.set(
                    f"完了: {n_rows} 件をグリッド表示（解析結果.md を出力） / {ratio_text}"
                )
                self._refresh_reanalysis_buttons_state()

                if chain_error:
                    self._status_var.set(
                        f"新規解析が完了しました。エラー再解析を開始します… / {ratio_text}"
                    )
                    self.after(0, self._start_error_reanalysis)
                    return

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _on_row_double_click(self, event: tk.Event) -> None:
        cols = list(self._tree["columns"])
        ci = self._column_index_at_event(event)
        if 0 <= ci < len(cols) and cols[ci] == self.FINAL_JUDGMENT_COL:
            return
        if not (0 <= ci < len(cols) and cols[ci] == self.TARGET_FILE_NAME_COL):
            return
        rid = self._tree.identify_row(event.y)
        if not rid:
            return
        path = self._resolve_file_path_for_row(rid)
        if path is None:
            messagebox.showinfo(
                "ファイルを開けません",
                "この行に関連するファイルが見つかりません。\n"
                "データフォルダまたは resolved_path を確認してください。",
            )
            return
        self._tree.selection_set(rid)
        self._open_row_file(rid, force=True)


def _center_window_on_screen(win: tk.Misc) -> None:
    """現在のサイズのまま、プライマリディスプレイのおおよそ中央へ移動する。"""
    win.update_idletasks()
    w = max(win.winfo_width(), win.winfo_reqwidth())
    h = max(win.winfo_height(), win.winfo_reqheight())
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = max(0, (sw - w) // 2)
    y = max(0, (sh - h) // 2)
    win.geometry(f"+{x}+{y}")


def main() -> None:
    root = tk.Tk()
    root.title("勤務表解析")

    client = create_client()
    if not client.authenticate():
        messagebox.showerror(
            "認証エラー",
            "NewtonX の認証に失敗しました。",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    try:
        assistants = client.get_assistants() or []
    except APIError as e:
        err_text = str(e).strip()
        if "403" in err_text:
            user_msg = (
                "アシスタント一覧の取得がサーバーに拒否されました（HTTP 403）。\n"
                "利用アカウントの権限・ロール、またはトークン／セッションの状態を確認してください。\n\n"
                f"{err_text}"
            )
        else:
            user_msg = (
                "アシスタント一覧を取得できませんでした（NewtonX API）。\n\n"
                f"{err_text}"
            )
        messagebox.showerror("NewtonX API エラー", user_msg, parent=root)
        root.destroy()
        sys.exit(1)

    assistant_names = [
        str(a.get("name") or "").strip() for a in assistants if str(a.get("name") or "").strip()
    ]
    if not assistant_names:
        messagebox.showerror(
            "エラー",
            "アシスタント一覧を取得できませんでした。",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    app = KintaiApp(root, client=client, assistants=assistants)
    app.pack(fill=tk.BOTH, expand=True)
    root.update_idletasks()
    _center_window_on_screen(root)
    root.lift()
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
