"""勤務表画像・PDF 解析ツールの Windows GUI エントリ。"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import threading
import ctypes
from ctypes import wintypes
from dataclasses import dataclass
from datetime import date
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

from PIL import ImageTk

_CODE_DIR = Path(__file__).resolve().parent
_ROOT = _CODE_DIR.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from kintai_core import (
    DEFAULT_PARALLEL_ANALYSIS_CHATS,
    EXCEL_SUFFIXES,
    path_is_excluded_archive,
    LEGACY_COMPANY_COL,
    LEGACY_COMPANY_NAME_COL,
    LEGACY_COMPANY_READ_LONG_COL,
    LEGACY_FILE_NAME_COL,
    LEGACY_EMPLOYEE_NO_COL,
    LEGACY_MONTH_COL,
    LEGACY_PERSON_COL,
    LEGACY_TARGET_FILE_NAME_COL,
    LEGACY_YEAR_COL,
    LEGACY_SUFFIX_READ_YEAR_COL,
    LEGACY_SUFFIX_READ_MONTH_COL,
    LEGACY_SUFFIX_READ_COMPANY_COL,
    LEGACY_SUFFIX_READ_PERSON_COL,
    LEGACY_SUFFIX_READ_TOTAL_HOURS_RAW_COL,
    LEGACY_SUFFIX_READ_TRANSPORT_EXPENSE_COL,
    PARALLEL_WORKERS_MAX,
    LEGACY_USER_JUDGMENT_COL,
    LEGACY_BILLING_UPDATE_HOURS_COL,
    SUMMARY_BILLING_UPDATE_RESULT_COL,
    SUMMARY_BILLING_UPDATE_HOURS_COL,
    SUMMARY_BILLING_UPDATE_TRANSPORT_COL,
    SUMMARY_TOTAL_HOURS_RAW_COL,
    SUMMARY_TRANSPORT_EXPENSE_COL,
    SUMMARY_COMPANY_COL,
    SUMMARY_EMPLOYEE_NO_COL,
    SUMMARY_FINAL_JUDGMENT_COL,
    SUMMARY_MONTH_COL,
    SUMMARY_PERSON_COL,
    SUMMARY_MATCH_PERSON_COL,
    SUMMARY_MATCH_DOC_TYPE_COL,
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
    clear_billing_update_hours_column,
    apply_billing_create_for_rows,
    billing_duplicate_group_key,
    build_billing_duplicate_groups,
    BillingDuplicateResolution,
    BillingDuplicateGroupInfo,
    billing_row_update_candidates,
    resolve_billing_hours_from_group,
    resolve_billing_transport_from_group,
    populate_billing_update_columns,
    row_display_values,
    sync_row_ai_read_states,
    SUMMARY_LABOR_AI_READ_STATE_COL,
    SUMMARY_TRANSPORT_AI_READ_STATE_COL,
    AI_READ_STATE_COLUMN_HEADING,
    run_analysis,
    summary_header_cells,
    BillingTsWriteItem,
    BILLING_UPDATE_SKIP,
    apply_billing_ts_writes,
    merge_billing_ts_plan_results,
    plan_billing_ts_writes,
    update_billing_engineer_ts_sheet,
    add_company_alias,
    company_alias_lookup_key,
    default_company_alias_table_path,
    default_user_settings_path,
    load_user_settings,
    save_last_assistant_name,
    load_company_aliases,
    is_match_company_ok_for_ratio,
    is_match_person_manual,
    match_person_symbol_for_row,
    recalculate_match_company_for_row,
    recalculate_match_doc_type_for_row,
    recalculate_match_person_for_row,
    rename_file_to_excluded,
    rename_file_from_excluded,
    list_excluded_files,
    restored_file_name_from_excluded,
    build_row_after_restore_without_analysis,
    save_company_aliases,
    _parse_filename_company_and_person,
    _is_billing_aggregated_marker,
    _row_billing_update_hours_decimal,
    _row_billing_update_transport,
    _normalize_billing_update_copy,
    _normalize_billing_update_hours_copy,
    _normalize_billing_update_transport_copy,
    _billing_value_is_no_data,
    _set_billing_update_columns,
    _company_text_contains_seraku,
    _document_company_for_display,
    _is_valid_employee_no,
    _row_employee_no,
    _normalize_employee_no_cell_value,
    _row_file_name,
    _row_grid_no,
    _row_total_hours_decimal,
    _row_transport_expense_raw,
    load_billing_preview_pages,
    IMAGE_SUFFIXES,
    PDF_SUFFIX,
)
from newtonx_adk.exceptions import APIError


@dataclass(frozen=True)
class BillingReviewEntry:
    """請求前照合セッション用の行スナップショット。"""

    iid: str
    file_name: str
    grid_no: int
    employee_no: str
    hours_write: str
    transport_write: str

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
    MATCH_PERSON_COL = SUMMARY_MATCH_PERSON_COL
    MATCH_DOC_TYPE_COL = SUMMARY_MATCH_DOC_TYPE_COL
    TOTAL_HOURS_DECIMAL_COL = "合計勤務時間（10進）"
    TOTAL_HOURS_RAW_COL = SUMMARY_TOTAL_HOURS_RAW_COL
    TRANSPORT_EXPENSE_COL = SUMMARY_TRANSPORT_EXPENSE_COL
    LABOR_AI_READ_STATE_COL = SUMMARY_LABOR_AI_READ_STATE_COL
    TRANSPORT_AI_READ_STATE_COL = SUMMARY_TRANSPORT_AI_READ_STATE_COL
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
        MATCH_COMPANY_COL: 100,
        MATCH_PERSON_COL: 72,
        MATCH_DOC_TYPE_COL: 72,
        LABOR_AI_READ_STATE_COL: 88,
        TRANSPORT_AI_READ_STATE_COL: 88,
        "押印有無": 72,
    }
    _TAG_REANALYSIS_ACTIVE = "reanalysis_active"
    _TAG_BILLING_REVIEW_ACTIVE = "billing_review_active"
    _TAG_DUP_EMPLOYEE_NO = "dup_employee_no"
    _ROW_EXTRA_KEYS = (
        "name_company_from_file",
        "name_company_from_doc",
        "match_company_manual_ok",
        "match_person_manual",
    )

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
        self._user_settings_path = default_user_settings_path(_ROOT)
        default_name = self._initial_assistant_name(assistant_names)
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
        self._billing_review_photos: list[ImageTk.PhotoImage] = []
        self._billing_review_highlighted_iids: list[str] = []
        self._cancel_event: threading.Event | None = None

        self._loaded_rows: list[dict[str, str]] = []
        self._loaded_json_path: Path | None = None
        self._last_saved_snapshot: str = ""
        self._chain_error_reanalysis_after_new = False
        self._sort_column: str | None = None
        self._sort_reverse: bool = False
        self._tree_column_headings: tuple[str, ...] = ()
        self._grid_row_no_seq: int = 0
        self._company_alias_table_path = default_company_alias_table_path(_ROOT)
        self._company_aliases: list[dict[str, str]] = load_company_aliases(
            self._company_alias_table_path
        )
        self._row_extra: dict[str, dict[str, str]] = {}

        self._build_ui()

    def _tree_heading_font(self) -> tkfont.Font:
        spec = ttk.Style().lookup("Treeview.Heading", "font")
        if spec:
            return tkfont.Font(root=self._root, font=spec)
        return tkfont.nametofont("TkDefaultFont")

    def _tree_body_font(self) -> tkfont.Font:
        spec = ttk.Style().lookup("Treeview", "font")
        if spec:
            return tkfont.Font(root=self._root, font=spec)
        return tkfont.nametofont("TkDefaultFont")

    def _column_width_for_cell_text(
        self, text: str, *, font: tkfont.Font | None = None
    ) -> int:
        """セル文字列の描画幅に合わせた列幅（px）を返す（見出しは含めない）。"""
        f = font or self._tree_body_font()
        padding_px = 20
        min_px = 48
        return max(f.measure(str(text or "")) + padding_px, min_px)

    def _autofit_grid_column_widths(self) -> None:
        """表示中のデータ行に合わせて各列幅を調整する（見出し幅は使わない）。"""
        cols = list(self._tree["columns"])
        if not cols:
            return
        font = self._tree_body_font()
        max_px = [48] * len(cols)
        for iid in self._tree.get_children():
            vals = tuple(self._tree.item(iid, "values") or ())
            for i, _h in enumerate(cols):
                cell = vals[i] if i < len(vals) else ""
                max_px[i] = max(
                    max_px[i], self._column_width_for_cell_text(cell, font=font)
                )
        for h, w_px in zip(cols, max_px, strict=True):
            self._tree.column(h, width=w_px, minwidth=48, stretch=tk.NO)
        self.update_idletasks()
        if self._grid_y_scroll is not None:
            new_w = self._window_width_for_columns(max_px, y_scroll=self._grid_y_scroll)
            h = self._root.winfo_height()
            if h <= 1:
                h = 680
            self._root.geometry(f"{new_w}x{h}")

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

    @staticmethod
    def _initial_assistant_name(assistant_names: list[str]) -> str:
        path = default_user_settings_path(_ROOT)
        saved = (load_user_settings(path).get("last_assistant_name") or "").strip()
        if saved and saved in assistant_names:
            return saved
        if TARGET_ASSISTANT_NAME in assistant_names:
            return TARGET_ASSISTANT_NAME
        return assistant_names[0] if assistant_names else ""

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

        analysis_lf = ttk.LabelFrame(ctrl, text="解析", padding=(6, 4))
        analysis_lf.pack(side=tk.LEFT, anchor="n", padx=(0, 8))

        self._new_btn = ttk.Button(
            analysis_lf,
            text="新規解析",
            command=self._start_new_analysis,
            state=tk.DISABLED,
        )
        self._new_btn.pack(side=tk.LEFT)

        self._new_plus_error_btn = ttk.Button(
            analysis_lf,
            text="新規解析＋エラー再解析",
            command=self._start_new_analysis_then_error_reanalysis,
            state=tk.DISABLED,
        )
        self._new_plus_error_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._cont_btn = ttk.Button(
            analysis_lf,
            text="追加継続解析",
            command=self._start_continue_analysis,
            state=tk.DISABLED,
        )
        self._cont_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._selected_reanalysis_btn = ttk.Button(
            analysis_lf,
            text="選択行解析",
            command=self._start_selected_rows_reanalysis,
            state=tk.DISABLED,
        )
        self._selected_reanalysis_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._error_reanalysis_btn = ttk.Button(
            analysis_lf,
            text="エラー再解析",
            command=self._start_error_reanalysis,
            state=tk.DISABLED,
        )
        self._error_reanalysis_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._cancel_btn = ttk.Button(
            analysis_lf,
            text="中断",
            command=self._cancel_analysis,
            state=tk.DISABLED,
        )
        self._cancel_btn.pack(side=tk.LEFT, padx=(8, 0))

        billing_lf = ttk.LabelFrame(ctrl, text="請求", padding=(6, 4))
        billing_lf.pack(side=tk.LEFT, anchor="n", padx=(0, 8))

        self._billing_prepare_btn = ttk.Button(
            billing_lf,
            text="請求データ作成",
            command=self._create_billing_data,
            state=tk.DISABLED,
        )
        self._billing_prepare_btn.pack(side=tk.LEFT)

        self._billing_delete_btn = ttk.Button(
            billing_lf,
            text="請求データ削除",
            command=self._delete_billing_data,
            state=tk.DISABLED,
        )
        self._billing_delete_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._billing_review_btn = ttk.Button(
            billing_lf,
            text="画像とデータ照合",
            command=self._open_billing_image_data_review,
            state=tk.DISABLED,
        )
        self._billing_review_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._billing_update_btn = ttk.Button(
            billing_lf,
            text="請求ファイル更新",
            command=self._update_billing_file,
            state=tk.DISABLED,
        )
        self._billing_update_btn.pack(side=tk.LEFT, padx=(4, 0))

        data_lf = ttk.LabelFrame(ctrl, text="データ", padding=(6, 4))
        data_lf.pack(side=tk.LEFT, anchor="n", padx=(0, 8))

        self._save_btn = ttk.Button(
            data_lf, text="保存", command=self._save_json, state=tk.DISABLED
        )
        self._save_btn.pack(side=tk.LEFT)

        self._load_btn = ttk.Button(
            data_lf, text="読み込み", command=self._load_json, state=tk.NORMAL
        )
        self._load_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._restore_excluded_btn = ttk.Button(
            data_lf,
            text="削除データ復帰",
            command=self._prompt_restore_excluded_files,
            state=tk.DISABLED,
        )
        self._restore_excluded_btn.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(
            data_lf,
            text="グリッド幅調整",
            command=self._autofit_grid_column_widths,
        ).pack(side=tk.LEFT, padx=(4, 0))

        self._export_csv_btn = ttk.Button(
            data_lf,
            text="CSVエクスポート",
            command=self._export_grid_csv,
            state=tk.DISABLED,
        )
        self._export_csv_btn.pack(side=tk.LEFT, padx=(4, 0))

        self._progress_var = tk.StringVar(value="")
        self._status_var = tk.StringVar(value="準備完了")

        status_area = ttk.Frame(ctrl)
        status_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
        # 「実行済 100 / 対象 120」など3桁になっても欠けないよう、表示幅を広げる
        ttk.Label(status_area, textvariable=self._progress_var, width=30).pack(
            side=tk.LEFT, anchor="n", padx=(0, 12)
        )
        self._status_label = ttk.Label(
            status_area,
            textvariable=self._status_var,
            anchor="w",
            width=70,
            wraplength=900,
            justify="left",
        )
        self._status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, anchor="n")

        grid_frame = ttk.Frame(self, padding=(8, 0, 8, 8))
        grid_frame.pack(fill=tk.BOTH, expand=True)

        headings = summary_header_cells()
        self._tree_column_headings = headings
        y_scroll = ttk.Scrollbar(grid_frame)
        self._grid_y_scroll = y_scroll
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
        self._tree.tag_configure(
            self._TAG_BILLING_REVIEW_ACTIVE,
            background="#1f2937",
            foreground="#ffffff",
        )
        self._tree.tag_configure(self._TAG_DUP_EMPLOYEE_NO, foreground="#c00000")
        y_scroll.configure(command=self._tree.yview)
        x_scroll.configure(command=self._tree.xview)

        for w_px, h in zip(col_px, headings, strict=True):
            self._tree.column(
                h, width=w_px, minwidth=48, stretch=tk.NO, anchor="w"
            )
            self._tree.heading(
                h,
                text=self._column_heading_display_text(h),
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
        self._tree.bind("<Control-a>", self._on_tree_select_all)
        self._tree.bind("<Control-A>", self._on_tree_select_all)

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
        has_grid_rows = bool(self._tree.get_children())
        self._save_btn.configure(
            state=(tk.NORMAL if has_grid_rows else tk.DISABLED)
        )
        self._export_csv_btn.configure(
            state=(tk.NORMAL if has_grid_rows else tk.DISABLED)
        )
        self._cancel_btn.configure(state=tk.DISABLED)
        has_data_dir = (
            self._data_dir is not None and self._data_dir.is_dir()
        )
        self._restore_excluded_btn.configure(
            state=(tk.NORMAL if has_data_dir else tk.DISABLED)
        )
        self._refresh_reanalysis_buttons_state()

    def _refresh_billing_buttons_state(self) -> None:
        if self._busy:
            self._billing_prepare_btn.configure(state=tk.DISABLED)
            self._billing_delete_btn.configure(state=tk.DISABLED)
            self._billing_review_btn.configure(state=tk.DISABLED)
            self._billing_update_btn.configure(state=tk.DISABLED)
            return
        has_rows = bool(self._tree.get_children())
        self._billing_prepare_btn.configure(
            state=(tk.NORMAL if has_rows else tk.DISABLED)
        )
        self._billing_delete_btn.configure(
            state=(tk.NORMAL if has_rows else tk.DISABLED)
        )
        self._billing_review_btn.configure(
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

    def _on_tree_select_all(self, _event: tk.Event) -> str:
        """Ctrl+A でグリッドの全行を選択する。"""
        children = self._tree.get_children()
        if children:
            self._tree.selection_set(*children)
            self._tree.focus(children[0])
        return "break"

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

    def _billing_target_iids(self) -> list[str]:
        selected = self._selected_tree_iids()
        if selected:
            return selected
        return list(self._tree.get_children())

    def _grid_export_target_iids(self) -> list[str]:
        selected = self._selected_tree_iids()
        if selected:
            return selected
        return list(self._tree.get_children())

    def _billing_target_scope_label(self) -> str:
        if self._selected_tree_iids():
            return f"選択行（{len(self._billing_target_iids())} 行）"
        return "全行"

    def _delete_target_iids(self, rid: str) -> list[str]:
        selected = self._selected_tree_iids()
        if selected:
            return list(selected)
        return [rid] if rid else []

    def _prompt_delete_grid_rows(self, rid: str) -> None:
        iids = self._delete_target_iids(rid)
        if not iids:
            return
        names: list[str] = []
        for iid in iids:
            row = self._current_row_dict_from_iid(iid)
            fn = self._file_name_from_row(row)
            names.append(fn or "（ファイル名なし）")
        preview = "\n".join(names[:15])
        if len(names) > 15:
            preview += f"\n... 他 {len(names) - 15} 件"
        if not messagebox.askokcancel(
            "行を削除",
            f"対象: {len(iids)} 行\n\n"
            f"{preview}\n\n"
            "グリッドから削除します。ファイル名の末尾に .bak を付け、"
            "再解析の対象外にします（例: 勤務表.pdf → 勤務表.pdf.bak）。\n"
            "元に戻す場合は「削除データ復帰」ボタンを使用してください。",
            parent=self._root,
        ):
            return
        self._delete_grid_rows(iids)

    def _delete_grid_rows(self, iids: list[str]) -> None:
        to_remove: list[str] = []
        rename_ok = 0
        missing = 0
        already_excluded = 0
        failed: list[str] = []

        for iid in iids:
            path = self._resolve_file_path_for_row(iid)
            row = self._current_row_dict_from_iid(iid)
            fn = self._file_name_from_row(row) or iid
            if path is None:
                missing += 1
                to_remove.append(iid)
                continue
            if path_is_excluded_archive(path):
                already_excluded += 1
                to_remove.append(iid)
                continue
            try:
                rename_file_to_excluded(path)
                rename_ok += 1
                to_remove.append(iid)
            except FileExistsError:
                failed.append(f"{fn}: リネーム先（末尾 .bak）が既に存在します")
            except OSError as e:
                failed.append(f"{fn}: {e}")

        for iid in to_remove:
            if iid == self._preview_row_iid:
                self._cancel_preview_hover_timer()
                self._close_row_file()
            self._row_extra.pop(iid, None)
            self._item_paths.pop(iid, None)
            try:
                self._tree.delete(iid)
            except tk.TclError:
                pass

        if to_remove:
            self._renumber_grid_rows()
            self._loaded_rows = self._current_grid_rows()
            ratio_text = self._company_match_ratio_text(self._loaded_rows)
            self._status_var.set(
                f"行削除完了: {len(to_remove)} 行 / {ratio_text}"
            )
            self._refresh_reanalysis_buttons_state()
            self._refresh_billing_buttons_state()

        lines = [f"グリッドから削除: {len(to_remove)} 行"]
        if rename_ok:
            lines.append(f"ファイル名末尾に .bak を付与: {rename_ok} 件")
        if already_excluded:
            lines.append(f"既に除外済み（グリッドのみ削除）: {already_excluded} 件")
        if missing:
            lines.append(f"ファイル不在（グリッドのみ削除）: {missing} 件")
        if failed:
            lines.append(f"リネーム失敗（行は残しました）: {len(failed)} 件")
            for msg in failed[:8]:
                lines.append(f" ・{msg}")
            if len(failed) > 8:
                lines.append(f"  ... 他 {len(failed) - 8} 件")
        if to_remove and self._tree.get_children():
            lines.append("残り行の No を 1 から振り直しました。")
        lines.append("")
        lines.append(
            "請求用列を直すときは「請求データ作成」を再実行してください。"
        )
        if failed:
            messagebox.showwarning(
                "行を削除（一部失敗）", "\n".join(lines), parent=self._root
            )
        elif to_remove:
            messagebox.showinfo("行を削除", "\n".join(lines), parent=self._root)

    def _grid_file_names_set(self) -> set[str]:
        names: set[str] = set()
        for iid in self._tree.get_children():
            row = self._current_row_dict_from_iid(iid)
            fn = self._file_name_from_row(row)
            if fn:
                names.add(fn)
        return names

    def _insert_minimal_restored_row(self, restored_path: Path) -> str:
        """復帰直後のプレースホルダ行をグリッド末尾に追加し iid を返す。"""
        row: dict[str, str] = {
            "upload_ok": "",
            "file_name": restored_path.name,
            "resolved_path": str(restored_path.resolve()),
            "analysis": "",
        }
        row["user_judgment_company"] = auto_judgment_symbol(row)
        vals = self._grid_values_from_row(row)
        iid = self._tree.insert("", tk.END, values=vals)
        self._item_paths[iid] = row["resolved_path"]
        self._sync_row_extra_from_row(iid, row)
        self._assign_row_no_to_iid(iid)
        return iid

    def _prompt_restore_excluded_files(self) -> None:
        if self._busy:
            return
        if self._data_dir is None or not self._data_dir.is_dir():
            messagebox.showwarning(
                "削除データ復帰",
                "データフォルダを選択してください。",
                parent=self._root,
            )
            return
        bak_paths = list_excluded_files(self._data_dir)
        if not bak_paths:
            messagebox.showinfo(
                "削除データ復帰",
                "削除済みファイルはありません。",
                parent=self._root,
            )
            return

        top = tk.Toplevel(self._root)
        top.title("削除データ復帰")
        top.transient(self._root)
        top.geometry("720x420")
        top.minsize(480, 280)

        outer = ttk.Frame(top, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            outer,
            text=(
                "データフォルダ直下の .bak ファイルです。"
                "復帰する行を選択してください（複数可）。"
            ),
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        list_wrap = ttk.Frame(outer)
        list_wrap.pack(fill=tk.BOTH, expand=True)
        y_scroll = ttk.Scrollbar(list_wrap, orient=tk.VERTICAL)
        listbox = tk.Listbox(
            list_wrap,
            selectmode=tk.EXTENDED,
            yscrollcommand=y_scroll.set,
            height=14,
        )
        y_scroll.config(command=listbox.yview)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        display_names = [restored_file_name_from_excluded(p) for p in bak_paths]
        for name in display_names:
            listbox.insert(tk.END, name)

        reanalyze_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            outer,
            text="復帰後に再解析する",
            variable=reanalyze_var,
        ).pack(anchor="w", pady=(8, 0))

        btns = ttk.Frame(outer)
        btns.pack(fill=tk.X, pady=(12, 0))

        def on_cancel() -> None:
            top.destroy()

        def on_restore() -> None:
            sel = list(listbox.curselection())
            if not sel:
                messagebox.showwarning(
                    "削除データ復帰",
                    "復帰するファイルを1件以上選択してください。",
                    parent=top,
                )
                return
            reanalyze = reanalyze_var.get()
            if reanalyze and not self._analysis_prerequisites_met():
                messagebox.showwarning(
                    "削除データ復帰",
                    "再解析するには、データフォルダと請求用ファイルの両方を"
                    "選択してください。\n"
                    "再解析しない場合はチェックを外してください。",
                    parent=top,
                )
                return
            chosen = [bak_paths[i] for i in sel]
            top.destroy()
            self._restore_excluded_files(chosen, reanalyze=reanalyze)

        ttk.Button(btns, text="復帰", command=on_restore).pack(side=tk.RIGHT)
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(
            side=tk.RIGHT, padx=(0, 8)
        )
        top.bind("<Escape>", lambda _e: on_cancel())
        top.grab_set()
        listbox.focus_set()

    def _restore_excluded_files(
        self, bak_paths: list[Path], *, reanalyze: bool
    ) -> None:
        if self._busy:
            return
        grid_names = self._grid_file_names_set()
        restored_iids: dict[str, str] = {}
        skipped_duplicate: list[str] = []
        failed: list[str] = []
        restored_count = 0

        for bak_path in bak_paths:
            restored_name = restored_file_name_from_excluded(bak_path)
            if restored_name in grid_names:
                skipped_duplicate.append(restored_name)
                continue
            try:
                restored_path = rename_file_from_excluded(bak_path)
            except (FileNotFoundError, FileExistsError, ValueError, OSError) as e:
                failed.append(f"{bak_path.name}: {e}")
                continue
            iid = self._insert_minimal_restored_row(restored_path)
            restored_iids[restored_name] = iid
            grid_names.add(restored_name)
            restored_count += 1

        if restored_count:
            self._renumber_grid_rows()
            self._loaded_rows = self._current_grid_rows()

        exp_y, exp_m = self._expected_year_month()

        if restored_count and not reanalyze:
            for fn, iid in restored_iids.items():
                row = build_row_after_restore_without_analysis(
                    Path(self._item_paths.get(iid, "")),
                    company_aliases=self._company_aliases,
                    expected_year=exp_y,
                    expected_month=exp_m,
                )
                row["user_judgment_company"] = auto_judgment_symbol(row)
                self._replace_row_with_result(iid, row)
            self._loaded_rows = self._current_grid_rows()
            ratio_text = self._company_match_ratio_text(self._loaded_rows)
            self._status_var.set(
                f"削除データ復帰完了: {restored_count} 行 / {ratio_text}"
            )
            self._refresh_reanalysis_buttons_state()
            self._refresh_billing_buttons_state()

        lines: list[str] = []
        if restored_count:
            lines.append(f"グリッドに復帰: {restored_count} 行")
            lines.append("No を表示順に振り直しました。")
        else:
            lines.append("復帰できた行はありません。")
        if skipped_duplicate:
            lines.append(f"グリッドに同名あり（スキップ）: {len(skipped_duplicate)} 件")
            for name in skipped_duplicate[:8]:
                lines.append(f" ・{name}")
            if len(skipped_duplicate) > 8:
                lines.append(f"  ... 他 {len(skipped_duplicate) - 8} 件")
        if failed:
            lines.append(f"リネーム失敗: {len(failed)} 件")
            for msg in failed[:8]:
                lines.append(f" ・{msg}")
            if len(failed) > 8:
                lines.append(f"  ... 他 {len(failed) - 8} 件")

        if reanalyze and restored_iids:
            if lines:
                messagebox.showinfo(
                    "削除データ復帰",
                    "\n".join(lines) + "\n\n再解析を開始します。",
                    parent=self._root,
                )
            self._run_targeted_reanalysis(
                restored_iids,
                label="削除データ復帰",
                complete_message=(
                    "選択した削除データの復帰と再解析が完了しました。"
                ),
            )
            return

        if restored_count or skipped_duplicate or failed:
            title = "削除データ復帰"
            if failed and not restored_count:
                messagebox.showwarning(title, "\n".join(lines), parent=self._root)
            else:
                messagebox.showinfo(title, "\n".join(lines), parent=self._root)

    def _billing_cell_preview(self, value: str) -> str:
        t = (value or "").strip()
        return t if t else "（空）"

    def _collect_billing_review_entries(self) -> list[BillingReviewEntry]:
        entries: list[BillingReviewEntry] = []
        for iid in self._billing_target_iids():
            ui_row = self._current_row_dict_from_iid(iid)
            if self._final_judgment_symbol_from_row(ui_row) != "〇":
                continue
            core = self._row_dict_to_core(ui_row)
            hours_val = _row_billing_update_hours_decimal(core)
            transport_val = _row_billing_update_transport(core)
            if _is_billing_aggregated_marker(hours_val):
                continue
            if not (hours_val or "").strip() and not (transport_val or "").strip():
                continue
            entries.append(
                BillingReviewEntry(
                    iid=iid,
                    file_name=self._file_name_from_row(ui_row),
                    grid_no=_row_grid_no(core),
                    employee_no=_row_employee_no(core),
                    hours_write=(hours_val or "").strip(),
                    transport_write=(transport_val or "").strip(),
                )
            )
        return entries

    def _billing_review_preview_iids(self, rep_iid: str) -> list[str]:
        """照合プレビュー用。重複グループなら同一キーの全行 iid（No 昇順）。"""
        rep_core = self._row_dict_to_core(self._current_row_dict_from_iid(rep_iid))
        key = billing_duplicate_group_key(rep_core)
        if key is None:
            return [rep_iid]
        matched: list[tuple[int, int, str]] = []
        for order, iid in enumerate(self._billing_target_iids()):
            core = self._row_dict_to_core(self._current_row_dict_from_iid(iid))
            if billing_duplicate_group_key(core) != key:
                continue
            matched.append((_row_grid_no(core), order, iid))
        if len(matched) < 2:
            return [rep_iid]
        matched.sort(key=lambda t: (t[0], t[1]))
        return [iid for _, _, iid in matched]

    def _billing_review_file_names_display(self, rep_iid: str) -> str:
        iids = self._billing_review_preview_iids(rep_iid)
        if len(iids) <= 1:
            ui_row = self._current_row_dict_from_iid(rep_iid)
            return self._file_name_from_row(ui_row) or "（なし）"
        lines: list[str] = []
        for iid in iids:
            ui_row = self._current_row_dict_from_iid(iid)
            core = self._row_dict_to_core(ui_row)
            no = _row_grid_no(core)
            fn = self._file_name_from_row(ui_row) or "（なし）"
            lines.append(f"No.{no}: {fn}")
        return "\n".join(lines)

    def _set_row_billing_review_highlight(self, rid: str, active: bool) -> None:
        try:
            cur = tuple(self._tree.item(rid, "tags") or ())
        except tk.TclError:
            return
        tag = self._TAG_BILLING_REVIEW_ACTIVE
        if active:
            if tag not in cur:
                self._tree.item(rid, tags=cur + (tag,))
        else:
            if tag in cur:
                self._tree.item(rid, tags=tuple(t for t in cur if t != tag))

    def _clear_billing_review_grid_highlight(self) -> None:
        for rid in self._billing_review_highlighted_iids:
            self._set_row_billing_review_highlight(rid, False)
        self._billing_review_highlighted_iids = []

    def _apply_billing_review_grid_highlight(self, iids: list[str]) -> None:
        self._clear_billing_review_grid_highlight()
        for rid in iids:
            self._set_row_billing_review_highlight(rid, True)
        self._billing_review_highlighted_iids = list(iids)
        if iids:
            try:
                self._tree.see(iids[0])
            except tk.TclError:
                pass

    _BILLING_REVIEW_EXCEL_CANVAS_MSG = (
        "Excel を開きました。内容は Excel ウィンドウで確認してください。"
    )

    def _open_excel_files_for_billing_review(self, iids: list[str]) -> None:
        """照合件表示時にグループ内の Excel を既定アプリで開く。"""
        self._close_row_file()
        tracked = False
        for iid in iids:
            path = self._resolve_file_path_for_row(iid)
            if path is None or path.suffix.lower() not in EXCEL_SUFFIXES:
                continue
            try:
                if not tracked:
                    self._open_row_file(iid, force=True)
                    tracked = True
                else:
                    _win_shell_open_file(path)
            except OSError as e:
                messagebox.showerror(
                    "起動できませんでした",
                    str(e),
                    parent=self._root,
                )

    def _open_billing_image_data_review(self) -> None:
        if self._busy:
            return
        initial = self._collect_billing_review_entries()
        if not initial:
            scope = self._billing_target_scope_label()
            messagebox.showinfo(
                "画像とデータ照合",
                f"対象: {scope}\n\n"
                "最終判断が「〇」かつ更新用列（勤務時間または交通費）が"
                "設定された行がありません。",
                parent=self._root,
            )
            return

        entries: list[BillingReviewEntry] = list(initial)
        state = {"index": 0}

        top = tk.Toplevel(self._root)
        top.title("画像とデータ照合")
        top.minsize(720, 520)
        try:
            top.state("zoomed")
        except tk.TclError:
            top.geometry("960x720")
        top.grab_set()

        outer = ttk.Frame(top, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            outer,
            text=(
                "請求 Excel に書き込む更新用の値と、元ファイルの画像を照合します。"
                "（空欄は Excel では空セルになります）"
            ),
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        body = ttk.Panedwindow(outer, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        img_frame = ttk.LabelFrame(body, text="プレビュー", padding=4)
        body.add(img_frame, weight=3)
        img_inner = ttk.Frame(img_frame)
        img_inner.pack(fill=tk.BOTH, expand=True)
        preview_scroll = ttk.Frame(img_inner)
        preview_scroll.pack(fill=tk.BOTH, expand=True, padx=4, pady=(4, 0))
        preview_canvas = tk.Canvas(preview_scroll, highlightthickness=0)
        preview_y_scroll = ttk.Scrollbar(
            preview_scroll, orient=tk.VERTICAL, command=preview_canvas.yview
        )
        preview_x_scroll = ttk.Scrollbar(
            preview_scroll, orient=tk.HORIZONTAL, command=preview_canvas.xview
        )
        preview_canvas.configure(
            xscrollcommand=preview_x_scroll.set,
            yscrollcommand=preview_y_scroll.set,
        )
        preview_canvas.grid(row=0, column=0, sticky="nsew")
        preview_y_scroll.grid(row=0, column=1, sticky="ns")
        preview_x_scroll.grid(row=1, column=0, sticky="ew")
        preview_scroll.rowconfigure(0, weight=1)
        preview_scroll.columnconfigure(0, weight=1)

        def _on_preview_wheel(event: tk.Event) -> None:
            delta = getattr(event, "delta", 0)
            if not delta:
                return
            steps = int(-1 * (delta / 120))
            if event.state & 0x1:
                preview_canvas.xview_scroll(steps, "units")
            else:
                preview_canvas.yview_scroll(steps, "units")

        def _on_preview_wheel_up(_event: tk.Event) -> None:
            preview_canvas.yview_scroll(-1, "units")

        def _on_preview_wheel_down(_event: tk.Event) -> None:
            preview_canvas.yview_scroll(1, "units")

        for _wheel_widget in (preview_canvas, preview_scroll):
            _wheel_widget.bind("<MouseWheel>", _on_preview_wheel)
            _wheel_widget.bind("<Button-4>", _on_preview_wheel_up)
            _wheel_widget.bind("<Button-5>", _on_preview_wheel_down)

        def _unbind_preview_wheel() -> None:
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                try:
                    top.unbind_all(seq)
                except tk.TclError:
                    pass

        top.bind_all("<MouseWheel>", _on_preview_wheel)
        top.bind_all("<Button-4>", _on_preview_wheel_up)
        top.bind_all("<Button-5>", _on_preview_wheel_down)

        open_file_btn = ttk.Button(img_inner, text="ファイルを開く")
        open_file_btn.pack(pady=(4, 4))

        info_frame = ttk.LabelFrame(body, text="書き込みデータ", padding=8)
        body.add(info_frame, weight=2)

        nav_var = tk.StringVar(value="")
        ttk.Label(info_frame, textvariable=nav_var, font=("", 10, "bold")).pack(
            anchor="w", pady=(0, 8)
        )
        no_var = tk.StringVar()
        file_var = tk.StringVar()
        emp_var = tk.StringVar()
        hours_var = tk.StringVar()
        transport_var = tk.StringVar()
        for label, var in (
            ("No", no_var),
            ("ファイル名", file_var),
            ("社員番号", emp_var),
        ):
            row_f = ttk.Frame(info_frame)
            row_f.pack(fill=tk.X, pady=2)
            ttk.Label(row_f, text=f"{label}:", width=28).pack(side=tk.LEFT)
            ttk.Label(row_f, textvariable=var, wraplength=320, justify="left").pack(
                side=tk.LEFT, fill=tk.X, expand=True
            )
        for label, var in (
            (self.BILLING_UPDATE_HOURS_COL, hours_var),
            (self.BILLING_UPDATE_TRANSPORT_COL, transport_var),
        ):
            row_f = ttk.Frame(info_frame)
            row_f.pack(fill=tk.X, pady=2)
            ttk.Label(row_f, text=f"{label}:", width=28).pack(side=tk.LEFT)
            ttk.Entry(row_f, textvariable=var, width=24).pack(
                side=tk.LEFT, fill=tk.X, expand=True
            )
        apply_grid_btn = ttk.Button(info_frame, text="グリッド更新")
        apply_grid_btn.pack(anchor="w", pady=(8, 0))

        nav_row = ttk.Frame(outer)
        nav_row.pack(fill=tk.X, pady=(8, 0))
        prev_btn = ttk.Button(nav_row, text="前へ")
        prev_btn.pack(side=tk.LEFT)
        next_btn = ttk.Button(nav_row, text="次へ")
        next_btn.pack(side=tk.LEFT, padx=(8, 0))

        action_row = ttk.Frame(outer)
        action_row.pack(fill=tk.X, pady=(8, 0))
        exclude_btn = ttk.Button(action_row, text="最終判断を ✖ にする")
        exclude_btn.pack(side=tk.LEFT)

        def current_entry() -> BillingReviewEntry:
            return entries[state["index"]]

        def _review_fields_dirty() -> bool:
            entry = current_entry()
            return (
                hours_var.get().strip() != (entry.hours_write or "").strip()
                or transport_var.get().strip() != (entry.transport_write or "").strip()
            )

        def _confirm_discard_if_dirty() -> bool:
            if not _review_fields_dirty():
                return True
            answer = messagebox.askyesnocancel(
                "未反映の編集",
                "グリッドに反映していない変更があります。\n破棄して続行しますか？",
                parent=top,
            )
            return answer is True

        def _close_review_window() -> None:
            if not _confirm_discard_if_dirty():
                return
            _unbind_preview_wheel()
            self._close_row_file()
            self._clear_billing_review_grid_highlight()
            top.destroy()

        ttk.Button(action_row, text="閉じる", command=_close_review_window).pack(
            side=tk.RIGHT
        )
        top.protocol("WM_DELETE_WINDOW", _close_review_window)

        _PREVIEW_PAD_X = 8
        _PREVIEW_TEXT_WIDTH = 520
        _PREVIEW_CAPTION_H = 22
        _PREVIEW_MSG_H = 56
        _PREVIEW_SECTION_GAP = 16
        _PREVIEW_PAGE_CAPTION_H = 18
        _PREVIEW_PAGE_GAP = 8

        def _clear_preview_canvas() -> None:
            preview_canvas.delete("all")
            self._billing_review_photos = []
            preview_canvas.configure(scrollregion=(0, 0, 0, 0))

        def _show_preview_message(text: str) -> None:
            _clear_preview_canvas()

            def _redraw(_event: tk.Event | None = None) -> None:
                preview_canvas.delete("all")
                w = preview_canvas.winfo_width()
                h = preview_canvas.winfo_height()
                if w <= 1 or h <= 1:
                    return
                preview_canvas.create_text(
                    w / 2,
                    h / 2,
                    text=text,
                    anchor="center",
                    width=max(w - 24, 80),
                    justify="center",
                )
                preview_canvas.configure(scrollregion=(0, 0, w, h))

            preview_canvas.bind("<Configure>", _redraw)
            _redraw()

        def refresh_preview_image(entry: BillingReviewEntry) -> None:
            _show_preview_message("読み込み中…")
            top.update_idletasks()
            preview_iids = self._billing_review_preview_iids(entry.iid)
            any_openable = any(
                self._resolve_file_path_for_row(iid) is not None for iid in preview_iids
            )
            open_file_btn.configure(state=tk.NORMAL if any_openable else tk.DISABLED)

            if preview_iids and all(
                self._resolve_file_path_for_row(iid) is None for iid in preview_iids
            ):
                _show_preview_message("ファイルが見つかりません")
                return

            preview_canvas.unbind("<Configure>")
            _clear_preview_canvas()
            y = _PREVIEW_PAD_X
            max_w = _PREVIEW_TEXT_WIDTH + _PREVIEW_PAD_X

            for iid in preview_iids:
                ui_row = self._current_row_dict_from_iid(iid)
                core = self._row_dict_to_core(ui_row)
                grid_no = _row_grid_no(core)
                file_name = self._file_name_from_row(ui_row) or "（なし）"
                caption = f"No.{grid_no} — {file_name}"
                preview_canvas.create_text(
                    _PREVIEW_PAD_X,
                    y,
                    text=caption,
                    anchor="nw",
                    font=("", 9, "bold"),
                )
                y += _PREVIEW_CAPTION_H

                path = self._resolve_file_path_for_row(iid)
                if path is None:
                    preview_canvas.create_text(
                        _PREVIEW_PAD_X,
                        y,
                        text="ファイルが見つかりません",
                        anchor="nw",
                        width=_PREVIEW_TEXT_WIDTH,
                        justify="left",
                    )
                    y += _PREVIEW_MSG_H
                    y += _PREVIEW_SECTION_GAP
                    continue

                suffix = path.suffix.lower()
                if suffix in EXCEL_SUFFIXES:
                    preview_canvas.create_text(
                        _PREVIEW_PAD_X,
                        y,
                        text=self._BILLING_REVIEW_EXCEL_CANVAS_MSG,
                        anchor="nw",
                        width=_PREVIEW_TEXT_WIDTH,
                        justify="left",
                    )
                    y += _PREVIEW_MSG_H
                    y += _PREVIEW_SECTION_GAP
                    continue
                if suffix != PDF_SUFFIX and suffix not in IMAGE_SUFFIXES:
                    preview_canvas.create_text(
                        _PREVIEW_PAD_X,
                        y,
                        text="この形式はプレビューできません",
                        anchor="nw",
                        width=_PREVIEW_TEXT_WIDTH,
                        justify="left",
                    )
                    y += _PREVIEW_MSG_H
                    y += _PREVIEW_SECTION_GAP
                    continue

                page_images = load_billing_preview_pages(path, max_edge_px=900)
                if not page_images:
                    preview_canvas.create_text(
                        _PREVIEW_PAD_X,
                        y,
                        text="プレビューを表示できません",
                        anchor="nw",
                        width=_PREVIEW_TEXT_WIDTH,
                        justify="left",
                    )
                    y += _PREVIEW_MSG_H
                    y += _PREVIEW_SECTION_GAP
                    continue

                page_total = len(page_images)
                for page_index, pil in enumerate(page_images, start=1):
                    if page_total > 1:
                        preview_canvas.create_text(
                            _PREVIEW_PAD_X,
                            y,
                            text=f"ページ {page_index} / {page_total}",
                            anchor="nw",
                            font=("", 9),
                        )
                        y += _PREVIEW_PAGE_CAPTION_H

                    photo = ImageTk.PhotoImage(pil)
                    self._billing_review_photos.append(photo)
                    iw, ih = pil.size
                    preview_canvas.create_image(
                        _PREVIEW_PAD_X, y, anchor="nw", image=photo
                    )
                    y += ih
                    max_w = max(max_w, _PREVIEW_PAD_X + iw)
                    if page_index < page_total:
                        y += _PREVIEW_PAGE_GAP

                y += _PREVIEW_SECTION_GAP

            preview_canvas.configure(scrollregion=(0, 0, max_w, max(y, 1)))
            preview_canvas.xview_moveto(0)
            preview_canvas.yview_moveto(0)

        def show_at_index(idx: int) -> None:
            if not entries:
                top.destroy()
                return
            state["index"] = max(0, min(idx, len(entries) - 1))
            entry = current_entry()
            nav_var.set(f"{state['index'] + 1} / {len(entries)}")
            no_disp = (
                str(entry.grid_no)
                if entry.grid_no < 10**9
                else "—"
            )
            no_var.set(no_disp)
            file_var.set(self._billing_review_file_names_display(entry.iid))
            emp_var.set(entry.employee_no or "（なし）")
            hours_var.set((entry.hours_write or "").strip())
            transport_var.set((entry.transport_write or "").strip())
            prev_btn.configure(state=(tk.NORMAL if state["index"] > 0 else tk.DISABLED))
            next_btn.configure(
                state=(tk.NORMAL if state["index"] < len(entries) - 1 else tk.DISABLED)
            )
            preview_iids = self._billing_review_preview_iids(entry.iid)
            self._apply_billing_review_grid_highlight(preview_iids)
            refresh_preview_image(entry)
            self._open_excel_files_for_billing_review(preview_iids)

        def _parse_review_update_fields(
            *,
            show_errors: bool,
        ) -> tuple[str, str] | None:
            h_raw = hours_var.get().strip()
            t_raw = transport_var.get().strip()
            hours = _normalize_billing_update_hours_copy(h_raw)
            transport = _normalize_billing_update_transport_copy(t_raw)
            if h_raw and not hours and not _billing_value_is_no_data(h_raw):
                if show_errors:
                    messagebox.showerror(
                        "グリッド更新",
                        f"{self.BILLING_UPDATE_HOURS_COL}の形式が正しくありません。\n"
                        "整数、または小数点以下2桁までの数値を入力してください。",
                        parent=top,
                    )
                return None
            if t_raw and not transport:
                if show_errors:
                    messagebox.showerror(
                        "グリッド更新",
                        f"{self.BILLING_UPDATE_TRANSPORT_COL}の形式が正しくありません。",
                        parent=top,
                    )
                return None
            return hours, transport

        def on_apply_grid() -> None:
            parsed = _parse_review_update_fields(show_errors=True)
            if parsed is None:
                return
            hours, transport = parsed
            entry = current_entry()
            core = self._row_dict_to_core(
                self._current_row_dict_from_iid(entry.iid)
            )
            _set_billing_update_columns(core, hours=hours, transport=transport)
            self._replace_row_with_result(entry.iid, core)
            self._loaded_rows = self._current_grid_rows()
            entries[state["index"]] = BillingReviewEntry(
                iid=entry.iid,
                file_name=entry.file_name,
                grid_no=entry.grid_no,
                employee_no=entry.employee_no,
                hours_write=hours,
                transport_write=transport,
            )
            hours_var.set(hours)
            transport_var.set(transport)
            self._refresh_billing_buttons_state()

        def on_prev() -> None:
            if state["index"] <= 0:
                return
            if not _confirm_discard_if_dirty():
                return
            show_at_index(state["index"] - 1)

        def on_next() -> None:
            if state["index"] >= len(entries) - 1:
                return
            if not _confirm_discard_if_dirty():
                return
            show_at_index(state["index"] + 1)

        def on_open_file() -> None:
            entry = current_entry()
            for iid in self._billing_review_preview_iids(entry.iid):
                if self._resolve_file_path_for_row(iid) is not None:
                    self._open_row_file(iid, force=True)

        def on_exclude() -> None:
            entry = current_entry()
            core = self._row_dict_to_core(self._current_row_dict_from_iid(entry.iid))
            core["user_judgment_company"] = "✖"
            core[self.FINAL_JUDGMENT_COL] = "✖"
            self._replace_row_with_result(entry.iid, core)
            self._loaded_rows = self._current_grid_rows()
            idx = state["index"]
            entries.pop(idx)
            if not entries:
                messagebox.showinfo(
                    "画像とデータ照合",
                    "照合対象がなくなりました。",
                    parent=top,
                )
                top.destroy()
                return
            show_at_index(min(idx, len(entries) - 1))

        prev_btn.configure(command=on_prev)
        next_btn.configure(command=on_next)
        apply_grid_btn.configure(command=on_apply_grid)
        open_file_btn.configure(command=on_open_file)
        exclude_btn.configure(command=on_exclude)
        top.bind("<Left>", lambda _e: on_prev())
        top.bind("<Right>", lambda _e: on_next())

        try:
            show_at_index(0)
            top.wait_window()
        finally:
            _unbind_preview_wheel()
            self._close_row_file()
            self._clear_billing_review_grid_highlight()

    def _prompt_billing_duplicate_resolution(
        self,
        rows: list[dict[str, str]],
        groups: list[BillingDuplicateGroupInfo],
    ) -> dict[tuple[str, str], BillingDuplicateResolution] | None:
        if not groups:
            return {}

        top = tk.Toplevel(self._root)
        top.title("請求データ作成 — 重複の解決")
        top.transient(self._root)
        top.geometry("920x520")
        top.minsize(640, 320)

        outer = ttk.Frame(top, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            outer,
            text=(
                "ファイル名の会社名と社員番号が同じ行があります。\n"
                "勤務時間と交通費は別々に選べます。「合算」はその項目だけ"
                "全行を足した値を代表行に載せます。\n"
                "結果は No が小さい行の更新用列に設定し、その行を最終判断「〇」、"
                "他行を「✖」にします。"
            ),
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        table_wrap = ttk.Frame(outer)
        table_wrap.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(table_wrap, highlightthickness=0)
        y_scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas)
        def _on_inner_configure(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        inner.bind("<Configure>", _on_inner_configure)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_canvas_configure(event: tk.Event) -> None:
            if event.width > 1:
                canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=y_scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_dialog_mousewheel(event: tk.Event) -> None:
            delta = getattr(event, "delta", 0)
            if delta:
                canvas.yview_scroll(int(-1 * (delta / 120)), "units")

        def _on_dialog_mousewheel_linux_up(_event: tk.Event) -> None:
            canvas.yview_scroll(-1, "units")

        def _on_dialog_mousewheel_linux_down(_event: tk.Event) -> None:
            canvas.yview_scroll(1, "units")

        def _unbind_dialog_mousewheel() -> None:
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                try:
                    top.unbind_all(seq)
                except tk.TclError:
                    pass

        top.bind_all("<MouseWheel>", _on_dialog_mousewheel)
        top.bind_all("<Button-4>", _on_dialog_mousewheel_linux_up)
        top.bind_all("<Button-5>", _on_dialog_mousewheel_linux_down)

        hours_vars: list[tk.StringVar] = []
        transport_vars: list[tk.StringVar] = []
        grid_row = 0

        for gi, group in enumerate(groups):
            members = [rows[i] for i in group.member_indices]
            ttk.Label(
                inner,
                text=(
                    f"【{gi + 1}】 {group.company_display} / 社員番号 {group.employee_no}"
                    f"（{len(members)} 件）"
                ),
                font=("", 9, "bold"),
            ).grid(row=grid_row, column=0, columnspan=4, sticky="w", padx=4, pady=(8, 2))
            grid_row += 1
            ttk.Label(
                inner,
                text=(
                    "下で「代表行（No 最小）に載せる値」を選びます。"
                    "勤務と交通費は別々に選べます。"
                ),
            ).grid(row=grid_row, column=0, columnspan=4, sticky="w", padx=4, pady=(0, 4))
            grid_row += 1

            headers = ("No", "ファイル名", "勤務（10進）", "交通費（読取）")
            for col, title in enumerate(headers):
                ttk.Label(inner, text=title, font=("", 9, "bold")).grid(
                    row=grid_row, column=col, sticky="w", padx=4, pady=2
                )
            grid_row += 1

            hours_pick_choices: list[tuple[str, str]] = []
            transport_pick_choices: list[tuple[str, str]] = []
            for mi, row_idx in enumerate(group.member_indices):
                row = rows[row_idx]
                no = _row_grid_no(row)
                if no >= 10**9:
                    no_disp = "—"
                else:
                    no_disp = str(no)
                fn = _row_file_name(row)
                if len(fn) > 36:
                    fn = fn[:33] + "..."
                h_cand, t_cand = billing_row_update_candidates(row)
                for col, text in enumerate(
                    (
                        no_disp,
                        fn,
                        self._billing_cell_preview(_row_total_hours_decimal(row)),
                        self._billing_cell_preview(_row_transport_expense_raw(row)),
                    )
                ):
                    ttk.Label(inner, text=text).grid(
                        row=grid_row, column=col, sticky="w", padx=4, pady=2
                    )
                hours_pick_choices.append(
                    (
                        f"pick:{mi}",
                        f"No.{no_disp} の勤務: {self._billing_cell_preview(h_cand)}",
                    )
                )
                transport_pick_choices.append(
                    (
                        f"pick:{mi}",
                        f"No.{no_disp} の交通費: {self._billing_cell_preview(t_cand)}",
                    )
                )
                grid_row += 1

            rep_row = rows[group.member_indices[0]]
            rep_no = _row_grid_no(rep_row)
            rep_no_disp = str(rep_no) if rep_no < 10**9 else "—"
            hours_sum = resolve_billing_hours_from_group(members, "sum", 0)
            transport_sum = resolve_billing_transport_from_group(members, "sum", 0)
            hours_pick_choices.append(
                (
                    "sum:",
                    f"No.{rep_no_disp} に勤務時間を合算: "
                    f"{self._billing_cell_preview(hours_sum)}",
                )
            )
            transport_pick_choices.append(
                (
                    "sum:",
                    f"No.{rep_no_disp} に交通費を合算: "
                    f"{self._billing_cell_preview(transport_sum)}",
                )
            )

            hv = tk.StringVar(value="pick:0")
            tv = tk.StringVar(value="pick:0")
            hours_vars.append(hv)
            transport_vars.append(tv)

            hf = ttk.LabelFrame(inner, text="勤務時間（代表行へ反映）", padding=(6, 4))
            hf.grid(row=grid_row, column=0, columnspan=4, sticky="ew", padx=4, pady=4)
            for pi, (val, label) in enumerate(hours_pick_choices):
                ttk.Radiobutton(hf, text=label, variable=hv, value=val).grid(
                    row=pi, column=0, sticky="w", padx=(0, 12), pady=1
                )
            grid_row += 1

            tf = ttk.LabelFrame(inner, text="交通費（代表行へ反映）", padding=(6, 4))
            tf.grid(row=grid_row, column=0, columnspan=4, sticky="ew", padx=4, pady=(0, 8))
            for pi, (val, label) in enumerate(transport_pick_choices):
                ttk.Radiobutton(tf, text=label, variable=tv, value=val).grid(
                    row=pi, column=0, sticky="w", padx=(0, 12), pady=1
                )
            grid_row += 1

        result: dict[str, object] = {"ok": False}

        def _parse_choice(raw: str) -> tuple[str, int]:
            if raw.startswith("sum"):
                return "sum", 0
            if raw.startswith("pick:"):
                try:
                    return "pick", int(raw.split(":", 1)[1])
                except ValueError:
                    return "pick", 0
            return "pick", 0

        def on_ok() -> None:
            out: dict[tuple[str, str], BillingDuplicateResolution] = {}
            for group, hv, tv in zip(groups, hours_vars, transport_vars, strict=True):
                hm, hi = _parse_choice(hv.get())
                tm, ti = _parse_choice(tv.get())
                out[group.key] = BillingDuplicateResolution(
                    hours_mode=hm,
                    hours_pick_index=hi,
                    transport_mode=tm,
                    transport_pick_index=ti,
                )
            result["ok"] = True
            result["data"] = out
            _unbind_dialog_mousewheel()
            top.destroy()

        def on_cancel() -> None:
            _unbind_dialog_mousewheel()
            top.destroy()

        btns = ttk.Frame(outer)
        btns.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(
            side=tk.RIGHT, padx=(0, 8)
        )
        top.bind("<Escape>", lambda _e: on_cancel())
        try:
            top.wait_window()
        finally:
            _unbind_dialog_mousewheel()
        if not result.get("ok"):
            return None
        return result.get("data") or {}

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
        scope = self._billing_target_scope_label()
        if not messagebox.askokcancel(
            "請求データ作成",
            f"対象: {scope}\n\n"
            "各行について、更新用合計勤務時間（10進）・更新用交通費合計を\n"
            "読取値から作成します。\n"
            "ファイル名の会社名と社員番号が同じ行がある場合は、\n"
            "採用・合算を選ぶダイアログを表示します（代表は No が小さい行）。",
            parent=self._root,
        ):
            return

        ordered: list[tuple[str, dict[str, str]]] = []
        for iid in self._billing_target_iids():
            ordered.append((iid, self._row_dict_to_core(self._current_row_dict_from_iid(iid))))

        cores = [core for _, core in ordered]
        dup_groups = build_billing_duplicate_groups(cores)
        dialog_groups = [g for g in dup_groups if g.needs_dialog]
        resolutions = self._prompt_billing_duplicate_resolution(cores, dialog_groups)
        if resolutions is None:
            return

        updated_count = apply_billing_create_for_rows(cores, resolutions)
        for (iid, _), core in zip(ordered, cores, strict=True):
            self._replace_row_with_result(iid, core)

        self._loaded_rows = self._current_grid_rows()
        self._status_var.set(
            f"請求データ作成完了（{scope}）: {updated_count} 行に更新用列を設定しました"
        )
        messagebox.showinfo(
            "請求データ作成",
            f"対象: {scope}\n更新用列を設定しました（{updated_count} 行）。",
            parent=self._root,
        )

    def _delete_billing_data(self) -> None:
        if self._busy:
            return
        if not self._tree.get_children():
            messagebox.showinfo(
                "請求データ削除",
                "グリッドに行がありません。",
                parent=self._root,
            )
            return
        scope = self._billing_target_scope_label()
        if not messagebox.askokcancel(
            "請求データ削除",
            f"対象: {scope}\n\n"
            "更新用合計勤務時間（10進）と更新用交通費合計をクリアします。",
            parent=self._root,
        ):
            return

        cleared = 0
        for iid in self._billing_target_iids():
            core = self._row_dict_to_core(self._current_row_dict_from_iid(iid))
            if _row_billing_update_hours_decimal(core) or _row_billing_update_transport(
                core
            ):
                cleared += 1
            clear_billing_update_hours_column(core)
            self._replace_row_with_result(iid, core)

        self._loaded_rows = self._current_grid_rows()
        self._status_var.set(
            f"請求データ削除完了（{scope}）: {cleared} 行の更新用列をクリアしました"
        )
        messagebox.showinfo(
            "請求データ削除",
            f"対象: {scope}\n"
            f"更新用合計勤務時間（10進）・更新用交通費合計をクリアしました（{cleared} 行）。",
            parent=self._root,
        )

    def _billing_ts_cell_label(self, value: str) -> str:
        t = (value or "").strip()
        return t if t else "（空）"

    def _show_billing_file_update_result_dialog(
        self,
        *,
        scope: str,
        billing_file_name: str,
        ok_count: int,
        fail_count: int,
        skip_count: int,
        detail_rows: list[tuple[str, str, str]],
    ) -> None:
        top = tk.Toplevel(self._root)
        top.title("請求ファイル更新 — 結果")
        top.transient(self._root)
        top.geometry("760x520")
        top.minsize(560, 360)

        outer = ttk.Frame(top, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        summary = (
            f"対象: {scope}\n"
            f"請求用ファイル: {billing_file_name}\n\n"
            f"成功: {ok_count} 件\n"
            f"失敗: {fail_count} 件\n"
            f"スキップ: {skip_count} 件"
        )
        ttk.Label(outer, text=summary, justify="left").pack(anchor="w", pady=(0, 8))

        list_frame = ttk.LabelFrame(
            outer, text="成功以外の行（グリッドの請求用ファイル更新列も参照）", padding=4
        )
        list_frame.pack(fill=tk.BOTH, expand=True)

        def _unbind_result_mousewheel() -> None:
            pass

        if not detail_rows:
            ttk.Label(
                list_frame,
                text="失敗・スキップはありません。",
                justify="left",
            ).pack(anchor="w", padx=4, pady=4)
        else:
            table_wrap = ttk.Frame(list_frame)
            table_wrap.pack(fill=tk.BOTH, expand=True)
            cols = ("no", "file_name", "result")
            tree = ttk.Treeview(
                table_wrap,
                columns=cols,
                show="headings",
                selectmode="browse",
                height=12,
            )
            tree.heading("no", text="No")
            tree.heading("file_name", text="ファイル名")
            tree.heading("result", text="結果")
            tree.column("no", width=56, stretch=False, minwidth=40)
            tree.column("file_name", width=360, stretch=True, minwidth=120)
            tree.column("result", width=200, stretch=True, minwidth=80)
            y_scroll = ttk.Scrollbar(
                table_wrap, orient=tk.VERTICAL, command=tree.yview
            )
            x_scroll = ttk.Scrollbar(
                table_wrap, orient=tk.HORIZONTAL, command=tree.xview
            )
            tree.configure(
                yscrollcommand=y_scroll.set,
                xscrollcommand=x_scroll.set,
            )
            tree.grid(row=0, column=0, sticky="nsew")
            y_scroll.grid(row=0, column=1, sticky="ns")
            x_scroll.grid(row=1, column=0, sticky="ew")
            table_wrap.rowconfigure(0, weight=1)
            table_wrap.columnconfigure(0, weight=1)
            for no_disp, file_name, symbol in detail_rows:
                tree.insert("", tk.END, values=(no_disp, file_name, symbol))

            def _on_result_mousewheel(event: tk.Event) -> None:
                delta = getattr(event, "delta", 0)
                if not delta:
                    return
                steps = int(-1 * (delta / 120))
                if event.state & 0x1:
                    tree.xview_scroll(steps, "units")
                else:
                    tree.yview_scroll(steps, "units")

            def _on_result_mousewheel_linux_up(_event: tk.Event) -> None:
                tree.yview_scroll(-1, "units")

            def _on_result_mousewheel_linux_down(_event: tk.Event) -> None:
                tree.yview_scroll(1, "units")

            def _do_unbind_result_mousewheel() -> None:
                for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                    try:
                        top.unbind_all(seq)
                    except tk.TclError:
                        pass

            _unbind_result_mousewheel = _do_unbind_result_mousewheel

            top.bind_all("<MouseWheel>", _on_result_mousewheel)
            top.bind_all("<Button-4>", _on_result_mousewheel_linux_up)
            top.bind_all("<Button-5>", _on_result_mousewheel_linux_down)

        def _close() -> None:
            _unbind_result_mousewheel()
            top.destroy()

        btn_row = ttk.Frame(outer)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="閉じる", command=_close).pack(side=tk.RIGHT)
        top.protocol("WM_DELETE_WINDOW", _close)
        top.grab_set()
        try:
            self._root.wait_window(top)
        finally:
            _unbind_result_mousewheel()

    def _prompt_billing_ts_overwrite_confirm(
        self, items: list[BillingTsWriteItem]
    ) -> bool:
        top = tk.Toplevel(self._root)
        top.title("請求ファイル更新 — 上書き確認")
        top.transient(self._root)
        top.geometry("980x440")
        top.minsize(720, 280)

        outer = ttk.Frame(top, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            outer,
            text="Excel に既存値がある行です。上書きする列を選択してください。",
        ).pack(anchor="w", pady=(0, 8))

        table_wrap = ttk.Frame(outer)
        table_wrap.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(table_wrap, highlightthickness=0)
        y_scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=canvas.yview)
        x_scroll = ttk.Scrollbar(
            table_wrap, orient=tk.HORIZONTAL, command=canvas.xview
        )
        inner = ttk.Frame(canvas)

        def _on_inner_configure(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        inner.bind("<Configure>", _on_inner_configure)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(
            xscrollcommand=x_scroll.set,
            yscrollcommand=y_scroll.set,
        )
        canvas.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        table_wrap.rowconfigure(0, weight=1)
        table_wrap.columnconfigure(0, weight=1)

        def _on_overwrite_mousewheel(event: tk.Event) -> None:
            delta = getattr(event, "delta", 0)
            if not delta:
                return
            steps = int(-1 * (delta / 120))
            if event.state & 0x1:
                canvas.xview_scroll(steps, "units")
            else:
                canvas.yview_scroll(steps, "units")

        def _on_overwrite_mousewheel_linux_up(_event: tk.Event) -> None:
            canvas.yview_scroll(-1, "units")

        def _on_overwrite_mousewheel_linux_down(_event: tk.Event) -> None:
            canvas.yview_scroll(1, "units")

        def _unbind_overwrite_mousewheel() -> None:
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                try:
                    top.unbind_all(seq)
                except tk.TclError:
                    pass

        top.bind_all("<MouseWheel>", _on_overwrite_mousewheel)
        top.bind_all("<Button-4>", _on_overwrite_mousewheel_linux_up)
        top.bind_all("<Button-5>", _on_overwrite_mousewheel_linux_down)

        headers = (
            "対象ファイル名",
            "Excel行",
            "通常請求時間（現在）",
            "通常請求時間（更新後）",
            "旅費交通費（現在）",
            "旅費交通費（更新後）",
            "勤務時間を上書き",
            "交通費を上書き",
        )
        for col, title in enumerate(headers):
            ttk.Label(inner, text=title, font=("", 9, "bold")).grid(
                row=0, column=col, sticky="w", padx=4, pady=2
            )

        hour_vars: list[tk.BooleanVar] = []
        transport_vars: list[tk.BooleanVar] = []
        for row_i, item in enumerate(items, start=1):
            vh = tk.BooleanVar(value=True)
            vt = tk.BooleanVar(value=True)
            hour_vars.append(vh)
            transport_vars.append(vt)
            fn = item.file_name
            if len(fn) > 42:
                fn = fn[:39] + "..."
            row_nums = ",".join(str(n) for n in item.sheet_row_indices)
            values = (
                fn,
                row_nums,
                self._billing_ts_cell_label(item.old_hours_display),
                self._billing_ts_cell_label(item.new_hours_display),
                self._billing_ts_cell_label(item.old_transport_display),
                self._billing_ts_cell_label(item.new_transport_display),
            )
            for col, text in enumerate(values):
                ttk.Label(inner, text=text).grid(
                    row=row_i, column=col, sticky="w", padx=4, pady=2
                )
            ttk.Checkbutton(inner, variable=vh).grid(
                row=row_i, column=6, padx=4, pady=2
            )
            ttk.Checkbutton(inner, variable=vt).grid(
                row=row_i, column=7, padx=4, pady=2
            )

        result = {"ok": False}

        def on_ok() -> None:
            for idx, item in enumerate(items):
                item.write_hours = hour_vars[idx].get()
                item.write_transport = transport_vars[idx].get()
            result["ok"] = True
            _unbind_overwrite_mousewheel()
            top.destroy()

        def on_cancel() -> None:
            _unbind_overwrite_mousewheel()
            top.destroy()

        btns = ttk.Frame(outer)
        btns.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="選択した内容で実行", command=on_ok).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

        top.protocol("WM_DELETE_WINDOW", on_cancel)
        top.grab_set()
        try:
            self._root.wait_window(top)
        finally:
            _unbind_overwrite_mousewheel()
        return result["ok"]

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
        scope = self._billing_target_scope_label()
        targets: list[tuple[str, dict[str, str]]] = []
        for iid in self._billing_target_iids():
            ui_row = self._current_row_dict_from_iid(iid)
            if self._final_judgment_symbol_from_row(ui_row) != "〇":
                continue
            core_row = self._row_dict_to_core(ui_row)
            hours_val = _row_billing_update_hours_decimal(core_row)
            transport_val = _row_billing_update_transport(core_row)
            if _is_billing_aggregated_marker(hours_val):
                continue
            if not (hours_val or "").strip() and not (transport_val or "").strip():
                continue
            targets.append((iid, core_row))
        if not targets:
            messagebox.showinfo(
                "請求ファイル更新",
                f"対象: {scope}\n\n"
                "この範囲のうち、最終判断が「〇」かつ更新用列（勤務時間または交通費）が"
                "設定された行がありません。\n"
                "先に「請求データ作成」を実行してください。",
                parent=self._root,
            )
            return
        if not messagebox.askokcancel(
            "請求ファイル更新",
            f"対象: {scope}\n\n"
            "最終判断が「〇」のレコードについて、更新用合計勤務時間（10進）・"
            "更新用交通費合計を請求用ファイルへ反映します。",
            parent=self._root,
        ):
            return
        core_rows = [core_row for _, core_row in targets]
        try:
            planned = plan_billing_ts_writes(self._billing_file_path, core_rows)
        except OSError as e:
            messagebox.showerror(
                "請求ファイル更新",
                f"請求用ファイルを開けませんでした。\n\n{e}",
                parent=self._root,
            )
            return
        except Exception as e:
            messagebox.showerror(
                "請求ファイル更新",
                f"請求用ファイルの読み込みに失敗しました。\n\n{e}",
                parent=self._root,
            )
            return

        confirm_items = [
            x
            for x in planned
            if isinstance(x, BillingTsWriteItem) and x.needs_confirm
        ]
        if confirm_items and not self._prompt_billing_ts_overwrite_confirm(
            confirm_items
        ):
            return

        write_items = [x for x in planned if isinstance(x, BillingTsWriteItem)]
        try:
            apply_results = (
                apply_billing_ts_writes(self._billing_file_path, write_items)
                if write_items
                else []
            )
            results = merge_billing_ts_plan_results(planned, apply_results)
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
        skip_count = 0
        detail_rows: list[tuple[str, str, str]] = []
        for (_iid, core_row), symbol in zip(targets, results, strict=True):
            self._set_billing_update_result_cell(_iid, symbol)
            if symbol == "〇":
                ok_count += 1
            elif symbol == BILLING_UPDATE_SKIP:
                skip_count += 1
            if symbol != "〇":
                grid_no = _row_grid_no(core_row)
                no_disp = str(grid_no) if grid_no < 10**9 else "—"
                file_name = (
                    (core_row.get("file_name") or "").strip() or "（ファイル名なし）"
                )
                detail_rows.append((no_disp, file_name, symbol))
        fail_count = len(targets) - ok_count - skip_count
        self._loaded_rows = self._current_grid_rows()
        self._status_var.set(
            f"請求ファイル更新完了（{scope}）: 成功 {ok_count} 件 / 失敗 {fail_count} 件 "
            f"/ スキップ {skip_count} 件（対象 {len(targets)} 件）"
        )
        self._show_billing_file_update_result_dialog(
            scope=scope,
            billing_file_name=self._billing_file_path.name,
            ok_count=ok_count,
            fail_count=fail_count,
            skip_count=skip_count,
            detail_rows=detail_rows,
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
        self._row_extra.clear()
        self._grid_row_no_seq = 0
        self._sort_column = None
        self._sort_reverse = False
        self._update_column_heading_labels()

    def _grid_sort_empty_marker(self, value: str) -> bool:
        t = (value or "").strip()
        return not t or t in (
            "（なし）",
            "不明",
            "（不明）",
            "（データなし）",
            "（対象レコードなし）",
        )

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

    def _column_heading_display_text(self, column_id: str) -> str:
        if column_id in (
            self.LABOR_AI_READ_STATE_COL,
            self.TRANSPORT_AI_READ_STATE_COL,
        ):
            return AI_READ_STATE_COLUMN_HEADING
        return column_id

    def _update_column_heading_labels(self) -> None:
        for h in self._tree_column_headings:
            text = self._column_heading_display_text(h)
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

        self._update_column_heading_labels()
        self._refresh_duplicate_employee_no_highlight()

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
            (self.MATCH_PERSON_COL, "match_person"),
            (self.MATCH_DOC_TYPE_COL, "match_doc_type"),
            (self.EMPLOYEE_NO_COL, "employee_no"),
            (self.BILLING_UPDATE_HOURS_COL, "billing_update_hours_decimal"),
            (LEGACY_BILLING_UPDATE_HOURS_COL, "billing_update_hours_decimal"),
            (self.TOTAL_HOURS_DECIMAL_COL, "total_hours_decimal"),
            (self.TOTAL_HOURS_RAW_COL, "total_hours_raw"),
            (self.LABOR_AI_READ_STATE_COL, "labor_ai_read_state"),
            (self.BILLING_UPDATE_TRANSPORT_COL, "billing_update_transport"),
            (self.TRANSPORT_EXPENSE_COL, "transport_expense_raw"),
            (self.TRANSPORT_AI_READ_STATE_COL, "transport_ai_read_state"),
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
                        or row.get(LEGACY_SUFFIX_READ_YEAR_COL)
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
                        or row.get(LEGACY_SUFFIX_READ_MONTH_COL)
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
            elif core_key == "match_person":
                val = str(
                    row.get(core_key)
                    or row.get(self.MATCH_PERSON_COL)
                    or ""
                ).strip()
                val = normalize_judgment_symbol(val) if val else ""
            elif core_key == "match_doc_type":
                val = str(
                    row.get(core_key)
                    or row.get(self.MATCH_DOC_TYPE_COL)
                    or ""
                ).strip()
                val = normalize_judgment_symbol(val) if val in ("〇", "△", "✖") else val
            elif core_key == "name_company_1":
                val = _document_company_for_display(
                    str(
                        row.get(self.COMPANY_COL)
                        or row.get(LEGACY_SUFFIX_READ_COMPANY_COL)
                        or row.get(self.LEGACY_COMPANY_READ_LONG_COL)
                        or row.get(self.LEGACY_COMPANY_NAME_COL)
                        or row.get(self.LEGACY_COMPANY_COL)
                        or row.get(core_key)
                        or ""
                    ).strip()
                )
            elif core_key == "name_person_from_doc":
                val = str(
                    row.get(self.PERSON_COL)
                    or row.get(LEGACY_SUFFIX_READ_PERSON_COL)
                    or row.get(self.LEGACY_PERSON_COL)
                    or row.get(core_key)
                    or ""
                ).strip()
            elif core_key == "total_hours_raw":
                val = str(
                    row.get("total_hours_raw")
                    or row.get(self.TOTAL_HOURS_RAW_COL)
                    or row.get(LEGACY_SUFFIX_READ_TOTAL_HOURS_RAW_COL)
                    or ""
                ).strip()
            elif core_key == "transport_expense_raw":
                val = str(
                    row.get("transport_expense_raw")
                    or row.get(self.TRANSPORT_EXPENSE_COL)
                    or row.get(LEGACY_SUFFIX_READ_TRANSPORT_EXPENSE_COL)
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
        grid_no = str(row.get("grid_row_no") or "").strip()
        if grid_no:
            out["grid_row_no"] = grid_no
        if "name_company_from_doc" in out:
            out["name_company_from_doc"] = _document_company_for_display(
                str(out.get("name_company_from_doc") or "")
            )
        return out

    def _grid_values_from_row(
        self, row: dict[str, str], *, sync_user_judgment_to_auto: bool = False
    ) -> tuple[str, ...]:
        """内部row形式・UI保存形式のどちらでも Treeview 表示値へ変換する。"""
        return row_display_values(
            self._row_dict_to_core(row),
            sync_user_judgment_to_auto=sync_user_judgment_to_auto,
        )

    def _row_no_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.ROW_NO_COL)

    def _tree_row_no(self, iid: str) -> str:
        vals = self._tree.item(iid, "values") or ()
        ci = self._row_no_column_index()
        return str(vals[ci]).strip() if ci < len(vals) else ""

    def _assign_row_no_to_iid(self, iid: str, no: int | None = None) -> int:
        """行に No を付与する（既存 No は上書きしない用途以外で no を指定）。"""
        if no is None:
            self._grid_row_no_seq += 1
            no = self._grid_row_no_seq
        else:
            self._grid_row_no_seq = max(self._grid_row_no_seq, no)
        vals = list(self._tree.item(iid, "values") or ())
        ci = self._row_no_column_index()
        while len(vals) <= ci:
            vals.append("")
        vals[ci] = str(no)
        self._tree.item(iid, values=tuple(vals))
        return no

    def _merge_preserved_row_no(
        self, rid: str, new_vals: tuple[str, ...]
    ) -> tuple[str, ...]:
        """行更新時に No 列だけ保持する（未設定なら新規採番）。"""
        ci = self._row_no_column_index()
        old_vals = self._tree.item(rid, "values") or ()
        old_no = str(old_vals[ci]).strip() if ci < len(old_vals) else ""
        vals = list(new_vals)
        while len(vals) <= ci:
            vals.append("")
        if old_no:
            vals[ci] = old_no
            if old_no.isdigit():
                self._grid_row_no_seq = max(self._grid_row_no_seq, int(old_no))
        else:
            self._grid_row_no_seq += 1
            vals[ci] = str(self._grid_row_no_seq)
        return tuple(vals)

    def _sync_grid_row_no_seq_from_tree(self) -> None:
        """グリッド上の No 最大値から採番カウンタを復元する。"""
        max_no = 0
        for iid in self._tree.get_children():
            no_s = self._tree_row_no(iid)
            if no_s.isdigit():
                max_no = max(max_no, int(no_s))
        self._grid_row_no_seq = max_no

    def _renumber_grid_rows(self) -> None:
        """表示順に No を 1…N に振り直す。"""
        children = list(self._tree.get_children())
        if not children:
            self._grid_row_no_seq = 0
            return

        ci_no = self._row_no_column_index()
        for n, iid in enumerate(children, start=1):
            vals = list(self._tree.item(iid, "values") or ())
            while len(vals) <= ci_no:
                vals.append("")
            vals[ci_no] = str(n)
            self._tree.item(iid, values=tuple(vals))

        self._grid_row_no_seq = len(children)
        self._refresh_duplicate_employee_no_highlight()

    def _refresh_duplicate_employee_no_highlight(self) -> None:
        """有効社員番号が2行以上あるとき、該当行全体の文字色を赤にする。"""
        tag_dup = self._TAG_DUP_EMPLOYEE_NO
        tag_reanalysis = self._TAG_REANALYSIS_ACTIVE
        tag_billing_review = self._TAG_BILLING_REVIEW_ACTIVE
        preserve_tags = (tag_reanalysis, tag_billing_review)
        counts: dict[str, int] = {}
        for iid in self._tree.get_children():
            row = self._row_dict_to_core(
                self._merge_row_extra(iid, self._current_row_dict_from_iid(iid))
            )
            key = _normalize_employee_no_cell_value(_row_employee_no(row))
            if not _is_valid_employee_no(key):
                continue
            counts[key] = counts.get(key, 0) + 1
        dup_keys = {k for k, v in counts.items() if v >= 2}

        for iid in self._tree.get_children():
            try:
                cur = tuple(self._tree.item(iid, "tags") or ())
            except tk.TclError:
                continue
            row = self._row_dict_to_core(
                self._merge_row_extra(iid, self._current_row_dict_from_iid(iid))
            )
            key = _normalize_employee_no_cell_value(_row_employee_no(row))
            is_dup = _is_valid_employee_no(key) and key in dup_keys
            without_dup = tuple(t for t in cur if t != tag_dup)
            other = tuple(t for t in without_dup if t not in preserve_tags)
            kept = tuple(t for t in preserve_tags if t in without_dup)
            if is_dup:
                new_tags = (tag_dup,) + other + kept
            else:
                new_tags = other + kept
            if new_tags != cur:
                self._tree.item(iid, tags=new_tags)

    def _user_judgment_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.FINAL_JUDGMENT_COL)

    def _employee_no_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.EMPLOYEE_NO_COL)

    def _total_hours_decimal_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.TOTAL_HOURS_DECIMAL_COL)

    def _total_hours_raw_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.TOTAL_HOURS_RAW_COL)

    def _transport_expense_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.TRANSPORT_EXPENSE_COL)

    def _match_company_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.MATCH_COMPANY_COL)

    def _match_person_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.MATCH_PERSON_COL)

    def _sync_row_extra_from_row(self, iid: str, row: dict[str, str]) -> None:
        extra = dict(self._row_extra.get(iid, {}))
        for key in self._ROW_EXTRA_KEYS:
            val = str(row.get(key) or "").strip()
            if key == "name_company_from_doc":
                val = _document_company_for_display(val)
            if val:
                extra[key] = val
            elif key in extra:
                del extra[key]
        if extra:
            self._row_extra[iid] = extra
        elif iid in self._row_extra:
            del self._row_extra[iid]

    def _merge_row_extra(self, iid: str, row: dict[str, str]) -> dict[str, str]:
        out = dict(row)
        for key, val in self._row_extra.get(iid, {}).items():
            if key not in self._ROW_EXTRA_KEYS:
                continue
            if key == "name_company_from_doc":
                val = _document_company_for_display(val)
            if key == "match_company_manual_ok":
                out[key] = val
                continue
            if key == "match_person_manual":
                out[key] = val
                continue
            if val and not str(out.get(key) or "").strip():
                out[key] = val
        return out

    def _is_match_company_manual_ok(self, row: dict[str, str]) -> bool:
        return (row.get("match_company_manual_ok") or "").strip() == "1"

    def _is_match_person_manual(self, row: dict[str, str]) -> bool:
        return is_match_person_manual(row)

    def _reload_company_aliases(self) -> None:
        self._company_aliases = load_company_aliases(self._company_alias_table_path)

    def _persist_company_aliases(self) -> None:
        save_company_aliases(self._company_alias_table_path, self._company_aliases)

    def _company_pair_from_row(self, row: dict[str, str]) -> tuple[str, str]:
        fn = self._file_name_from_row(row)
        file_co = (row.get("name_company_from_file") or "").strip()
        if not file_co and fn:
            file_co = _parse_filename_company_and_person(fn)[0].strip()
        doc_co = ""
        for src in (
            row.get("name_company_from_doc"),
            row.get("name_company_1"),
            row.get(self.COMPANY_COL),
            row.get(self.LEGACY_COMPANY_READ_LONG_COL),
            row.get(LEGACY_COMPANY_NAME_COL),
            row.get(LEGACY_COMPANY_COL),
        ):
            doc_co = _document_company_for_display(str(src or "").strip())
            if doc_co:
                break
        return file_co, doc_co

    def _format_doc_company_for_message(self, doc_co: str) -> str:
        return doc_co if doc_co else "（空欄）"

    def _recalculate_company_match_for_row(
        self, rid: str, row: dict[str, str] | None = None
    ) -> None:
        core = self._row_dict_to_core(
            self._merge_row_extra(
                rid, row if row is not None else self._current_row_dict_from_iid(rid)
            )
        )
        core["match_company"] = recalculate_match_company_for_row(
            core, company_aliases=self._company_aliases
        )
        if self._is_match_company_manual_ok(core):
            core["match_company"] = "〇"
        core[self.MATCH_COMPANY_COL] = core["match_company"]
        if not self._is_match_person_manual(core):
            core["match_person"] = recalculate_match_person_for_row(core)
        core[self.MATCH_PERSON_COL] = core.get("match_person", "")
        core[self.MATCH_DOC_TYPE_COL] = (
            core.get("match_doc_type")
            or core.get(self.MATCH_DOC_TYPE_COL)
            or ""
        )
        self._recalculate_auto_judgment_for_row(rid, core)

    def _set_match_person_symbol(self, rid: str, value: str) -> None:
        """氏名比較を右クリックで手動設定し、自動判断・最終判断を連動更新する。"""
        value = normalize_judgment_symbol(value)
        if value not in ("〇", "△", "✖"):
            return
        row = self._merge_row_extra(rid, self._current_row_dict_from_iid(rid))
        core = self._row_dict_to_core(row)
        core["match_person_manual"] = "1"
        core["match_person"] = value
        core[self.MATCH_PERSON_COL] = value
        self._recalculate_auto_judgment_for_row(rid, core)
        self._sync_row_extra_from_row(rid, core)
        self._loaded_rows = self._current_grid_rows()
        self._status_var.set(f"氏名比較を {value} に変更しました")

    def _set_match_company_temp_ok(self, rid: str) -> None:
        """会社名比較を対応表登録なしで一旦〇にする。"""
        row = self._merge_row_extra(rid, self._current_row_dict_from_iid(rid))
        core = self._row_dict_to_core(row)
        core["match_company_manual_ok"] = "1"
        core["match_company"] = "〇"
        core[self.MATCH_COMPANY_COL] = "〇"
        if not self._is_match_person_manual(core):
            core["match_person"] = recalculate_match_person_for_row(core)
        core[self.MATCH_PERSON_COL] = core.get("match_person", "")
        core[self.MATCH_DOC_TYPE_COL] = (
            core.get("match_doc_type")
            or core.get(self.MATCH_DOC_TYPE_COL)
            or ""
        )
        self._recalculate_auto_judgment_for_row(rid, core)
        self._sync_row_extra_from_row(rid, core)
        self._status_var.set("会社名比較を一旦〇にしました（対応表には登録しません）")

    def _file_company_lookup_key(self, file_co: str) -> str:
        return company_alias_lookup_key(file_co, "")[0]

    def _find_empty_doc_company_mismatch_rows(
        self,
        file_co: str,
        *,
        exclude_rid: str | None = None,
    ) -> list[str]:
        """同じファイル名会社で文書会社が未取得かつ会社名比較が✖の行。"""
        target_key = self._file_company_lookup_key(file_co)
        if not target_key:
            return []
        hits: list[str] = []
        for iid in self._tree.get_children():
            if exclude_rid and iid == exclude_rid:
                continue
            row = self._merge_row_extra(iid, self._current_row_dict_from_iid(iid))
            core = self._row_dict_to_core(row)
            if (core.get("match_company") or "").strip() != "✖":
                continue
            row_file_co, row_doc_co = self._company_pair_from_row(row)
            if self._file_company_lookup_key(row_file_co) != target_key:
                continue
            if row_doc_co:
                continue
            hits.append(iid)
        return hits

    def _apply_doc_company_to_row(self, rid: str, doc_co: str) -> None:
        row = self._merge_row_extra(rid, self._current_row_dict_from_iid(rid))
        row[self.COMPANY_COL] = doc_co
        row["name_company_1"] = doc_co
        row["name_company_from_doc"] = doc_co
        core = self._row_dict_to_core(row)
        core["name_company_1"] = doc_co
        core["name_company_from_doc"] = doc_co
        self._recalculate_company_match_for_row(rid, core)

    def _offer_bulk_doc_company_fill(
        self,
        *,
        source_rid: str,
        file_co: str,
        doc_co: str,
    ) -> int:
        """文書会社未取得の類似行へ、登録した文書会社を設定するか確認する。"""
        targets = self._find_empty_doc_company_mismatch_rows(
            file_co, exclude_rid=source_rid
        )
        if not targets:
            return 0
        sample_names = [
            self._file_name_from_row(
                self._merge_row_extra(iid, self._current_row_dict_from_iid(iid))
            )
            for iid in targets[:5]
        ]
        sample_names = [n for n in sample_names if n]
        lines = [
            f"同じファイル名会社で文書会社が未取得の行が {len(targets)} 件あります。",
            f"登録した文書会社「{doc_co}」を設定しますか？",
            "",
            f"ファイル名会社: {file_co}",
        ]
        if sample_names:
            lines.append("")
            lines.append("対象例:")
            lines.extend(f"・{name}" for name in sample_names)
            if len(targets) > len(sample_names):
                lines.append(f"…他 {len(targets) - len(sample_names)} 件")
        if not messagebox.askyesno(
            "会社名対応の一括反映",
            "\n".join(lines),
            parent=self._root,
        ):
            return 0
        for iid in targets:
            self._apply_doc_company_to_row(iid, doc_co)
        return len(targets)

    def _recalculate_all_company_matches(self) -> None:
        for iid in self._tree.get_children():
            self._recalculate_company_match_for_row(iid)

    def _register_company_alias_from_row(self, rid: str) -> None:
        row = self._merge_row_extra(rid, self._current_row_dict_from_iid(rid))
        core = self._row_dict_to_core(row)
        file_co, doc_co = self._company_pair_from_row(core)
        if not file_co:
            messagebox.showwarning(
                "会社名対応登録",
                "ファイル名から会社名を取得できません。",
                parent=self._root,
            )
            return
        if _company_text_contains_seraku(file_co):
            messagebox.showwarning(
                "会社名対応登録",
                "ファイル名の会社名がセラクのため、登録できません。",
                parent=self._root,
            )
            return
        if not messagebox.askokcancel(
            "会社名対応登録",
            f"次の対応を「会社名対応表.json」に登録します。\n\n"
            f"ファイル名会社: {file_co}\n"
            f"文書会社: {self._format_doc_company_for_message(doc_co)}",
            parent=self._root,
        ):
            return
        added = add_company_alias(
            self._company_aliases,
            file_company=file_co,
            doc_company=doc_co,
        )
        try:
            self._persist_company_aliases()
        except OSError as e:
            messagebox.showerror(
                "会社名対応登録",
                f"対応表の保存に失敗しました。\n\n{e}",
                parent=self._root,
            )
            return
        self._recalculate_all_company_matches()
        applied = 0
        if doc_co:
            applied = self._offer_bulk_doc_company_fill(
                source_rid=rid,
                file_co=file_co,
                doc_co=doc_co,
            )
        action = "登録" if added else "更新"
        self._loaded_rows = self._current_grid_rows()
        ratio_text = self._company_match_ratio_text(self._loaded_rows)
        status = f"会社名対応を{action}しました: {file_co} ↔ {self._format_doc_company_for_message(doc_co)} / {ratio_text}"
        if applied:
            status += f" / 類似行 {applied} 件に文書会社を設定"
        self._status_var.set(status)

    def _open_company_alias_table_file(self) -> None:
        path = self._company_alias_table_path
        if not path.is_file():
            save_company_aliases(path, self._company_aliases)
        try:
            _win_shell_open_file(path)
        except OSError as e:
            messagebox.showerror(
                "会社名対応表",
                f"ファイルを開けませんでした。\n\n{e}",
                parent=self._root,
            )

    def _open_company_alias_editor(self) -> None:
        top = tk.Toplevel(self)
        top.title("会社名対応表")
        top.transient(self)
        top.geometry("720x360")

        frame = ttk.Frame(top, padding=8)
        frame.pack(fill=tk.BOTH, expand=True)

        cols = ("file_company", "doc_company", "note")
        tree = ttk.Treeview(
            frame,
            columns=cols,
            show="headings",
            selectmode="browse",
            height=10,
        )
        tree.heading("file_company", text="ファイル名会社")
        tree.heading("doc_company", text="文書会社")
        tree.heading("note", text="メモ")
        tree.column("file_company", width=220, stretch=True)
        tree.column("doc_company", width=220, stretch=True)
        tree.column("note", width=180, stretch=True)
        y_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=y_scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        working_aliases: list[dict[str, str]] = [
            dict(item) for item in self._company_aliases
        ]

        def refresh_tree() -> None:
            for iid in tree.get_children():
                tree.delete(iid)
            for idx, item in enumerate(working_aliases):
                tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(
                        item.get("file_company", ""),
                        item.get("doc_company", "") or "（空欄）",
                        item.get("note", ""),
                    ),
                )

        def prompt_pair(
            *,
            title: str,
            initial_file: str = "",
            initial_doc: str = "",
            initial_note: str = "",
        ) -> tuple[str, str, str] | None:
            dlg = tk.Toplevel(top)
            dlg.title(title)
            dlg.transient(top)
            dlg.grab_set()
            ttk.Label(dlg, text="ファイル名会社:").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
            file_var = tk.StringVar(value=initial_file)
            ttk.Entry(dlg, textvariable=file_var, width=48).grid(row=0, column=1, padx=8, pady=(8, 4))
            ttk.Label(dlg, text="文書会社:").grid(row=1, column=0, sticky="w", padx=8, pady=4)
            doc_var = tk.StringVar(value=initial_doc)
            ttk.Entry(dlg, textvariable=doc_var, width=48).grid(row=1, column=1, padx=8, pady=4)
            ttk.Label(dlg, text="メモ:").grid(row=2, column=0, sticky="w", padx=8, pady=4)
            note_var = tk.StringVar(value=initial_note)
            ttk.Entry(dlg, textvariable=note_var, width=48).grid(row=2, column=1, padx=8, pady=4)
            result: dict[str, tuple[str, str, str] | None] = {"value": None}

            def on_ok() -> None:
                fc = file_var.get().strip()
                dc = doc_var.get().strip()
                if not fc:
                    messagebox.showwarning("入力不足", "ファイル名会社を入力してください。", parent=dlg)
                    return
                result["value"] = (fc, dc, note_var.get().strip())
                dlg.destroy()

            def on_cancel() -> None:
                dlg.destroy()

            btns = ttk.Frame(dlg)
            btns.grid(row=3, column=0, columnspan=2, sticky="e", padx=8, pady=8)
            ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
            ttk.Button(btns, text="キャンセル", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 8))
            dlg.wait_window()
            return result["value"]

        def on_add() -> None:
            pair = prompt_pair(title="対応を追加")
            if pair is None:
                return
            fc, dc, note = pair
            add_company_alias(working_aliases, file_company=fc, doc_company=dc, note=note)
            refresh_tree()

        def on_edit() -> None:
            sel = tree.selection()
            if not sel:
                return
            idx = int(sel[0])
            item = working_aliases[idx]
            pair = prompt_pair(
                title="対応を編集",
                initial_file=str(item.get("file_company") or ""),
                initial_doc=str(item.get("doc_company") or ""),
                initial_note=str(item.get("note") or ""),
            )
            if pair is None:
                return
            fc, dc, note = pair
            item["file_company"] = fc
            item["doc_company"] = dc
            item["note"] = note
            refresh_tree()

        def on_delete() -> None:
            sel = tree.selection()
            if not sel:
                return
            idx = int(sel[0])
            del working_aliases[idx]
            refresh_tree()

        def on_save() -> None:
            try:
                save_company_aliases(self._company_alias_table_path, working_aliases)
            except OSError as e:
                messagebox.showerror("保存失敗", str(e), parent=top)
                return
            self._company_aliases = [dict(item) for item in working_aliases]
            self._recalculate_all_company_matches()
            self._loaded_rows = self._current_grid_rows()
            ratio_text = self._company_match_ratio_text(self._loaded_rows)
            self._status_var.set(f"会社名対応表を保存しました / {ratio_text}")
            top.destroy()

        def on_reload() -> None:
            self._reload_company_aliases()
            working_aliases[:] = [dict(item) for item in self._company_aliases]
            refresh_tree()

        btns = ttk.Frame(top, padding=(8, 0, 8, 8))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="追加", command=on_add).pack(side=tk.LEFT)
        ttk.Button(btns, text="編集", command=on_edit).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="削除", command=on_delete).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="ファイルを開く", command=self._open_company_alias_table_file).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Button(btns, text="再読み込み", command=on_reload).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="保存して閉じる", command=on_save).pack(side=tk.RIGHT)
        ttk.Button(btns, text="閉じる", command=top.destroy).pack(side=tk.RIGHT, padx=(0, 8))

        refresh_tree()

    def _recalculate_auto_judgment_for_row(
        self, rid: str, row: dict[str, str] | None = None
    ) -> None:
        """列編集後に自動判断を再計算し、ユーザ判断も同じ結果で上書きする。"""
        base = row if row is not None else self._current_row_dict_from_iid(rid)
        core = self._row_dict_to_core(self._merge_row_extra(rid, base))
        if self._is_match_company_manual_ok(core):
            core["match_company"] = "〇"
            core[self.MATCH_COMPANY_COL] = "〇"
        if not self._is_match_person_manual(core):
            core["match_person"] = recalculate_match_person_for_row(core)
        core[self.MATCH_PERSON_COL] = core.get("match_person", "")
        aj = auto_judgment_symbol(core)
        core["auto_judgment"] = aj
        core["user_judgment_company"] = aj
        core[self.FINAL_JUDGMENT_COL] = aj
        core[self.MATCH_DOC_TYPE_COL] = (
            core.get("match_doc_type")
            or core.get(self.MATCH_DOC_TYPE_COL)
            or ""
        )
        sync_row_ai_read_states(core)
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
                label=f"{self.TOTAL_HOURS_RAW_COL}を編集",
                command=lambda: self._prompt_edit_total_hours_raw(rid),
            )
            menu.add_separator()

        if cols[ci] == self.TRANSPORT_EXPENSE_COL:
            menu.add_command(
                label=f"{self.TRANSPORT_EXPENSE_COL}を編集",
                command=lambda: self._prompt_edit_transport_expense_raw(rid),
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

        if cols[ci] == self.MATCH_COMPANY_COL:
            menu.add_command(
                label="この対応をOK登録",
                command=lambda: self._register_company_alias_from_row(rid),
            )
            menu.add_command(
                label="一旦〇にする（対応表に登録しない）",
                command=lambda: self._set_match_company_temp_ok(rid),
            )
            menu.add_command(
                label="会社名対応表を編集",
                command=self._open_company_alias_editor,
            )
            menu.add_separator()

        if cols[ci] == self.MATCH_PERSON_COL:
            for label, inner in (("〇", "〇"), ("△", "△"), ("✖", "✖")):
                menu.add_command(
                    label=label,
                    command=lambda v=inner: self._set_match_person_symbol(rid, v),
                )
            menu.add_separator()

        menu.add_command(
            label="行を削除",
            command=lambda: self._prompt_delete_grid_rows(rid),
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
        return self._merge_row_extra(rid, row)

    def _replace_row_with_result(
        self,
        rid: str,
        row: dict[str, str],
        *,
        sync_user_judgment_to_auto: bool = False,
    ) -> None:
        self._tree.item(
            rid,
            values=self._merge_preserved_row_no(
                rid,
                self._grid_values_from_row(
                    row, sync_user_judgment_to_auto=sync_user_judgment_to_auto
                ),
            ),
        )
        self._item_paths[rid] = row.get("resolved_path", "")
        self._sync_row_extra_from_row(rid, row)

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
        self._export_csv_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._restore_excluded_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._error_reanalysis_btn.configure(state=tk.DISABLED)
        self._billing_prepare_btn.configure(state=tk.DISABLED)
        self._billing_delete_btn.configure(state=tk.DISABLED)
        self._billing_review_btn.configure(state=tk.DISABLED)
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
                    company_aliases=self._company_aliases,
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
                self._refresh_duplicate_employee_no_highlight()
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
        self._export_csv_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._restore_excluded_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._error_reanalysis_btn.configure(state=tk.DISABLED)
        self._billing_prepare_btn.configure(state=tk.DISABLED)
        self._billing_delete_btn.configure(state=tk.DISABLED)
        self._billing_review_btn.configure(state=tk.DISABLED)
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
                    company_aliases=self._company_aliases,
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
                self._refresh_duplicate_employee_no_highlight()

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
            self._refresh_duplicate_employee_no_highlight()
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
        top.title(f"{self.TOTAL_HOURS_RAW_COL}を編集")
        top.transient(self)
        top.grab_set()

        ttk.Label(
            top,
            text=(
                f"{self.TOTAL_HOURS_RAW_COL}を入力してください。\n"
                "例: 8:20 / 8時間20分 / 8.20 / 101_10H / 86.17H / 0 / 0時間\n"
                "空欄は（データなし）として保存します（AI読取状態: データなし）。\n"
                "入力後、10進列・AI読取状態を同じ規則で自動更新します。"
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
            row["total_hours_raw"] = new_raw or "（データなし）"
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

    def _prompt_edit_transport_expense_raw(self, rid: str) -> None:
        """交通費合計（読取）を右クリックから編集し、更新用交通費合計も連動更新する。"""
        try:
            te_ci = self._transport_expense_column_index()
        except ValueError:
            return

        vals = list(self._tree.item(rid, "values") or [])
        cur = vals[te_ci] if te_ci < len(vals) else ""

        top = tk.Toplevel(self)
        top.title(f"{self.TRANSPORT_EXPENSE_COL}を編集")
        top.transient(self)
        top.grab_set()

        ttk.Label(
            top,
            text=(
                f"{self.TRANSPORT_EXPENSE_COL}を入力してください。\n"
                "例: 12,345円 / 5000 / 0 / 0円\n"
                "記載がない場合は（データなし）と入力するか、空欄のまま OK してください。\n"
                "（保存時は（データなし）、AI読取状態: データなし）。\n"
                "0・0円はゼロとして扱います。更新用交通費合計・AI読取状態を自動反映します"
                "（更新用勤務時間は変更しません）。"
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
            stored = new_raw or "（データなし）"
            row = self._row_dict_to_core(self._current_row_dict_from_iid(rid))
            row["transport_expense_raw"] = stored
            row[self.TRANSPORT_EXPENSE_COL] = stored
            transport_billing = _normalize_billing_update_copy(stored)
            row["billing_update_transport"] = transport_billing
            row[self.BILLING_UPDATE_TRANSPORT_COL] = transport_billing
            mdt = recalculate_match_doc_type_for_row(row)
            row["match_doc_type"] = mdt
            row[self.MATCH_DOC_TYPE_COL] = mdt
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
            no_val = self._tree_row_no(iid)
            if no_val:
                row["grid_row_no"] = no_val
            row.pop(self.ROW_NO_COL, None)
            # 内部データ
            row["resolved_path"] = self._item_paths.get(iid, "")
            row = self._merge_row_extra(iid, row)
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
            if is_match_company_ok_for_ratio(symbol):
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

    def _export_grid_csv(self) -> None:
        if self._busy:
            return
        iids = self._grid_export_target_iids()
        if not iids:
            messagebox.showinfo(
                "CSVエクスポート",
                "エクスポートする行がありません。",
                parent=self._root,
            )
            return
        self._prepare_native_dialog()
        fp = filedialog.asksaveasfilename(
            title="グリッドをCSVで保存",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("すべて", "*.*")],
            initialfile=f"kintai_grid_{self._now_ts()}.csv",
            parent=self._root,
        )
        if not fp:
            return
        out = Path(fp)
        cols = list(self._tree["columns"])
        header = [self._column_heading_display_text(h) for h in cols]
        try:
            with out.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(header)
                for iid in iids:
                    values = list(self._tree.item(iid, "values") or ())
                    row = [
                        str(values[i]) if i < len(values) and values[i] is not None else ""
                        for i in range(len(cols))
                    ]
                    writer.writerow(row)
        except OSError as e:
            messagebox.showerror(
                "CSVエクスポート",
                f"ファイルを保存できませんでした。\n\n{e}",
                parent=self._root,
            )
            return
        scope = (
            f"選択行 {len(iids)} 行"
            if self._selected_tree_iids()
            else f"全行 {len(iids)} 行"
        )
        self._status_var.set(f"CSVをエクスポートしました（{scope}）: {out}")

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
        self._grid_row_no_seq = 0
        for r in rows:
            vals = self._grid_values_from_row(r)
            iid = self._tree.insert("", tk.END, values=vals)
            self._item_paths[iid] = (r.get("resolved_path") or "").strip()
            self._sync_row_extra_from_row(iid, r)
            saved_no = str(r.get("grid_row_no") or "").strip()
            if saved_no.isdigit():
                self._assign_row_no_to_iid(iid, int(saved_no))
            else:
                self._assign_row_no_to_iid(iid)
        self._sync_grid_row_no_seq_from_tree()
        self._refresh_duplicate_employee_no_highlight()

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
        save_last_assistant_name(
            self._user_settings_path, (self._assistant_var.get() or "").strip()
        )
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
        # 追加継続解析: 読み込み済み（または現在表示）の状態を起点
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
        self._export_csv_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._restore_excluded_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._error_reanalysis_btn.configure(state=tk.DISABLED)
        self._billing_prepare_btn.configure(state=tk.DISABLED)
        self._billing_delete_btn.configure(state=tk.DISABLED)
        self._billing_review_btn.configure(state=tk.DISABLED)
        self._billing_update_btn.configure(state=tk.DISABLED)
        self._progress_var.set("")
        self._status_var.set("解析を準備しています…")

        # 追加継続解析のときは、まず既存行をグリッドに反映（ユーザ判断も含む）
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
            # 追加継続解析時は読み込み済み件数を加算して表示
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
                self._sync_row_extra_from_row(iid, row)
                self._assign_row_no_to_iid(iid)
                try:
                    self._tree.yview_moveto(1)
                except tk.TclError:
                    pass
                refresh_running_ratio_status()

            self.after(0, lambda rr=r: append_row(rr))

        def _as_core_row(r: dict[str, str]) -> dict[str, str]:
            """UI行(dict) -> kintai_core の row(dict) に寄せる"""
            # 追加継続解析時に既存行の情報を落とさないよう、UI行をできるだけ保持したまま
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

            # 起点行（追加継続解析）: file_name -> row
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
                    company_aliases=self._company_aliases,
                )
            except BaseException as e:
                err = e
                rows_result = None

            rows_final = rows_result or []

            # 追加継続解析: base を起点に、結果を追加（既存は skip_file_names でスキップされる想定）
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

                self._refresh_duplicate_employee_no_highlight()

                # --- UIボタン復帰 ---
                self._update_data_dir_dependent_buttons()
                self._load_btn.configure(state=tk.NORMAL)
                if self._loaded_rows:
                    self._save_btn.configure(state=tk.NORMAL)
                    self._export_csv_btn.configure(state=tk.NORMAL)

                # --- ステータス表示 ---
                n_rows = len(self._loaded_rows or [])
                ratio_text = self._company_match_ratio_text(self._loaded_rows)
                if err is not None:
                    # 要件: 送信エラー等でも中断状態・保存/追加継続解析を有効にする
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