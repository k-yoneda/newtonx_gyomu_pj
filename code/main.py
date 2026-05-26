"""勤務表画像・PDF 解析ツールの Windows GUI エントリ。"""

from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

_CODE_DIR = Path(__file__).resolve().parent
_ROOT = _CODE_DIR.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from kintai_core import (
    DEFAULT_PARALLEL_ANALYSIS_CHATS,
    PARALLEL_WORKERS_MAX,
    TARGET_ASSISTANT_NAME,
    _decimal_for_table_display,
    _work_hours_string_to_decimal,
    auto_judgment_symbol,
    create_client,
    is_manual_user_judgment,
    normalize_judgment_symbol,
    row_display_values,
    run_analysis,
    summary_header_cells,
)
from newtonx_adk.exceptions import APIError


class KintaiApp(tk.Frame):
    USER_JUDGMENT_COL = "ユーザ判断"
    AUTO_JUDGMENT_COL = "自動判断"
    EMPLOYEE_NO_COL = "社員番号"
    TOTAL_HOURS_DECIMAL_COL = "合計勤務時間（10進）"
    TOTAL_HOURS_RAW_COL = "合計勤務時間（読取）"
    MATCH_COMPANY_COL = "会社名比較"
    LEGACY_MATCH_COMPANY_COL = "会社名比較（ファイル名✖文書）"
    COMPANY1_COL = "会社名1"
    # 「エラー再解析」対象: ユーザ判断が「〇」以外の行
    _ERROR_REANALYSIS_OK_VALUES = ("〇",)
    # 記号列（〇/△/✖ 等）の列幅（px）
    _SYMBOL_COLUMN_WIDTHS: dict[str, int] = {
        "アップロード": 72,
        "対象シート有無": 88,
        USER_JUDGMENT_COL: 72,
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
        assistant_uid: str,
        parallel_workers: int = DEFAULT_PARALLEL_ANALYSIS_CHATS,
    ) -> None:
        super().__init__(master)
        self._root = master
        self._root.title("勤務表解析")
        self._root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._client = client
        self._assistant_uid = assistant_uid
        nw = int(parallel_workers)
        self._parallel_workers = max(1, min(nw, PARALLEL_WORKERS_MAX))
        self._data_dir: Path | None = None
        self._busy = False
        self._item_paths: dict[str, str] = {}
        self._cancel_event: threading.Event | None = None

        self._loaded_rows: list[dict[str, str]] = []
        self._loaded_json_path: Path | None = None
        self._last_saved_snapshot: str = ""

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

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        self._folder_var = tk.StringVar(value="（未選択）")
        ttk.Label(top, text="データフォルダ:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self._folder_var, anchor="w").grid(
            row=0, column=1, sticky="ew", padx=(4, 4)
        )
        ttk.Button(top, text="参照…", command=self._browse_folder).grid(
            row=0, column=2, sticky="e"
        )
        top.columnconfigure(1, weight=1)

        ctrl = ttk.Frame(self, padding=(8, 0, 8, 8))
        ctrl.pack(fill=tk.X)

        self._new_btn = ttk.Button(
            ctrl, text="新規解析", command=self._start_new_analysis, state=tk.DISABLED
        )
        self._new_btn.grid(row=0, column=0, sticky="w")

        self._cont_btn = ttk.Button(
            ctrl, text="継続解析", command=self._start_continue_analysis, state=tk.DISABLED
        )
        self._cont_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self._save_btn = ttk.Button(
            ctrl, text="保存", command=self._save_json, state=tk.DISABLED
        )
        self._save_btn.grid(row=0, column=2, sticky="w", padx=(8, 0))

        self._load_btn = ttk.Button(
            ctrl, text="読み込み", command=self._load_json, state=tk.NORMAL
        )
        self._load_btn.grid(row=0, column=3, sticky="w", padx=(8, 0))

        self._cancel_btn = ttk.Button(
            ctrl, text="中断", command=self._cancel_analysis, state=tk.DISABLED
        )
        self._cancel_btn.grid(row=0, column=4, sticky="w", padx=(16, 0))

        self._progress_var = tk.StringVar(value="")
        self._status_var = tk.StringVar(value="準備完了")
        # 「実行済 100 / 対象 120」など3桁になっても欠けないよう、表示幅を広げる
        # ttk.Label の width は“文字数”ベースなので、minsize と合わせて余裕を持たせる。
        # （環境によってフォントが少し太く、26文字だと末尾が欠けるケースがあったため更に増やす）
        ttk.Label(ctrl, textvariable=self._progress_var, width=30).grid(
            row=0, column=5, sticky="w", padx=(16, 0)
        )
        # 進捗表示（実行済/対象）は桁数により伸びるため、最低幅を確保して欠けを防ぐ
        ctrl.columnconfigure(5, minsize=240)

        # ステータスは可変長だが、要求幅が大きくなりすぎると右端ボタンが見えなくなる。
        # 右端に必ず「エラー再解析」を表示するため、ラベルの要求幅を抑制（固定幅＋折り返し）する。
        self._status_label = ttk.Label(
            ctrl,
            textvariable=self._status_var,
            anchor="w",
            width=70,
            wraplength=900,
            justify="left",
        )
        self._status_label.grid(row=0, column=6, sticky="ew", padx=(12, 0))
        ctrl.columnconfigure(6, weight=1)
        # 右端ボタン領域の最低幅を確保
        ctrl.columnconfigure(7, minsize=120)

        # 右側: エラー再解析（ユーザ判断が「〇」以外の行のみを対象に再解析）
        self._error_reanalysis_btn = ttk.Button(
            ctrl,
            text="エラー再解析",
            command=self._start_error_reanalysis,
            state=tk.DISABLED,
        )
        self._error_reanalysis_btn.grid(row=0, column=7, sticky="e", padx=(12, 0))

        grid_frame = ttk.Frame(self, padding=(8, 0, 8, 8))
        grid_frame.pack(fill=tk.BOTH, expand=True)

        headings = summary_header_cells()
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
            self._tree.heading(h, text=h, anchor="w")

        self._tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        grid_frame.rowconfigure(0, weight=1)
        grid_frame.columnconfigure(0, weight=1)

        self._tree.bind("<Button-3>", self._on_tree_right_click)
        self._tree.bind("<Double-1>", self._on_row_double_click)

        initial_w = self._window_width_for_columns(col_px, y_scroll=y_scroll)
        min_w = min(960, initial_w)
        self._root.minsize(min_w, 520)
        self._root.geometry(f"{initial_w}x680")

    def _browse_folder(self) -> None:
        self._prepare_native_dialog()
        d = filedialog.askdirectory(
            title="データを読み込むフォルダを選択",
            parent=self._root,
        )
        if not d:
            return
        self._data_dir = Path(d)
        self._folder_var.set(str(self._data_dir.resolve()))
        if not self._busy:
            self._progress_var.set("")
            self._new_btn.configure(state=tk.NORMAL)
            self._cont_btn.configure(
                state=(tk.NORMAL if self._loaded_rows else tk.DISABLED)
            )
            self._save_btn.configure(state=(tk.NORMAL if self._tree.get_children() else tk.DISABLED))
            self._cancel_btn.configure(state=tk.DISABLED)
            self._refresh_error_reanalysis_button_state()

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
                row.get(self.USER_JUDGMENT_COL)
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

    def _refresh_error_reanalysis_button_state(self) -> None:
        if self._busy or self._data_dir is None:
            self._error_reanalysis_btn.configure(state=tk.DISABLED)
            return
        # 対象行が1件でもあれば有効
        enabled = bool(self._eligible_error_reanalysis_iids())
        self._error_reanalysis_btn.configure(state=(tk.NORMAL if enabled else tk.DISABLED))

    def _clear_grid(self) -> None:
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._item_paths.clear()

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

    def _row_dict_to_core(self, row: dict[str, str]) -> dict[str, str]:
        """Treeview/JSON 行を kintai_core 互換 dict に変換する。"""
        out = dict(row)
        pairs = (
            ("画像ファイル名", "file_name"),
            ("アップロード", "upload_ok"),
            ("対象シート有無", "target_sheet_exists"),
            (self.USER_JUDGMENT_COL, "user_judgment_company"),
            (self.COMPANY1_COL, "name_company_1"),
            ("氏名", "name_person_from_doc"),
            (self.EMPLOYEE_NO_COL, "employee_no"),
            (self.TOTAL_HOURS_DECIMAL_COL, "total_hours_decimal"),
            (self.TOTAL_HOURS_RAW_COL, "total_hours_raw"),
            (self.MATCH_COMPANY_COL, "match_company"),
            ("押印有無", "seal_in_doc"),
        )
        for ui_key, core_key in pairs:
            if ui_key in row:
                val = str(row.get(ui_key) or "").strip()
            elif core_key == "match_company":
                val = str(
                    row.get(core_key)
                    or row.get(self.LEGACY_MATCH_COMPANY_COL)
                    or ""
                ).strip()
            else:
                val = str(row.get(core_key) or "").strip()
            if core_key == "user_judgment_company":
                out[core_key] = normalize_judgment_symbol(val) if val else ""
            else:
                out[core_key] = val
        return out

    def _grid_values_from_row(self, row: dict[str, str]) -> tuple[str, ...]:
        """内部row形式・UI保存形式のどちらでも Treeview 表示値へ変換する。"""
        return row_display_values(self._row_dict_to_core(row))

    def _user_judgment_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.USER_JUDGMENT_COL)

    def _employee_no_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.EMPLOYEE_NO_COL)

    def _total_hours_decimal_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.TOTAL_HOURS_DECIMAL_COL)

    def _total_hours_raw_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.TOTAL_HOURS_RAW_COL)

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

        # ユーザ判断列: 右クリックで 〇/△/✖ を選択
        if cols[ci] == self.USER_JUDGMENT_COL:
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

    def _replace_row_with_result(self, rid: str, row: dict[str, str]) -> None:
        self._tree.item(rid, values=row_display_values(row))
        self._item_paths[rid] = row.get("resolved_path", "")

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
        if self._busy or self._data_dir is None:
            return

        current_row = self._current_row_dict_from_iid(rid)

        # 選択行の「再解析」は、会社名1の内容に関わらず実行可能とする。
        # （不明/（存在しない）のみ一括で再解析したい場合は「エラー再解析」ボタンを使用。）

        file_name = (current_row.get("画像ファイル名") or current_row.get("file_name") or "").strip()
        if not file_name:
            messagebox.showinfo("再解析", "選択行のファイル名を取得できません。")
            return

        td = self._data_dir.resolve()
        client = self._client
        aid = self._assistant_uid

        self._busy = True
        self._cancel_event = threading.Event()
        self._new_btn.configure(state=tk.DISABLED)
        self._cont_btn.configure(state=tk.DISABLED)
        self._save_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
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

                self._new_btn.configure(state=(tk.NORMAL if self._data_dir else tk.DISABLED))
                self._cont_btn.configure(state=(tk.NORMAL if (self._data_dir and (self._loaded_rows or self._tree.get_children())) else tk.DISABLED))
                self._save_btn.configure(state=(tk.NORMAL if self._tree.get_children() else tk.DISABLED))
                self._load_btn.configure(state=tk.NORMAL)

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
                self._refresh_error_reanalysis_button_state()
                messagebox.showinfo("再解析完了", f"選択行の再解析が完了しました。\n{file_name}")

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _start_error_reanalysis(self) -> None:
        """ユーザ判断が「〇」以外の行だけをまとめて再解析する。"""
        if self._busy or self._data_dir is None:
            return

        target_iids = self._eligible_error_reanalysis_iids()
        if not target_iids:
            messagebox.showinfo("エラー再解析", "対象行がありません（ユーザ判断が『〇』以外の行）。")
            self._refresh_error_reanalysis_button_state()
            return

        # iid -> file_name
        iid_by_file: dict[str, str] = {}
        for iid in target_iids:
            row = self._current_row_dict_from_iid(iid)
            file_name = (row.get("画像ファイル名") or row.get("file_name") or "").strip()
            if not file_name:
                continue
            iid_by_file[file_name] = iid

        if not iid_by_file:
            messagebox.showwarning("エラー再解析", "対象行のファイル名を取得できませんでした。")
            self._refresh_error_reanalysis_button_state()
            return

        file_names = set(iid_by_file.keys())
        total = len(file_names)

        td = self._data_dir.resolve()
        client = self._client
        aid = self._assistant_uid

        self._busy = True
        self._cancel_event = threading.Event()
        self._new_btn.configure(state=tk.DISABLED)
        self._cont_btn.configure(state=tk.DISABLED)
        self._save_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._error_reanalysis_btn.configure(state=tk.DISABLED)
        self._progress_var.set(f"エラー再解析中 0 / {total}")
        self._status_var.set(f"エラー再解析しています… {total} 件")

        # 念のため、前回の異常終了等で反転が残っていた場合に備えて全解除してから開始する。
        self._set_rows_reanalysis_highlight(target_iids, False)

        # 「現在再解析中の1行」だけを反転表示する。
        # ※複数並列だと同時に複数行が“処理中”になり得るため、ここでは逐次処理(1並列)で運用する。
        current_active: dict[str, str] = {"rid": ""}

        def log_line(message: str) -> None:
            def apply_log(m: str = message) -> None:
                if self._should_ignore_status_log(m):
                    return
                self._status_var.set(m[:800])

            self.after(0, apply_log)

        def on_progress(done: int, t: int) -> None:
            self.after(0, lambda d=done, tt=t: self._progress_var.set(f"エラー再解析中 {d} / {tt}"))

        def on_file_started(file_name: str) -> None:
            """core側が「このファイルの処理を開始した」タイミングで呼ばれる。

            GUI側では、今処理中の1行のみ反転表示する。
            """
            fn = (file_name or "").strip()
            if not fn:
                return
            rid = iid_by_file.get(fn) or ""
            if not rid:
                return

            def apply_started() -> None:
                prev = current_active.get("rid") or ""
                if prev and prev != rid:
                    self._set_row_reanalysis_highlight(prev, False)
                current_active["rid"] = rid
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
            # 再解析時は自動判断でユーザ判断を上書きする
            row["user_judgment_company"] = auto_judgment_symbol(row)

            rid = iid_by_file.get(fn)
            if not rid:
                return

            def apply_row() -> None:
                self._replace_row_with_result(rid, row)
                # この行の処理が終わったので反転を戻す（次の on_file_started で次行が反転する）
                self._set_row_reanalysis_highlight(rid, False)
                if (current_active.get("rid") or "") == rid:
                    current_active["rid"] = ""
                # 途中経過でも割合等を更新
                current_rows = self._current_grid_rows()
                ratio_text = self._company_match_ratio_text(current_rows, total_target_count=len(current_rows))
                self._status_var.set(f"エラー再解析中… / {ratio_text}")

            self.after(0, apply_row)

        def worker() -> None:
            err: BaseException | None = None
            try:
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
                    parallel_chats=1,
                )
            except BaseException as e:
                err = e

            def finish() -> None:
                self._busy = False
                self._cancel_btn.configure(state=tk.DISABLED)
                cancelled = self._cancel_event is not None and self._cancel_event.is_set()
                self._cancel_event = None

                # 反転表示を解除（成功/エラー/中断いずれも）
                cur = (current_active.get("rid") or "").strip()
                if cur:
                    self._set_row_reanalysis_highlight(cur, False)
                self._set_rows_reanalysis_highlight(target_iids, False)

                self._new_btn.configure(state=(tk.NORMAL if self._data_dir else tk.DISABLED))
                self._cont_btn.configure(state=(tk.NORMAL if (self._data_dir and (self._loaded_rows or self._tree.get_children())) else tk.DISABLED))
                self._save_btn.configure(state=(tk.NORMAL if self._tree.get_children() else tk.DISABLED))
                self._load_btn.configure(state=tk.NORMAL)

                # 確定: 現在グリッドの内容を loaded_rows に反映
                self._loaded_rows = self._current_grid_rows()
                self._progress_var.set(f"エラー再解析完了 {total} / {total}")
                ratio_text = self._company_match_ratio_text(self._loaded_rows)

                # ボタン状態更新
                self._refresh_error_reanalysis_button_state()

                if err is not None:
                    messagebox.showerror("エラー再解析エラー", str(err))
                    self._status_var.set(f"エラー再解析エラー（途中結果は保持） / {ratio_text}")
                    return
                if cancelled:
                    self._status_var.set(f"エラー再解析を中断しました（途中結果は保持） / {ratio_text}")
                    return

                self._status_var.set(f"エラー再解析完了: {total} 件 / {ratio_text}")
                messagebox.showinfo("エラー再解析完了", f"ユーザ判断が〇以外の行を再解析しました。\n対象: {total} 件")

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

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
            if ci >= len(vals):
                while len(vals) <= ci:
                    vals.append("")
            vals[ci] = new_v
            self._tree.item(rid, values=tuple(vals))
            top.destroy()

        def on_cancel() -> None:
            top.destroy()

        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
        ttk.Button(btns, text="キャンセル", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 8))

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
            new_dec = _decimal_for_table_display(_work_hours_string_to_decimal(new_raw)) if new_raw else "（なし）"
            max_ci = max(raw_ci, dec_ci)
            if max_ci >= len(vals):
                while len(vals) <= max_ci:
                    vals.append("")
            vals[raw_ci] = new_raw or "（なし）"
            vals[dec_ci] = new_dec
            self._tree.item(rid, values=tuple(vals))
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

    def _write_json_file(self, path: Path, rows: list[dict[str, str]]) -> None:
        payload = {
            "version": 1,
            "data_dir": str(self._data_dir.resolve()) if self._data_dir else "",
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

        # data_dir は読み込みJSONを優先してセット（空なら維持）
        dd = (data.get("data_dir") or "").strip()
        if dd:
            self._data_dir = Path(dd)
            self._folder_var.set(str(self._data_dir.resolve()))

        self._rebuild_grid_from_rows(self._loaded_rows)
        self._save_btn.configure(state=(tk.NORMAL if self._loaded_rows else tk.DISABLED))
        if not self._busy:
            self._new_btn.configure(state=(tk.NORMAL if self._data_dir else tk.DISABLED))
            self._cont_btn.configure(state=(tk.NORMAL if (self._data_dir and self._loaded_rows) else tk.DISABLED))
        self._refresh_error_reanalysis_button_state()
        ratio_text = self._company_match_ratio_text(self._loaded_rows)
        self._status_var.set(f"読み込みました: {self._loaded_json_path} / {ratio_text}")
        self._update_title()

    def _rebuild_grid_from_rows(self, rows: list[dict[str, str]]) -> None:
        self._clear_grid()
        for r in rows:
            vals = self._grid_values_from_row(r)
            iid = self._tree.insert("", tk.END, values=vals)
            self._item_paths[iid] = (r.get("resolved_path") or "").strip()

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

    def _auto_save_on_exit(self) -> None:
        rows = self._current_grid_rows()
        if not rows:
            return
        # 終了時の自動保存は常に日時付きの新規ファイル
        out_name = self._suggest_timestamped_name()
        out_path = Path.cwd() / out_name
        self._write_json_file(out_path, rows)
        self._loaded_json_path = out_path
        self._loaded_rows = rows
        self._last_saved_snapshot = self._compute_snapshot(rows)

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
        self._root.destroy()

    def _start_new_analysis(self) -> None:
        # 新規解析: グリッドをクリアして最初から
        self._loaded_rows = []
        self._loaded_json_path = None
        self._last_saved_snapshot = ""
        self._update_title()
        self._start_analysis(base_rows=None)

    def _start_continue_analysis(self) -> None:
        # 継続解析: 読み込み済み（または現在表示）の状態を起点
        base = self._loaded_rows or self._current_grid_rows()
        self._start_analysis(base_rows=base)

    def _start_analysis(self, *, base_rows: list[dict[str, str]] | None) -> None:
        if self._busy or self._data_dir is None:
            return
        self._busy = True
        self._cancel_event = threading.Event()
        self._new_btn.configure(state=tk.DISABLED)
        self._cont_btn.configure(state=tk.DISABLED)
        self._save_btn.configure(state=tk.DISABLED)
        self._load_btn.configure(state=tk.DISABLED)
        self._cancel_btn.configure(state=tk.NORMAL)
        self._progress_var.set("")
        self._status_var.set("解析を準備しています…")

        # 継続解析のときは、まず既存行をグリッドに反映（ユーザ判断も含む）
        if base_rows is None:
            self._clear_grid()
        else:
            self._rebuild_grid_from_rows(base_rows)

        td = self._data_dir.resolve()
        client = self._client
        aid = self._assistant_uid

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
                vals = row_display_values(row)
                iid = self._tree.insert("", tk.END, values=vals)
                self._item_paths[iid] = row.get("resolved_path", "")
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
            out["file_name"] = (r.get("画像ファイル名") or r.get("file_name") or "").strip()
            out["resolved_path"] = (r.get("resolved_path") or "").strip()
            uj = (r.get(self.USER_JUDGMENT_COL) or r.get("user_judgment_company") or "").strip()
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
                            or base_map[fn].get(self.USER_JUDGMENT_COL)
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
                        fn = (br.get("画像ファイル名") or br.get("file_name") or "").strip()
                        if fn:
                            skip_names.add(fn)

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
                    parallel_chats=self._parallel_workers,
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
                if self._data_dir is not None:
                    self._new_btn.configure(state=tk.NORMAL)
                else:
                    self._new_btn.configure(state=tk.DISABLED)

                # 継続解析は「途中結果がある」ならエラー/中断でも有効
                self._cont_btn.configure(state=(tk.NORMAL if self._loaded_rows else tk.DISABLED))
                # 保存も同様に有効
                self._save_btn.configure(state=(tk.NORMAL if self._loaded_rows else tk.DISABLED))
                self._load_btn.configure(state=tk.NORMAL)

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
                    self._refresh_error_reanalysis_button_state()
                    return

                if cancelled:
                    self._status_var.set(
                        f"中断しました: {n_rows} 件を保持（JSON保存で続きから再開できます） / {ratio_text}"
                    )
                else:
                    self._status_var.set(
                        f"完了: {n_rows} 件をグリッド表示（解析結果.md を出力） / {ratio_text}"
                    )

                self._refresh_error_reanalysis_button_state()

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _on_row_double_click(self, event: tk.Event) -> None:
        if self._tree.identify_region(event.x, event.y) == "cell":
            col_w = self._tree.identify_column(event.x)
            try:
                ci = int(col_w.replace("#", "")) - 1
            except ValueError:
                ci = -1
            cols = list(self._tree["columns"])
            if 0 <= ci < len(cols) and cols[ci] == self.USER_JUDGMENT_COL:
                return
        rid = self._tree.identify_row(event.y)
        if not rid:
            return
        path_str = self._item_paths.get(rid, "").strip()
        if not path_str:
            messagebox.showinfo(
                "ファイルを開けません",
                "この行に関連するファイルパスがありません。",
            )
            return
        p = Path(path_str)
        if not p.is_file():
            messagebox.showerror(
                "ファイルがありません",
                f"見つかりません: {p}",
            )
            return
        try:
            os.startfile(os.path.normpath(str(p)))
        except OSError as e:
            messagebox.showerror(
                "起動できませんでした",
                str(e),
            )


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
    # withdraw() のルートは環境により messagebox / Toplevel が出ず、メインにも進めないことがある。
    # 画面外の 1x1 として「表示済み」の親にする。
    root.geometry("1x1+-10000+-10000")
    root.deiconify()
    root.update_idletasks()

    client = create_client()
    if not client.authenticate():
        messagebox.showerror(
            "認証エラー",
            "NewtonX の認証に失敗しました。",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    # --- アシスタント選択ダイアログ（コンボボックス） ---
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

    assistant_names = [str(a.get("name") or "").strip() for a in assistants]
    assistant_names = [n for n in assistant_names if n]

    if not assistant_names:
        messagebox.showerror(
            "エラー",
            "アシスタント一覧を取得できませんでした。",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    dlg = tk.Toplevel(root)
    dlg.title("アシスタント・並列設定")
    dlg.transient(root)
    dlg.grab_set()

    ttk.Label(dlg, text="使用するアシスタントを選択してください:").pack(
        fill=tk.X, padx=12, pady=(12, 6)
    )

    sel_var = tk.StringVar(value="")
    cb = ttk.Combobox(dlg, textvariable=sel_var, values=assistant_names, state="readonly")
    cb.pack(fill=tk.X, padx=12)

    # デフォルト: TARGET_ASSISTANT_NAME があればそれ、なければ先頭
    default_name = TARGET_ASSISTANT_NAME if TARGET_ASSISTANT_NAME in assistant_names else assistant_names[0]
    sel_var.set(default_name)

    wf = ttk.Frame(dlg)
    wf.pack(fill=tk.X, padx=12, pady=(12, 0))
    ttk.Label(wf, text="ワーカースレッド数（並列チャット数）:").pack(side=tk.LEFT)
    workers_var = tk.StringVar(value=str(DEFAULT_PARALLEL_ANALYSIS_CHATS))
    tk.Spinbox(
        wf,
        from_=1,
        to=PARALLEL_WORKERS_MAX,
        textvariable=workers_var,
        width=6,
        justify="center",
    ).pack(side=tk.LEFT, padx=(8, 0))
    ttk.Label(wf, text=f"（1〜{PARALLEL_WORKERS_MAX}、既定 {DEFAULT_PARALLEL_ANALYSIS_CHATS}）").pack(
        side=tk.LEFT, padx=(8, 0)
    )

    btns = ttk.Frame(dlg)
    btns.pack(fill=tk.X, padx=12, pady=12)

    chosen_name: dict[str, str | None] = {"value": None}
    chosen_workers: dict[str, int] = {"value": DEFAULT_PARALLEL_ANALYSIS_CHATS}

    def on_ok() -> None:
        chosen_name["value"] = (sel_var.get() or "").strip() or default_name
        try:
            nw = int(str(workers_var.get()).strip())
        except (TypeError, ValueError):
            nw = DEFAULT_PARALLEL_ANALYSIS_CHATS
        chosen_workers["value"] = max(1, min(nw, PARALLEL_WORKERS_MAX))
        try:
            dlg.grab_release()
        except tk.TclError:
            pass
        dlg.destroy()

    def on_cancel() -> None:
        chosen_name["value"] = None
        try:
            dlg.grab_release()
        except tk.TclError:
            pass
        dlg.destroy()

    ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
    ttk.Button(btns, text="キャンセル", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 8))

    dlg.bind("<Return>", lambda _e: on_ok())
    dlg.bind("<Escape>", lambda _e: on_cancel())
    cb.focus_set()

    dlg.update_idletasks()
    _center_window_on_screen(dlg)
    dlg.lift()
    dlg.focus_force()

    root.wait_window(dlg)

    try:
        root.grab_release()
    except tk.TclError:
        pass
    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.focus_force()

    if not chosen_name["value"]:
        root.destroy()
        sys.exit(0)

    selected_name = str(chosen_name["value"])
    selected = next(
        (a for a in assistants if str(a.get("name") or "").strip() == selected_name),
        None,
    )
    assistant_uid = ""
    if selected:
        raw = selected.get("uid")
        if raw is None or str(raw).strip() == "":
            raw = selected.get("uuid")
        assistant_uid = str(raw).strip() if raw is not None else ""

    if not assistant_uid:
        messagebox.showerror(
            "エラー",
            f"アシスタント「{selected_name}」が見つかりません。",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    # Windows では「最初の Tk を destroy したあとに 2 つ目の Tk を作る」と
    # メインウィンドウが表示されないことがあるため、起動〜本画面まで同一の root を使う。
    app = KintaiApp(
        root,
        client=client,
        assistant_uid=assistant_uid,
        parallel_workers=chosen_workers["value"],
    )
    app.pack(fill=tk.BOTH, expand=True)
    root.update_idletasks()
    _center_window_on_screen(root)
    root.lift()
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
