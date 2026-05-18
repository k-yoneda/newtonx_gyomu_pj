"""勤務表画像・PDF 解析ツールの Windows GUI エントリ。"""

from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

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
    create_client,
    row_display_values,
    run_analysis,
    summary_header_cells,
)
from newtonx_adk.exceptions import APIError


class KintaiApp(tk.Frame):
    USER_JUDGMENT_COL = "ユーザ判断"
    EMPLOYEE_NO_COL = "社員番号"
    TOTAL_HOURS_DECIMAL_COL = "合計勤務時間（10進）"
    TOTAL_HOURS_RAW_COL = "合計勤務時間（読取）"

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
        ttk.Label(ctrl, textvariable=self._progress_var, width=18).grid(
            row=0, column=5, sticky="w", padx=(16, 0)
        )
        ttk.Label(ctrl, textvariable=self._status_var).grid(
            row=0, column=6, sticky="w", padx=(12, 0)
        )

        grid_frame = ttk.Frame(self, padding=(8, 0, 8, 8))
        grid_frame.pack(fill=tk.BOTH, expand=True)

        headings = summary_header_cells()
        y_scroll = ttk.Scrollbar(grid_frame)
        x_scroll = ttk.Scrollbar(grid_frame, orient=tk.HORIZONTAL)

        # headings（列名）と col_px（列幅）の数がズレると zip(strict=True) で落ちるため、
        # headings の列数に追従して幅リストを整形する。
        default_w_px = 120
        col_px_base = [56, 320, 200, 120, 120, 140, 80, 220, 72]
        if len(col_px_base) < len(headings):
            col_px = col_px_base + [default_w_px] * (len(headings) - len(col_px_base))
        else:
            col_px = col_px_base[: len(headings)]

        total_w = sum(col_px) + 28
        initial_w = max(1920, total_w)
        min_w = min(960, initial_w)
        self._tree = ttk.Treeview(
            grid_frame,
            columns=list(headings),
            show="headings",
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set,
        )
        y_scroll.configure(command=self._tree.yview)
        x_scroll.configure(command=self._tree.xview)

        for w_px, h in zip(col_px, headings, strict=True):
            self._tree.column(h, width=w_px, minwidth=60, anchor="w")
            self._tree.heading(h, text=h, anchor="w")

        self._tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        grid_frame.rowconfigure(0, weight=1)
        grid_frame.columnconfigure(0, weight=1)

        self._tree.bind("<Button-3>", self._on_tree_right_click)
        self._tree.bind("<Double-1>", self._on_row_double_click)

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

    def _grid_values_from_row(self, row: dict[str, str]) -> tuple[str, ...]:
        """内部row形式・UI保存形式のどちらでも Treeview 表示値へ変換する。"""
        cols = list(self._tree["columns"])
        # JSON保存後の UI 行（日本語列名）
        if any((row.get(c) or "") != "" for c in cols):
            return tuple((row.get(c) or "") for c in cols)
        # 解析直後の内部 row 形式
        return row_display_values(row)

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

        # ユーザ判断列: 右クリックで 〇/△/✖ を選択
        if cols[ci] == self.USER_JUDGMENT_COL:
            menu = tk.Menu(self, tearoff=0)
            opts = (
                ("〇", "〇"),
                ("△", "△"),
                ("✖", "×"),
            )
            for label, inner in opts:
                menu.add_command(
                    label=label,
                    command=lambda v=inner: self._set_user_judgment_cell(rid, v),
                )
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
            return

        # 社員番号列: 右クリックで入力（空欄許可）
        if cols[ci] == self.EMPLOYEE_NO_COL:
            self._prompt_edit_employee_no(rid)
            return

        # 合計勤務時間（読取）列: 右クリックで入力し、10進列も更新
        if cols[ci] == self.TOTAL_HOURS_RAW_COL:
            self._prompt_edit_total_hours_raw(rid)
            return

    def _set_user_judgment_cell(self, rid: str, value: str) -> None:
        if value not in ("〇", "△", "×"):
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
        self._status_var.set(f"読み込みました: {self._loaded_json_path}")
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
            self.after(
                0,
                lambda m=message: self._status_var.set(m[:800]),
            )

        base_done = 0
        if base_rows:
            base_done = len([r for r in base_rows if isinstance(r, dict)])
            self._progress_var.set(f"実行済 {base_done} / 対象 ?")

        def on_progress(done: int, total: int) -> None:
            # 継続解析時は読み込み済み件数を加算して表示
            self.after(
                0,
                lambda d=done, t=total, b=base_done: self._progress_var.set(
                    f"実行済 {b + d} / 対象 {b + t}"
                ),
            )

        def on_row(r: dict[str, str]) -> None:
            def append_row(row: dict[str, str]) -> None:
                vals = row_display_values(row)
                iid = self._tree.insert("", tk.END, values=vals)
                self._item_paths[iid] = row.get("resolved_path", "")
                try:
                    self._tree.yview_moveto(1)
                except tk.TclError:
                    pass

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
                out["user_judgment_company"] = uj
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
                # user_judgment を引継ぎ（既存があれば上書きしない）
                fn = (new_row.get("file_name") or "").strip()
                if fn and fn in base_map:
                    uj = (base_map[fn].get("user_judgment_company") or "").strip()
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
                if err is not None:
                    # 要件: 送信エラー等でも中断状態・保存/継続解析を有効にする
                    messagebox.showerror(
                        "解析エラー",
                        str(err),
                    )
                    self._status_var.set(
                        f"エラーで停止しました: {n_rows} 件を保持（JSON保存で続きから再開できます）"
                    )
                    # 進捗は固定せず、最後に見えていた値を維持（空にはしない）
                    return

                if cancelled:
                    self._status_var.set(
                        f"中断しました: {n_rows} 件を保持（JSON保存で続きから再開できます）"
                    )
                else:
                    self._status_var.set(
                        f"完了: {n_rows} 件をグリッド表示（解析結果.md を出力）"
                    )

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
