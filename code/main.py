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
    TARGET_ASSISTANT_NAME,
    create_client,
    row_display_values,
    run_analysis,
    summary_header_cells,
)


class KintaiApp(tk.Tk):
    USER_JUDGMENT_COL = "ユーザ判断"
    EMPLOYEE_NO_COL = "社員番号"

    def __init__(
        self,
        *,
        client,
        assistant_uid: str,
    ) -> None:
        super().__init__()
        self.title("勤務表解析")
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._client = client
        self._assistant_uid = assistant_uid
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
        col_px_base = [320, 200, 120, 120, 140, 80, 220, 72]
        if len(col_px_base) < len(headings):
            col_px = col_px_base + [default_w_px] * (len(headings) - len(col_px_base))
        else:
            col_px = col_px_base[: len(headings)]

        total_min = sum(col_px) + 28
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

        self.minsize(total_min, 520)
        self.geometry(f"{total_min}x680")

    def _browse_folder(self) -> None:
        d = filedialog.askdirectory(title="データを読み込むフォルダを選択")
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

    def _user_judgment_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.USER_JUDGMENT_COL)

    def _employee_no_column_index(self) -> int:
        return list(self._tree["columns"]).index(self.EMPLOYEE_NO_COL)

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
                "社員番号（7桁）を入力してください。\n"
                "- 空欄: 未設定\n"
                "- 7桁数字: 社員番号\n"
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

    def _save_json(self) -> None:
        if self._busy:
            return
        rows = self._current_grid_rows()
        if not rows:
            messagebox.showinfo("保存", "保存する行がありません。")
            return
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
        )
        if not fp:
            return
        out = Path(fp)
        self._write_json_file(out, rows)
        self._loaded_json_path = out
        self._loaded_rows = rows
        self._last_saved_snapshot = self._compute_snapshot(rows)
        self._status_var.set(f"保存しました: {self._loaded_json_path}")
        self._update_title()

    def _load_json(self) -> None:
        if self._busy:
            return
        fp = filedialog.askopenfilename(
            title="結果JSONを読み込み",
            filetypes=[("JSON", "*.json"), ("すべて", "*.*")],
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
        cols = list(self._tree["columns"])
        for r in rows:
            vals = [r.get(c, "") for c in cols]
            iid = self._tree.insert("", tk.END, values=tuple(vals))
            self._item_paths[iid] = (r.get("resolved_path") or "").strip()

    def _update_title(self) -> None:
        base = "勤務表解析"
        if self._loaded_json_path is None:
            self.title(base)
        else:
            self.title(f"{base} - {self._loaded_json_path.name}")

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
                    self._auto_save_on_exit()
                except Exception as e:
                    messagebox.showerror("保存失敗", str(e))
                    return
        self.destroy()

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
            # kintai_core 側の row_display_values は key ベース。必要なものだけ入れる。
            # 継続解析で user_judgment_company を引き継ぐため、UIの「ユーザ判断」を反映。
            out: dict[str, str] = {}
            out["file_name"] = (r.get("画像ファイル名") or r.get("file_name") or "").strip()
            out["resolved_path"] = (r.get("resolved_path") or "").strip()
            uj = (r.get(self.USER_JUDGMENT_COL) or r.get("user_judgment_company") or "").strip()
            if uj:
                out["user_judgment_company"] = uj
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


def main() -> None:
    root = tk.Tk()
    root.withdraw()

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
    assistants = client.get_assistants() or []
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
    dlg.title("アシスタント選択")
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

    btns = ttk.Frame(dlg)
    btns.pack(fill=tk.X, padx=12, pady=12)

    chosen: dict[str, str | None] = {"name": None}

    def on_ok() -> None:
        chosen["name"] = (sel_var.get() or "").strip() or default_name
        dlg.destroy()

    def on_cancel() -> None:
        chosen["name"] = None
        dlg.destroy()

    ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT)
    ttk.Button(btns, text="キャンセル", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 8))

    dlg.bind("<Return>", lambda _e: on_ok())
    dlg.bind("<Escape>", lambda _e: on_cancel())
    cb.focus_set()

    root.wait_window(dlg)

    if not chosen["name"]:
        root.destroy()
        sys.exit(0)

    selected_name = str(chosen["name"])
    selected = next((a for a in assistants if (a.get("name") == selected_name)), None)
    assistant_uid = str(selected["uid"]) if selected and selected.get("uid") is not None else ""

    if not assistant_uid:
        messagebox.showerror(
            "エラー",
            f"アシスタント「{selected_name}」が見つかりません。",
            parent=root,
        )
        root.destroy()
        sys.exit(1)

    root.destroy()

    app = KintaiApp(client=client, assistant_uid=assistant_uid)
    app.mainloop()


if __name__ == "__main__":
    main()
