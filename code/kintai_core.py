"""勤務表画像・PDF解析（NewtonX）のコア処理。"""
from __future__ import annotations

from collections.abc import Callable

from newtonx_adk import NewtonXClient, ConfigManager, FileUploadError
from openpyxl import load_workbook
import json
import threading
import math
import re
import tempfile
import time
import types
import unicodedata
from datetime import datetime as DTDateTime, time as DTTime, timedelta as DTTimeDelta
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from PIL import Image, ImageOps, UnidentifiedImageError

# バイナリ換算の 1MiB。「未満」なのでこれ未満のサイズになるよう調整する
MIB = 1024 * 1024

IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff"}
PDF_SUFFIX = ".pdf"
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
TARGET_EXCEL_SHEET_NAME = "タイムシート兼作業報告書_お客様先用"

# アップロード: 初回試行後、最大 UPLOAD_MAX_RETRIES 回まで再試行（合計で最大 1 + UPLOAD_MAX_RETRIES 回）
UPLOAD_MAX_RETRIES = 3
UPLOAD_RETRY_DELAY_SEC = 0.35

DEFAULT_PARALLEL_ANALYSIS_CHATS = 2
PARALLEL_WORKERS_MAX = 32

# グリッド・集計表の先頭列（〇: リトライ上限内でアップロード成功 / ✖: それ以外）
SUMMARY_UPLOAD_COL = "アップロード"
SUMMARY_TARGET_SHEET_COL = "対象シート有無"

TARGET_ASSISTANT_NAME = "GPT-5.2(高性能)"
#TARGET_ASSISTANT_NAME = "GPT-5.4-mini(高速)"

def _process_sse_response_no_print(self, response) -> str:
    full_response = ""
    for line in response.iter_lines(decode_unicode=True):
        if line and line.startswith("data: "):
            data = line[6:]
            if data == "[DONE]":
                break
            try:
                parsed_data = json.loads(data)
                if parsed_data.get("type") == "content":
                    full_response += parsed_data.get("content", "")
            except json.JSONDecodeError:
                full_response += data
    return full_response


def create_client() -> NewtonXClient:
    """NewtonX クライアントを生成し、SSE の標準出力を抑制するパッチを当てる。"""
    config_manager = ConfigManager()
    _cfg = config_manager.get_config()
    _cfg.timeout = max(getattr(_cfg, "timeout", 30) or 30, 300)
    client = NewtonXClient(config_manager)
    client._process_sse_response = types.MethodType(
        _process_sse_response_no_print, client
    )
    return client


def resolve_assistant_uid(
    client: NewtonXClient, name: str = TARGET_ASSISTANT_NAME
) -> str | None:
    assistants = client.get_assistants()
    selected = next((a for a in assistants if a.get("name") == name), None)
    return str(selected["uid"]) if selected else None


def _compress_image_under_1mib(src: Path) -> Path:
    """元ファイルは変更しない。1MiB未満のJPEGを tempfile に書き、そのパスを返す。"""
    img = Image.open(src)
    img.load()
    img = ImageOps.exif_transpose(img)

    if getattr(img, "is_animated", False):
        img.seek(0)

    if img.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        img = bg
    elif img.mode == "P":
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[-1])
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")

    tmp = tempfile.NamedTemporaryFile(prefix="nx_upload_", suffix=".jpg", delete=False)
    tmp.close()
    out_path = Path(tmp.name)

    current = img
    for _ in range(20):
        for quality in range(92, 9, -4):
            current.save(out_path, format="JPEG", quality=quality, optimize=True)
            if out_path.stat().st_size < MIB:
                return out_path
        w, h = current.size
        if w <= 320 and h <= 320:
            break
        current = current.resize(
            (max(1, int(w * 0.88)), max(1, int(h * 0.88))),
            Image.Resampling.LANCZOS,
        )

    raise RuntimeError(f"1MiB未満に収められませんでした: {src}")


def _resolve_upload_path(original: Path) -> tuple[str, Path | None]:
    """アップロード用パスと、削除すべき一時ファイル（あれば）を返す。元ファイルは書き換えない。"""
    if original.stat().st_size < MIB:
        return str(original), None
    tmp = _compress_image_under_1mib(original)
    return str(tmp), tmp


def _build_check_message(display_file_name: str) -> str:
    """解析結果に出すファイル名をローカルの実名に固定する。"""
    return f"""
勤務表の画像を解析し、１画像１行で以下の内容にあたるものをmd形式で表で出力してください。
画像ファイル名の会社名は会社名として出力しないこと。
会社名とは、画像上に法人の会社名として認識できた名前のことである。会社名は契約先、常駐先、会社名という文字の近辺にある場合が多い。
会社名が複数存在する場合は、会社名：会社名2,会社名2とカンマで区切って出力してください。
会社名でセラクという名称を含むものは、株式会社セラクなどは会社名から除外してください。
画像ファイル名（アップロードファイル名）として、次の名前のみを記載してください（サーバー側のIDや別名は使わないこと）:
{display_file_name}

会社名：
氏名：
合計勤務時間：（8:20・8時間20分・8.20・101_10H・86.17H 等。必要に応じて実データの表記のまま）
押印有無：〇/✖

"""


def _build_pdf_check_message(display_file_name: str) -> str:
    """PDF用。アップロード直後は send_message にドキュメントIDを渡さなくても参照できる想定。"""
    return f"""
同じPDFに勤怠（出退勤・打刻・勤務時間等）の情報と経費精算（領収書・立替等）の情報の両方が含まれている場合は、経費精算は無視し、必ず勤怠（勤務表）の情報だけを根拠に回答してください。経費側の社名・氏名・金額は採用しないでください。
PDFファイル名の会社名は会社名として抽出しないこと。
勤務表（勤怠）のPDFを解析し、１ファイルあたり適切な行数で以下の内容にあたるものをmd形式で表で出力してください。出力する勤務先・氏名・合計勤務時間・押印はすべて勤怠部分の記載に基づきます。
会社名とは、ファイル上に法人の会社名として認識できた名前のことである。会社名は契約先、常駐先、会社名という文字の近辺にある場合が多い。
会社名が複数存在する場合は、会社名：会社名2,会社名2とカンマで区切って出力してください。
会社名でセラクという名称を含むものは、株式会社セラクなどは会社名から除外してください。
PDFファイル名（アップロードファイル名）として、次の名前のみを記載してください（サーバー側のIDや別名は使わないこと）:
{display_file_name}

会社名：
氏名：
合計勤務時間：（8:20・8時間20分・8.20・101_10H・86.17H 等。必要に応じて実データの表記のまま）
押印有無：〇/✖

"""


def _excel_target_sheet_symbol(
    file_path: Path,
    *,
    target_sheet_name: str = TARGET_EXCEL_SHEET_NAME,
) -> tuple[str, str]:
    """Excel ブック内に対象シートがあるかを判定し、(記号, 補足メッセージ) を返す。"""
    try:
        with zipfile.ZipFile(file_path) as zf:
            with zf.open("xl/workbook.xml") as fp:
                root = ET.parse(fp).getroot()
        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheet_names = [
            str(node.attrib.get("name") or "").strip()
            for node in root.findall(".//x:sheets/x:sheet", ns)
        ]
    except (FileNotFoundError, OSError, KeyError, ET.ParseError, zipfile.BadZipFile) as e:
        return "✖", f"Excelシート確認失敗: {e}"
    return ("〇", "") if target_sheet_name in sheet_names else ("✖", "")


def _excel_cell_value_to_raw_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, DTDateTime):
        return value.strftime("%H:%M")
    if isinstance(value, DTTime):
        return value.strftime("%H:%M")
    if isinstance(value, DTTimeDelta):
        total_minutes = max(0, int(round(value.total_seconds() / 60.0)))
        hh, mm = divmod(total_minutes, 60)
        return f"{hh}:{mm:02d}"
    return str(value).strip()


def _excel_cell_value_to_decimal_hours(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, DTDateTime):
        hours = (
            value.hour
            + value.minute / 60.0
            + value.second / 3600.0
            + value.microsecond / 3600000000.0
        )
        return _format_decimal_str(hours)
    if isinstance(value, DTTime):
        hours = (
            value.hour
            + value.minute / 60.0
            + value.second / 3600.0
            + value.microsecond / 3600000000.0
        )
        return _format_decimal_str(hours)
    if isinstance(value, DTTimeDelta):
        return _format_decimal_str(value.total_seconds() / 3600.0)
    if isinstance(value, (int, float)):
        return _format_decimal_str(float(value) * 24.0)
    return _work_hours_string_to_decimal(str(value))


def _employee_no_from_file_name(file_name: str) -> str:
    stem = Path(file_name or "").stem
    if len(stem) < 7:
        return ""
    tail7 = stem[-7:]
    if tail7.isdigit():
        return tail7
    if re.fullmatch(r"BP\d{5}", tail7, re.IGNORECASE):
        return tail7.upper()
    return "社員番号エラー"


def _extract_excel_target_sheet_row(
    file_path: Path,
    *,
    target_sheet_name: str = TARGET_EXCEL_SHEET_NAME,
) -> dict[str, str]:
    row = {
        "upload_ok": "",
        "file_name": file_path.name,
        "resolved_path": str(file_path.resolve()),
        "employee_no": _employee_no_from_file_name(file_path.name),
        "target_sheet_exists": "✖",
        "analysis": "",
        "source_kind": "excel",
    }

    sheet_symbol, note = _excel_target_sheet_symbol(
        file_path,
        target_sheet_name=target_sheet_name,
    )
    row["target_sheet_exists"] = sheet_symbol
    if note:
        row["analysis"] = note
        return row
    if sheet_symbol != "〇":
        return row

    wb = None
    try:
        wb = load_workbook(file_path, data_only=True, read_only=True)
        ws = wb[target_sheet_name]
        company = (ws["G5"].value or "")
        person = (ws["J7"].value or "")
        total_raw_value = ws["F42"].value

        file_company, _file_person = _parse_filename_company_and_person(file_path.name)
        doc_company_raw = str(company).strip() if company is not None else ""
        doc_company_match = _document_company_for_match(doc_company_raw)
        match_company = _match_company_symbol_single(file_company, doc_company_match)

        row["name_company_from_file"] = file_company
        row["name_company_from_doc"] = doc_company_raw
        row["name_company_1"] = str(company).strip() if company is not None else ""
        row["name_person_from_doc"] = str(person).strip() if person is not None else ""
        row["total_hours_raw"] = _excel_cell_value_to_raw_text(total_raw_value)
        row["total_hours_decimal"] = _excel_cell_value_to_decimal_hours(total_raw_value)
        row["match_company"] = match_company
        row["user_judgment_company"] = match_company
    except Exception as e:
        row["analysis"] = f"Excelセル読み取り失敗: {e}"
    finally:
        if wb is not None:
            wb.close()
    return row


# --- 合計勤務時間の10進化 -----------------------------------------------------------------
def _trunc2_down(x: float) -> float:
    """正の数について、小数点以下第3位以下を切り捨て、第2位まで有効化する（100倍floor/100）。"""
    n = 100
    return math.floor((x if x >= 0 else 0) * n + 1e-9) / n


def _format_decimal_str(value: float) -> str:
    """小数点以下第3位以下切り捨てのうえ、小数第2位まで 2 桁表示で保持。"""
    t = _trunc2_down(value)
    return f"{t:.2f}"


def _work_hours_string_to_decimal(raw: str) -> str:
    """
    合計（総）勤務時間の10進2桁化（第3位未満切り捨て）。
    - 8.35  … 十進（小数点はそのまま時間。8.35 時間 = 8.35）
    - 8.35H … 十進（N.NH は N.N 時間）
    - 8:35  … 60進 → H+MM/60
    - 8時間35分 … 60進
    - 8_35H  … 60進（時_分H）
    解釈できなければ空文字。
    """
    if not (raw and str(raw).strip()):
        return ""
    t = unicodedata.normalize("NFKC", str(raw).strip())
    t = t.split("(")[0].split("（")[0]
    t = re.sub(r"[\s　]+", "", t)
    t = t.replace("：", ":").replace("．", ".")
    if not t:
        return ""

    # 0) 60進 H_MMH（例: 101_10H → 101時間10分。末尾は H/h）
    m = re.match(r"^(\d+)_(\d{1,2})[Hh]$", t)
    if m:
        try:
            hh, mm = int(m.group(1)), int(m.group(2))
            if mm > 59:
                return ""
            return _format_decimal_str(float(hh) + mm / 60.0)
        except (ValueError, OverflowError):
            return ""

    # 1) HH:MM
    m = re.match(r"^(\d+):(\d{1,2})$", t)
    if m:
        try:
            hh, mm = int(m.group(1)), int(m.group(2))
            if mm > 59:
                return ""
            return _format_decimal_str(float(hh) + mm / 60.0)
        except (ValueError, OverflowError):
            return ""

    # 2) HH時間MM分
    m = re.match(r"^(\d+)時間(\d{1,2})分$", t)
    if m:
        try:
            hh, mm = int(m.group(1)), int(m.group(2))
            if mm > 59:
                return ""
            return _format_decimal_str(float(hh) + mm / 60.0)
        except (ValueError, OverflowError):
            return ""

    # 2.5) HH時間（分の記載なしは 0分 扱い）
    m = re.match(r"^(\d+)時間$", t)
    if m:
        try:
            hh = int(m.group(1))
            return _format_decimal_str(float(hh))
        except (ValueError, OverflowError):
            return ""

    # 3) H.xxxH（例: 174.75H = 174.75 時間の10進・末尾 H）
    m = re.match(r"^(\d+\.\d+)[Hh]$", t)
    if m:
        try:
            return _format_decimal_str(float(m.group(1)))
        except (ValueError, OverflowError):
            return ""

    # 4) H.xxx ドット（8.20 は 8.20h 風。小数部は桁数可変→浮動後に2位切り捨て）
    m = re.match(r"^(\d+)\.(\d+)$", t)
    if m:
        try:
            h, fr = m.group(1), m.group(2)
            val = int(h) + int(fr) / 10.0 ** len(fr)
            return _format_decimal_str(val)
        except (ValueError, OverflowError):
            return ""

    return ""


def _md_table_row_cells(line: str) -> list[str]:
    s = (line or "").strip()
    if not s.startswith("|"):
        return []
    inner = s[1:-1] if s.endswith("|") and len(s) > 1 else s[1:]
    return [c.strip() for c in inner.split("|")]


def _is_md_table_separator_row(cells: list[str]) -> bool:
    if not cells:
        return True
    for c in cells:
        t = c.replace(" ", "").replace("\u3000", "")
        if re.search(r"[^:｜\-\s─━]", t):
            return False
    return True


# 区切り行の隣接セルに誤ってマッチしないよう、列名そのものを弾く
_KINTAI_HEADER_CELLS: frozenset[str] = frozenset(
    {
        "会社名",
        "勤務先",
        "氏名",
        "合計勤務時間",
        "合計",
        "押印有無",
        "押印",
        "有無",
        "画像ファイル名",
        "ファイル名",
    }
)


def _is_table_headerish_cell(s: str) -> bool:
    t = unicodedata.normalize("NFKC", (s or "").strip())
    if t in _KINTAI_HEADER_CELLS or t in ("画像", "合計勤務", "労働時間"):
        return True
    if re.fullmatch(r"氏名[（(].+?[)）]?", t):
        return True
    return bool(re.match(r"^(合計|押印|勤務先|画像|ファイル)", t) and len(t) <= 20)


def _kintai_header_col_roles(cells: list[str]) -> dict[str, int] | None:
    """
    見出し行: 各列の役 (co, pe, th, se, fn) → 0始まりインデックス
    勤怠表らしい列が2つ以上特定できたときだけ返す。
    """
    roles: dict[str, int] = {}
    for j, c in enumerate(cells):
        t = unicodedata.normalize("NFKC", (c or "").strip())
        if not t:
            continue
        if "co" not in roles and (
            t in ("会社名", "勤務先", "勤務先会社", "就業先")
            or ("勤務" in t and "会社" in t)
        ):
            roles["co"] = j
        elif "pe" not in roles and (t == "氏名" or t.startswith("氏名")):
            roles["pe"] = j
        elif "th" not in roles and (
            t in ("合計勤務時間", "合計勤務", "総労働時間", "労働時間", "合計", "労働")
            or (("合計" in t or "総" in t) and ("勤務" in t or "時間" in t or "労働" in t))
        ):
            roles["th"] = j
        elif "se" not in roles and (t in ("押印有無",) or ("押印" in t) or t == "有無"):
            roles["se"] = j
        elif "fn" not in roles and (("画像" in t and "ファイル" in t) or t in ("ファイル名",) or t == "画像ファイル名"):
            roles["fn"] = j
    if "co" in roles and "pe" in roles:
        return roles
    if "pe" in roles and "th" in roles:
        return roles
    if "co" in roles and "th" in roles:
        return roles
    if "fn" in roles and (len(roles) >= 2):
        return roles
    if len(roles) >= 3:
        return roles
    return None


def _extract_kintai_from_markdown_table(
    text: str, expected_file_name: str
) -> dict[str, str] | None:
    """
    LLMの |...| 表から、見出し行の次以降のデータ行のセルを取る（隣接セルに列名が入る行はスキップ）。
    """
    if not (text and str(text).strip()):
        return None
    exp = Path(expected_file_name or "").name

    table_lines: list[list[str]] = []
    for _i, line in enumerate(str(text).splitlines()):
        s = line.strip()
        if not s.startswith("|"):
            continue
        cells = _md_table_row_cells(s)
        if not cells:
            continue
        if _is_md_table_separator_row(cells):
            continue
        table_lines.append(cells)

    def _row_looks_header_only(c: list[str]) -> bool:
        nonempty = [x for x in c if (x or "").strip()]
        if not nonempty:
            return True
        return all(_is_table_headerish_cell(x) for x in nonempty)

    def _cell_get(cells: list[str], key: str, roles: dict[str, int]) -> str:
        if key not in roles:
            return ""
        j = roles[key]
        if j < 0 or j >= len(cells):
            return ""
        v = (cells[j] or "").strip()
        if not v or _is_table_headerish_cell(v):
            return ""
        return v

    def _build_row(dcells: list[str], roles: dict[str, int]) -> dict[str, str]:
        return {
            "company": _cell_get(dcells, "co", roles),
            "person": _cell_get(dcells, "pe", roles),
            "total": _cell_get(dcells, "th", roles),
            "seal": _cell_get(dcells, "se", roles),
        }

    def _row_matches_filename(
        dcells: list[str], roles: dict[str, int]
    ) -> bool:
        if "fn" not in roles or not exp:
            return True
        vfn = _cell_get(dcells, "fn", roles)
        if not vfn:
            return True
        vfnn = unicodedata.normalize("NFKC", vfn)
        st = Path(exp).stem
        stn = unicodedata.normalize("NFKC", st)
        return not (
            st not in vfnn
            and stn not in vfnn
            and exp not in vfnn
            and vfnn not in exp
        )

    for require_fn in (True, False):
        for hi, hcells in enumerate(table_lines):
            roles = _kintai_header_col_roles(hcells)
            if not roles:
                continue
            for sub in range(hi + 1, len(table_lines)):
                dcells = table_lines[sub]
                if _kintai_header_col_roles(dcells) is not None:
                    break
                if not dcells or _row_looks_header_only(dcells):
                    continue
                if require_fn and not _row_matches_filename(dcells, roles):
                    continue
                d = _build_row(dcells, roles)
                if any(d.values()):
                    return d
    return None


def _is_plausible_total_labor_value(v: str) -> bool:
    """表のヘッダ行の隣セル（列名等）を拾わない: 数値必須。"""
    t = (v or "").strip()
    if not t or len(t) > 120:
        return False
    if t in ("-", "—", "未記載", "n/a", "N/A", "null", "要確認", "（なし）"):
        return False
    if _is_table_headerish_cell(t) and not re.search(r"\d", t):
        return False
    return bool(re.search(r"\d", t))


def _extract_total_work_hours_from_document(text: str) -> str:
    """解析文から 合計勤務時間 ラベルに続く1トークンを抜出（文面上の先頭候補で数値含むもののみ）。"""
    if not (text and str(text).strip()):
        return ""
    raw = str(text)
    cands: list[tuple[int, str]] = []
    _skip = ("", "-", "—", "未記載", "n/a", "N/A", "null", "要確認")
    for m in re.finditer(
        r"合計勤務時間\s*[：:｜\|]?\s*([^\n\r\|]+?)(?=(\s*\||$|\n))",
        raw,
    ):
        v = m.group(1).strip()
        if v and v not in _skip:
            cands.append((m.start(1), v))
    for m in re.finditer(
        r"\|[^\n]*?合計勤務時間(?:\s*\|)?\s*([^\n\|]+)", raw, re.IGNORECASE
    ):
        v = m.group(1).strip()
        if v and v not in _skip:
            cands.append((m.start(1), v))
    cands.sort(key=lambda x: (x[0], x[1]))
    for _pos, v in cands:
        if _is_plausible_total_labor_value(v):
            return v
    return ""


def _decimal_for_table_display(s: str) -> str:
    """10進列: 変換結果を表示（不要な末尾 .00 は落とし、8.35 / 98.5 のようになる）。"""
    t = (s or "").strip()
    if not t:
        return "（なし）"
    try:
        x = float(t)
    except ValueError:
        return t
    a = f"{x:.2f}"
    return a.rstrip("0").rstrip(".")


# --- ファイル名と解析文字列の 会社名・氏名 照合（〇／△／✖）----------------------------------------
# 法人格表記（文書内比較前に反復除去）
_LEGAL_TOKENS = (
    "株式会社", "（株）", "㈱", "有限会社", "（有）", "合名会社", "合資会社", "合同会社",
    "一般社団法人", "一般財団法人", "ＮＰＯ法人", "NPO法人", "医療法人", "学校法人", "宗教法人",
    "独立行政法人", "(株)", "（有)", "(有)",
)
_LEGAL_TOKENS_SORTED = tuple(sorted(_LEGAL_TOKENS, key=len, reverse=True))

# 系列・グループ名寄せ用キーワード
_GROUP_MARKERS = (
    "ホールディング", "ホールディングス", "グループ",
)


def _strip_trailing_sama(company_segment: str) -> str:
    """末尾の「様」を除去。ファイル名の会社部分に加え、照合用に読み取り側の会社名にも使う。"""
    t = unicodedata.normalize("NFKC", (company_segment or "").strip())
    if t.endswith("様"):
        t = t[: -1].rstrip(" 　\t")
    return t


def _parse_filename_company_and_person(file_name: str) -> tuple[str, str]:
    """ファイル名（拡張子除く）から会社名と氏名を分離。

    先頭 […] を繰り返し除去し、会社側の末尾「様」を除去。

    例:
      - [xxx]AAA様_山田太郎_123456.jpg -> (AAA, 山田太郎)
      - AAA様 山田太郎 123456.png -> (AAA, 山田太郎)

    ※社員番号は別ロジックで抽出して表示に付与する。
    """
    stem = Path(file_name).stem
    s = stem
    while True:
        m = re.match(r"^\[([^\]]+)\]\s*", s)
        if m:
            s = s[m.end() :]
        else:
            break
    s = s.strip()

    # 後ろに社員番号等が付いていても、会社名/氏名の分割は「最初の区切り」で行う
    for sep in ("_", "＿", "·", "・", " ", "　"):
        if sep in s:
            a, b = s.split(sep, 1)
            return _strip_trailing_sama(a), b.strip()
    return _strip_trailing_sama(s), ""


def _strip_leading_legal_tokens(t: str, tokens: tuple[str, ...]) -> str:
    """先頭に連続する法人格・組織形態を長い表記から順に剥がす。"""
    t = t.lstrip(" 　、，・.")
    while t:
        stripped = False
        for w in tokens:
            if t.startswith(w):
                t = t[len(w) :].lstrip(" 　、，・.")
                stripped = True
                break
        if not stripped:
            break
    return t


def _strip_legal_forms(company: str) -> str:
    """法人格・組織形態を先頭・文中・末尾のいずれでも除去（反復。長い表記を優先）。"""
    t = unicodedata.normalize("NFKC", str(company).strip())
    for _ in range(14):
        t0 = t
        t = _strip_leading_legal_tokens(t, _LEGAL_TOKENS_SORTED)
        for w in _LEGAL_TOKENS_SORTED:
            t = t.replace(w, "")
        t = re.sub(r"[\s　、，・\.]+", "", t)
        if t == t0:
            break
    return t


def _normalize_mixed(s: str) -> str:
    """数値・英字: 全半角・大文字小文字のゆらぎを吸収。"""
    t = unicodedata.normalize("NFKC", s or "")
    out: list[str] = []
    for c in t:
        if c.isascii() and c.isalpha():
            out.append(c.lower())
        else:
            out.append(c)
    return "".join(out)


def _normalize_person(s: str) -> str:
    return _normalize_mixed(re.sub(r"[\s\u3000\u00a0]+", "", s or ""))


def _company_core_for_match(s: str) -> str:
    """拠点語尾を外したコア比較用文字列。様→法人格除去の順（読取・ファイル名とも）。"""
    t = _strip_trailing_sama(s)
    t = _strip_legal_forms(t)
    t = _normalize_mixed(t)
    t = re.sub(
        r"(本社|支社|支店|営業所|工場|事業所|オフィス|店|エリア|地区)$", "", t
    )
    return t


def _prefer_kintai_text_for_extraction(text: str) -> str:
    """
    解析テキストに勤怠と経費精算の両方が混在する想定のとき、照合・勤務先・氏名の抽出元を勤怠側に寄せる。
    （LLMが両方出した場合のフォロー。見出し・先頭出現位置のヒューリスティクス。）
    """
    t = str(text) if text else ""
    if not t.strip():
        return t
    has_kin = any(
        k in t
        for k in (
            "勤怠",
            "勤務表",
            "出退勤",
            "出退",
            "合計勤務",
            "合計勤務時間",
            "打刻",
            "会社名",
            "押印",
        )
    )
    has_kei = "経費" in t or "精算" in t
    if not (has_kin and has_kei):
        return t
    keihi_pos: int | None = None
    for m in re.finditer(
        r"(?:^|\n)#{1,3}\s*[^\n]*?(?:経費精算|精算[（(]?経費|経費[（(]?精算|領収|立替|交通費?)",
        t,
        re.MULTILINE,
    ):
        p = m.start()
        if keihi_pos is None or p < keihi_pos:
            keihi_pos = p
    if keihi_pos is None:
        p0 = t.find("経費精算")
        if p0 >= 0:
            keihi_pos = p0
    if keihi_pos is None:
        return t
    head = t[:keihi_pos]
    if any(
        k in head
        for k in (
            "勤怠",
            "勤務表",
            "出退",
            "合計勤務",
            "合計勤務時間",
            "打刻",
            "勤務先",
        )
    ):
        return head.rstrip()
    # 文書上で経費が先・勤怠が後の可能性
    tail = t[keihi_pos + 1 :]
    for kw in ("勤怠", "勤務表", "出退勤", "出退", "合計勤務", "合計勤務時間", "会社名"):
        i = tail.find(kw)
        if i >= 0:
            return tail[i:].strip()
    return t


def _katakana_latin_mismatch_both_look_like_transcription(a: str, b: str) -> bool:
    """片方が主にカタカナ、他方が主にラテン字なら 読み揺れ △ 候補とみなす。"""
    def mostly_kat(x: str) -> bool:
        u = re.sub(r"[\s\w\-+]", "", x)
        if len(u) < 2:
            return False
        return not re.search(r"[^\u30a0-\u30ff\u30f4]", u)

    def mostly_lat(x: str) -> bool:
        return bool(re.search(r"[A-Za-z]{2,}", _normalize_mixed(x)))

    return (mostly_kat(a) and mostly_lat(b)) or (mostly_kat(b) and mostly_lat(a))


def _company_text_contains_seraku(s: str) -> bool:
    """社名文字列にセラクまたは英字 seraku（大小無視）が含まれるか。"""
    if not (s or "").strip():
        return False
    t = re.sub(r"\s+", "", unicodedata.normalize("NFKC", s))
    return bool(re.search(r"セラク|seraku", t, re.IGNORECASE))


def _file_co_for_match(company_from_file: str) -> str:
    """ファイル名由来の会社。セラク/seraku を含むときは照合対象外（空文字）。"""
    s = (company_from_file or "").strip()
    if not s or _company_text_contains_seraku(s):
        return ""
    return s


def _document_company_for_display(ktab_company: str) -> str:
    """会社名1列用。セラクを含む勤怠表の会社は表示しない（不明扱い）。"""
    k = (ktab_company or "").strip()
    if not k:
        return "不明"
    if _company_text_contains_seraku(k):
        return "不明"
    return k


def _document_company_for_match(ktab_company: str) -> str:
    """照合用の文書側会社。複数候補はカンマ区切りのまま返し、比較側で分解する。"""
    return (ktab_company or "").strip()


def _split_document_company_candidates(companies_text: str) -> list[str]:
    """会社名の複数候補を分解する。カンマ区切り想定だが、日本語読点等も緩く許容する。"""
    raw = unicodedata.normalize("NFKC", (companies_text or "").strip())
    if not raw:
        return []
    parts = re.split(r"\s*[,，、]\s*", raw)
    out: list[str] = []
    for p in parts:
        t = (p or "").strip()
        if not t:
            continue
        if _company_text_contains_seraku(t):
            continue
        out.append(t)
    return out


def _match_company_symbol_single(file_co: str, doc_company: str) -> str:
    """勤怠表から取れた会社名とファイル名会社を照合。複数社名は候補のいずれか一致で判定。"""
    fp = _file_co_for_match(file_co)
    doc_candidates = _split_document_company_candidates(doc_company)
    if not doc_candidates:
        return "〇" if not fp else "✖"
    if not fp:
        return "△"
    symbols = [_compare_company(fp, d) for d in doc_candidates]
    if "〇" in symbols:
        return "〇"
    if "△" in symbols:
        return "△"
    return "✖"


def _series_company_hint(f_raw: str, d_raw: str) -> bool:
    """系列・ホール等の表記違いを粗く △ 相当とする。"""
    c1, c2 = _company_core_for_match(f_raw), _company_core_for_match(d_raw)
    if c1 and c1 == c2:
        return False
    for g in _GROUP_MARKERS:
        f_has, d_has = g in f_raw, g in d_raw
        if f_has != d_has and (c1 in c2 or c2 in c1):
            return True
    return False



def _extract_person_name_from_document(text: str) -> str:
    """氏名を文から抽出。"""
    if not (text and str(text).strip()):
        return ""
    raw = str(text)
    cands: list[tuple[int, str]] = []
    for m in re.finditer(
        r"氏名\s*[：:｜\|]?\s*([^\n\r\|]+?)(?=(\s*\||$|\n))", raw
    ):
        v = m.group(1).strip()
        if v and v not in ("", "-", "—", "未記載") and not _is_table_headerish_cell(v):
            cands.append((m.start(1), v))
    for m in re.finditer(
        r"\|[^\n]*?氏名(?:\s*\|)?\s*([^\n\|]+)", raw, re.IGNORECASE
    ):
        v = m.group(1).strip()
        if v and not _is_table_headerish_cell(v):
            cands.append((m.start(1), v))
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands[0][1] if cands else ""


def _normalize_seal_phrase(s: str) -> str:
    """押印欄の短い文を 〇/✖ 等に寄せる。"""
    s = (s or "").strip()
    if not s:
        return ""
    # 見出しセル（「押印有無」等）を弾く。ただし記号自体は許可。
    if _is_table_headerish_cell(s) and s not in ("有", "無", "有無", "〇", "✖", "○", "×", "✕"):
        return ""

    t = unicodedata.normalize("NFKC", s).replace(" ", "")

    # 否定系 → ✖
    if re.search(r"^(無|不|否|未|な(い|し)?|印な|印の記載な|－|—|N/?A)", t):
        return "✖"
    # 肯定系 → 〇
    if re.search(r"^(有|是|印あり?|有印|○|〇)", t):
        return "〇"

    # 文中に含まれるケース
    if "無" in t and "有" not in t:
        return "✖"
    if "有" in t and "無" not in t:
        return "〇"

    # 既に記号のとき（× は NFKC で × になる）
    if t in ("×", "✕", "✖"):
        return "✖"
    if t in ("○", "〇"):
        return "〇"

    return s[:10]


def _extract_seal_in_from_document(text: str) -> str:
    """押印有無: 有 / 無 等を文から抜出（表記は短く揃える）。"""
    if not (text and str(text).strip()):
        return ""
    raw = str(text)
    cands: list[str] = []
    for m in re.finditer(
        r"押印有無\s*[：:｜\|]?\s*([^\n\r\|]+?)(?=(\s*\||$|\n))", raw
    ):
        v = m.group(1).strip()
        if v and v not in ("", "-", "—", "未記載", "N/A", "n/a", "null", "要確認"):
            if not _is_table_headerish_cell(v) or v in ("有", "無"):
                cands.append(v)
    for m in re.finditer(
        r"\|[^\n]*?押印有無(?:\s*\|)?\s*([^\n\|]+)", raw, re.IGNORECASE
    ):
        v = m.group(1).strip()
        if v and v not in ("", "-", "—") and (
            not _is_table_headerish_cell(v) or v in ("有", "無")
        ):
            cands.append(v)
    s = (cands[0] if cands else "").strip()
    if not s:
        return ""
    return _normalize_seal_phrase(s)


def _compare_person(filename_person: str, document_person: str) -> str:
    """氏名: 姓名間の空白は正規化して比較。一致→〇 部分一致→△ それ以外✖。"""
    f = (filename_person or "").strip()
    d = (document_person or "").strip()
    if not f or not d:
        return "✖"
    a, b = _normalize_person(f), _normalize_person(d)
    if a == b:
        return "〇"
    if a in b or b in a:
        return "△"
    return "✖"


def _compare_company(
    company_from_file: str,
    company_from_doc: str,
) -> str:
    """
    ファイル名の会社と文書の会社名を比較（_company_core_for_match で双方とも末尾 様・法人格を除去）。
    コアが一致→〇。
    コアが完全一致でなくても、一方が他方の部分文字列なら△（短い方2文字以上）。
    そのほか読み揺れ・系列風は△。それ以外✖。
    """
    f = (company_from_file or "").strip()
    d = (company_from_doc or "").strip()
    if not f or not d:
        return "✖"
    cfn = _company_core_for_match(f)
    cdoc = _company_core_for_match(d)
    if not cfn or not cdoc:
        return "✖"
    if cfn == cdoc:
        return "〇"
    shorter, longer = (cfn, cdoc) if len(cfn) <= len(cdoc) else (cdoc, cfn)
    if len(shorter) >= 2 and shorter in longer:
        return "△"
    if _katakana_latin_mismatch_both_look_like_transcription(cfn, cdoc):
        return "△"
    if _series_company_hint(f, d):
        return "△"
    return "✖"


def _enrich_with_match_scores(
    row: dict[str, str],
    analysis: str,
    *,
    prefer_kintai_section: bool = False,
) -> None:
    """比較記号（会社名・氏名）と参照文字列を row に追加。

    追加仕様:
      - 画像ファイル名の「拡張子の前にある7桁」を社員番号として扱い、氏名の右（別列）に表示する。
      - 末尾7文字が存在するのに数字7桁ではない場合、「社員番号エラー」と表示する。
    """

    text_for_extraction = (
        _prefer_kintai_text_for_extraction(analysis) if prefer_kintai_section else analysis
    )
    fn = row.get("file_name", "")
    fp_co, fp_pe = _parse_filename_company_and_person(fn)
    row["employee_no"] = _employee_no_from_file_name(fn)
    ktab = _extract_kintai_from_markdown_table(text_for_extraction, fn) or {}
    k_co = (ktab.get("company") or "").strip()
    name_company_1 = _document_company_for_display(k_co)
    doc_co_match = _document_company_for_match(k_co)
    d_pe = (ktab.get("person") or "").strip() or _extract_person_name_from_document(
        text_for_extraction
    )
    th_raw = (ktab.get("total") or "").strip() or _extract_total_work_hours_from_document(
        text_for_extraction
    )
    s_seal = (ktab.get("seal") or "").strip()
    _seal_norm = _normalize_seal_phrase(s_seal) if s_seal else ""
    row["name_company_from_file"] = fp_co
    row["name_person_from_file"] = fp_pe
    row["name_company_1"] = name_company_1
    # 互換: 従来キーも残す（勤怠表の会社セルのみ）
    row["name_company_from_doc"] = k_co
    row["name_person_from_doc"] = d_pe
    row["match_company"] = _match_company_symbol_single(fp_co, doc_co_match)
    row["user_judgment_company"] = row["match_company"]
    row["match_person"] = _compare_person(fp_pe, d_pe)
    th_dec = _work_hours_string_to_decimal(th_raw)
    row["total_hours_raw"] = th_raw
    row["total_hours_decimal"] = th_dec
    # 読取列は解析から抜いた文字列をそのまま用いる（60進/10進の別は _work_hours_string_to_decimal 側のルール）
    row["labor_read_display"] = (th_raw or "").strip() or "（なし）"
    row["seal_in_doc"] = _seal_norm or _extract_seal_in_from_document(
        text_for_extraction
    )


def _row_is_excel(row: dict[str, str]) -> bool:
    ext = Path(row.get("file_name") or "").suffix.lower()
    return ext in EXCEL_SUFFIXES or (row.get("source_kind") or "").strip() == "excel"


def _target_sheet_ok_symbol(row: dict[str, str]) -> str:
    if "target_sheet_exists" not in row:
        return ""
    v = (row.get("target_sheet_exists") or "").strip()
    return v if v in ("〇", "✖") else ""


def _escape_md_table_cell(text: str) -> str:
    """Markdown表セル内用に | と改行をエスケープする。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("|", "\\|")
    return text.replace("\n", "<br>")


# アップロードが 422 等で失敗した際、ワーカーのチャットを作り直す判断に使う
_UPLOAD_RECREATE_CHAT_STATUSES = frozenset(
    {400, 404, 407, 408, 409, 410, 422, 423, 424, 425, 426, 500}
)


def _extract_http_status_from_adk_exc(exc: BaseException) -> int | None:
    """NewtonX ADK が付ける FileUploadError 等から HTTP ステータスを推定。"""
    sc = getattr(exc, "status_code", None)
    if isinstance(sc, int):
        return sc
    msg = str(exc)
    m = re.search(r"に失敗:\s*(\d{3})\s*-", msg)
    if m:
        return int(m.group(1))
    return None


def _upload_http_error_should_recreate_chat(exc: BaseException) -> bool:
    code = _extract_http_status_from_adk_exc(exc)
    return code is not None and code in _UPLOAD_RECREATE_CHAT_STATUSES


SUMMARY_MD_HEADER = (
    f"| {SUMMARY_UPLOAD_COL} | 画像ファイル名 | {SUMMARY_TARGET_SHEET_COL} | 会社名1 | 氏名 | 社員番号 | "
    "合計勤務時間（10進） | 合計勤務時間（読取） | 押印有無 | "
    "会社名比較（ファイル名✖文書） | ユーザ判断 |"
)


def _upload_image_with_retries(
    wc: NewtonXClient,
    *,
    chat_uid_holder: dict[str, str],
    upload_src: str | Path,
    file_name: str,
    recreate_chat_fn: Callable[[BaseException], bool] | None = None,
    log_emit: Callable[[str], None] | None = None,
) -> tuple[str | None, bool]:
    """upload_image を試行する。422 などの場合は recreate_chat_fn でチャット刷新後・同一試行で再アップロード。"""
    last: str | None = None
    lg = log_emit if log_emit is not None else print

    def _recover_if_applicable(exc: BaseException, *, recovered_already: bool) -> bool:
        if not recreate_chat_fn or recovered_already:
            return False
        if not _upload_http_error_should_recreate_chat(exc):
            return False
        code = _extract_http_status_from_adk_exc(exc)
        code_s = str(code) if code is not None else "?"
        lg(f"画像アップロード失敗 HTTP {code_s} のためチャットを再作成して再試行します。")
        return bool(recreate_chat_fn(exc))

    for attempt in range(UPLOAD_MAX_RETRIES + 1):
        if attempt > 0 and UPLOAD_RETRY_DELAY_SEC > 0:
            time.sleep(UPLOAD_RETRY_DELAY_SEC)
        recovered_round = False
        while True:
            try:
                last = wc.upload_image(
                    chat_uid=chat_uid_holder["chat_uid"],
                    file_path=upload_src,
                    file_name=file_name,
                )
                if last:
                    return last, True
                break
            except FileUploadError as e:
                if _extract_http_status_from_adk_exc(e) is None:
                    lg(f"画像アップロード: リトライ不能な検証／IOエラーです — {e}")
                    return None, False
                if _recover_if_applicable(e, recovered_already=recovered_round):
                    recovered_round = True
                    continue
                last = None
                break
    return last, False


def _upload_document_with_retries(
    wc: NewtonXClient,
    *,
    chat_uid_holder: dict[str, str],
    file_path: str,
    file_name: str,
    recreate_chat_fn: Callable[[BaseException], bool] | None = None,
    log_emit: Callable[[str], None] | None = None,
) -> tuple[bool, bool]:
    """upload_document を試行する。チャット刷新ロジックは画像と同一。"""
    last_ok = False
    lg = log_emit if log_emit is not None else print

    def _recover_if_applicable(exc: BaseException, *, recovered_already: bool) -> bool:
        if not recreate_chat_fn or recovered_already:
            return False
        if not _upload_http_error_should_recreate_chat(exc):
            return False
        code = _extract_http_status_from_adk_exc(exc)
        code_s = str(code) if code is not None else "?"
        lg(f"PDFアップロード失敗 HTTP {code_s} のためチャットを再作成して再試行します。")
        return bool(recreate_chat_fn(exc))

    for attempt in range(UPLOAD_MAX_RETRIES + 1):
        if attempt > 0 and UPLOAD_RETRY_DELAY_SEC > 0:
            time.sleep(UPLOAD_RETRY_DELAY_SEC)
        recovered_round = False
        while True:
            try:
                last_ok = wc.upload_document(
                    chat_uid=chat_uid_holder["chat_uid"],
                    file_path=file_path,
                    file_name=file_name,
                )
                if last_ok:
                    return True, True
                break
            except FileUploadError as e:
                if _extract_http_status_from_adk_exc(e) is None:
                    lg(f"PDFアップロード: リトライ不能な検証／IOエラーです — {e}")
                    return False, False
                if _recover_if_applicable(e, recovered_already=recovered_round):
                    recovered_round = True
                    continue
                last_ok = False
                break
    return last_ok, False


def _upload_ok_symbol(row: dict[str, str]) -> str:
    """JSON に upload_ok が無い既存データは空欄。無効値も空欄に寄せる。"""
    if "upload_ok" not in row:
        return ""
    v = (row.get("upload_ok") or "").strip()
    return v if v in ("〇", "✖") else ""


def _one_summary_data_line(r: dict[str, str]) -> str:
    """11列1行分（集計用）。Excel 行は対象シート有無のみを埋め、AI読取列は空欄にする。"""
    is_excel = _row_is_excel(r)
    u_sym = _escape_md_table_cell(_upload_ok_symbol(r))
    fn = _escape_md_table_cell(r.get("file_name", ""))
    ts = _escape_md_table_cell(_target_sheet_ok_symbol(r))
    co1 = _escape_md_table_cell(
        ((r.get("name_company_1") or "").strip() or ("" if is_excel else "不明"))
    )
    pe = _escape_md_table_cell(
        ((r.get("name_person_from_doc") or "").strip() or ("" if is_excel else "不明"))
    )
    emp = _escape_md_table_cell(
        (r.get("employee_no") or "").strip() or ""
    )
    th = _escape_md_table_cell(
        ((
            _decimal_for_table_display((r.get("total_hours_decimal") or "").strip())
        ) if (r.get("total_hours_decimal") or "").strip() else ("" if is_excel else "（なし）"))
    )
    lr = _escape_md_table_cell(
        ((r.get("total_hours_raw") or "").strip() or ("" if is_excel else "（なし）"))
    )
    se = _escape_md_table_cell(
        "" if is_excel else ((r.get("seal_in_doc") or "").strip() or "不明")
    )
    mc = _escape_md_table_cell((r.get("match_company") or ("" if is_excel else "✖")))
    uj = (r.get("user_judgment_company") or "").strip() or (r.get("match_company") or ("" if is_excel else "✖"))
    uj = _escape_md_table_cell(uj)
    return (
        f"| {u_sym} | {fn} | {ts} | {co1} | {pe} | {emp} | {th} | {lr} | {se} | {mc} | {uj} |"
    )


def _summary_table_md_lines(results: list[dict[str, str]]) -> list[str]:
    """
    解析結果.md 用: ヘッダ1行＋区切り行なし。データは各1行。
    """
    return [SUMMARY_MD_HEADER] + [_one_summary_data_line(r) for r in results]


def summary_header_cells() -> tuple[str, ...]:
    """SUMMARY_MD_HEADER から見出し要素を返す（Treeview 列名用）。"""
    raw = SUMMARY_MD_HEADER.strip()
    if raw.startswith("|"):
        raw = raw[1:]
    if raw.endswith("|"):
        raw = raw[:-1]
    return tuple(c.strip() for c in raw.split("|"))


def row_display_values(r: dict[str, str]) -> tuple[str, ...]:
    """グリッド表示用（Markdown エスケープなし）。 _one_summary_data_line と同一ルール。"""
    is_excel = _row_is_excel(r)
    up_sym = _upload_ok_symbol(r)
    fn = r.get("file_name", "") or ""
    ts = _target_sheet_ok_symbol(r)
    co1 = ((r.get("name_company_1") or "").strip() or ("" if is_excel else "不明"))
    pe = ((r.get("name_person_from_doc") or "").strip() or ("" if is_excel else "不明"))
    emp = (r.get("employee_no") or "").strip() or ""
    th = (
        _decimal_for_table_display((r.get("total_hours_decimal") or "").strip())
        if (r.get("total_hours_decimal") or "").strip()
        else ("" if is_excel else "（なし）")
    )
    lr = ((r.get("total_hours_raw") or "").strip() or ("" if is_excel else "（なし）"))
    se = "" if is_excel else ((r.get("seal_in_doc") or "").strip() or "不明")
    mc = (r.get("match_company") or ("" if is_excel else "✖"))
    uj = (r.get("user_judgment_company") or "").strip() or mc
    return (up_sym, fn, ts, co1, pe, emp, th, lr, se, mc, uj)


def run_analysis(
    client: NewtonXClient,
    assistant_uid: str,
    data_dir: Path,
    *,
    save_md_path: Path | None = None,
    on_log: Callable[[str], None] | None = None,
    emit_progress_md_rows: bool = True,
    on_file_progress: Callable[[int, int], None] | None = None,
    on_row_completed: Callable[[dict[str, str]], None] | None = None,
    cancel_event=None,
    skip_file_names: set[str] | None = None,
    parallel_chats: int = DEFAULT_PARALLEL_ANALYSIS_CHATS,
) -> list[dict[str, str]]:
    """
    指定フォルダ直下の画像・PDF・Excel を名前順で処理し、結果行のリストを返す。
    NewtonX では parallel_chats 本のチャットを同一フォルダに作成し、
    ファイルをラウンドロビンで割り当ててスレッド並列処理する（終了時は全スレッド join）。
    各ワーカーは独自の NewtonXClient で API を呼び出す。
    アップロードは、失敗のたびに最大 UPLOAD_MAX_RETRIES 回まで再試行します（総試行は多くとも 1+UPLOAD_MAX_RETRIES 回）。
    Excel は生成AIへアップロードせず、対象シート名の有無のみを判定する。
    save_md_path が None のときは cwd に 解析結果.md を出力する。
    emit_progress_md_rows が False のとき、コンソール向け Markdown 行は on_log に流さない。
    on_file_progress: 処理が終わったファイル数を (実行済み数, 対象総数) で通知する（各ファイルの試行の末尾で1回）。
    on_row_completed: 解析結果またはアップロード失敗行を1件 results に載せた直後に呼ぶ。
    cancel_event: threading.Event 互換。set() されたら以降の解析を中断する（処理中の1ファイルは止まらず、次の境界で止まる）。
    skip_file_names: ここに含まれるファイル名は解析処理自体をスキップする（progress は進める）。
    """
    log = on_log if on_log is not None else print

    if not data_dir.is_dir():
        raise FileNotFoundError(f"フォルダが存在しません: {data_dir}")

    skip_names = {str(x).strip() for x in (skip_file_names or set()) if str(x).strip()}

    image_files = sorted(
        p for p in data_dir.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES and p.name not in skip_names
    )
    pdf_files = sorted(
        p for p in data_dir.iterdir()
        if p.is_file() and p.suffix.lower() == PDF_SUFFIX and p.name not in skip_names
    )
    excel_files = sorted(
        p for p in data_dir.iterdir()
        if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES and p.name not in skip_names
    )
    if not image_files and not pdf_files and not excel_files:
        raise ValueError(f"画像・PDF・Excelファイルが見つかりません: {data_dir}")

    total_files = len(image_files) + len(pdf_files) + len(excel_files)
    progress_done = 0
    progress_lock = threading.Lock()

    def is_cancelled() -> bool:
        try:
            return bool(cancel_event is not None and cancel_event.is_set())
        except Exception:
            return False

    def notify_file_finished() -> None:
        nonlocal progress_done
        with progress_lock:
            progress_done += 1
            cur = progress_done
        if on_file_progress is not None:
            on_file_progress(cur, total_files)

    if on_file_progress is not None:
        on_file_progress(0, total_files)

    if is_cancelled():
        log("中断: 解析を開始する前にキャンセルされました")
        return []

    target_folder_name = "業務課集計"
    folders = client.get_folders()
    matched = next(
        (
            f
            for f in folders
            if (f.get("name") or "").strip() == target_folder_name
        ),
        None,
    )
    if matched is not None:
        raw_id = (
            matched.get("uid")
            if matched.get("uid") is not None
            else matched.get("id")
        )
        folder_uid = str(raw_id) if raw_id is not None else None
        if not folder_uid:
            raise RuntimeError(
                "NewtonX のフォルダ一覧に該当がありますが、フォルダ ID を取得できませんでした。"
            )
        log(f"既存フォルダを使用します: {target_folder_name} ({folder_uid})")
    else:
        folder_uid = client.create_folder(target_folder_name)
        if not folder_uid:
            raise RuntimeError("NewtonX 上でフォルダの作成に失敗しました。")
        log(f"フォルダが作成されました: {folder_uid}")

    n_workers = int(parallel_chats)
    if n_workers < 1:
        n_workers = DEFAULT_PARALLEL_ANALYSIS_CHATS
    elif n_workers > PARALLEL_WORKERS_MAX:
        n_workers = PARALLEL_WORKERS_MAX

    chat_uids: list[str] = []
    for wi in range(n_workers):
        chat_uid = client.create_chat_in_folder(
            assistant_uid=assistant_uid,
            folder_uid=folder_uid,
            title=f"業務課集計 #{wi + 1}",
        )
        if chat_uid:
            client.move_chat_to_folder(chat_uid=chat_uid, folder_uid=folder_uid)
        if not chat_uid:
            raise RuntimeError(
                f"NewtonX 上でチャットの作成に失敗しました（{wi + 1} / {n_workers}）。"
            )
        log(f"チャットが作成されました（並列ワーカー {wi + 1}）: {chat_uid}")
        chat_uids.append(chat_uid)

    summary_header_done = False
    summary_lock = threading.Lock()

    def emit_summary_row_md(row_dict: dict[str, str]) -> None:
        nonlocal summary_header_done
        if not emit_progress_md_rows:
            return
        with summary_lock:
            if not summary_header_done:
                log(SUMMARY_MD_HEADER)
                summary_header_done = True
            log(_one_summary_data_line(row_dict))

    tasks_ordered: list[tuple[str, Path]] = [
        ("image", p) for p in image_files
    ] + [("pdf", p) for p in pdf_files] + [("excel", p) for p in excel_files]

    buckets: list[list[tuple[str, Path]]] = [[] for _ in range(n_workers)]
    for i, task in enumerate(tasks_ordered):
        buckets[i % n_workers].append(task)

    bucket_results: list[list[dict[str, str]]] = [[] for _ in range(n_workers)]
    worker_exc: list[BaseException | None] = [None] * n_workers

    def append_upload_failure_row(
        worker_ix: int,
        file_path: Path,
        *,
        prefer_pdf_section: bool,
        message: str,
    ) -> None:
        row = {
            "upload_ok": "✖",
            "file_name": file_path.name,
            "resolved_path": str(file_path.resolve()),
            "analysis": message,
        }
        _enrich_with_match_scores(
            row,
            row["analysis"],
            prefer_kintai_section=prefer_pdf_section,
        )
        bucket_results[worker_ix].append(row)
        emit_summary_row_md(row)
        if on_row_completed is not None:
            on_row_completed(row)

    def worker_loop(worker_idx: int, chat_uid: str, tasks: list[tuple[str, Path]]) -> None:
        wc = create_client()
        chat_holder: dict[str, str] = {"chat_uid": chat_uid}

        def recreate_worker_chat(_exc: BaseException) -> bool:
            """アップロード系 HTTP エラー時に、本ワーカーの作業チャットを作り直す。"""
            try:
                nc = wc.create_chat_in_folder(
                    assistant_uid=assistant_uid,
                    folder_uid=folder_uid,
                    title=f"業務課集計 #{worker_idx + 1}",
                )
                if not nc:
                    log(f"ワーカー {worker_idx + 1}: チャットの再作成に失敗しました（API）。")
                    return False
                try:
                    wc.move_chat_to_folder(chat_uid=nc, folder_uid=folder_uid)
                except Exception as me:
                    log(
                        f"ワーカー {worker_idx + 1}: "
                        f"再作成チャットのフォルダ移動に警告 — {me}"
                    )
                chat_holder["chat_uid"] = str(nc).strip()
                log(
                    f"ワーカー {worker_idx + 1}: "
                    f"チャットを再作成しました: {chat_holder['chat_uid']}"
                )
                return True
            except Exception as e:
                log(f"ワーカー {worker_idx + 1}: チャット再作成処理で異常 — {e}")
                return False

        def signal_fatal(exc: BaseException) -> None:
            worker_exc[worker_idx] = exc
            try:
                if cancel_event is not None:
                    cancel_event.set()
            except Exception:
                pass

        for kind, file_path in tasks:
            if is_cancelled():
                log(
                    f"中断: ワーカー {worker_idx + 1} がキャンセルにより停止しました"
                )
                break
            try:
                if kind == "excel":
                    row_excel = _extract_excel_target_sheet_row(file_path)
                    bucket_results[worker_idx].append(row_excel)
                    emit_summary_row_md(row_excel)
                    if on_row_completed is not None:
                        on_row_completed(row_excel)
                    if row_excel.get("analysis"):
                        log(f"Excel処理警告: {file_path.name}: {row_excel['analysis']}")
                elif kind == "image":
                    try:
                        upload_src, tmp_upload = _resolve_upload_path(file_path)
                    except (
                        OSError,
                        RuntimeError,
                        ValueError,
                        UnidentifiedImageError,
                    ) as e:
                        log(
                            f"スキップ: 画像の読み込み／圧縮に失敗しました — "
                            f"{file_path.name}: {e}"
                        )
                        append_upload_failure_row(
                            worker_idx,
                            file_path,
                            prefer_pdf_section=False,
                            message="（画像の読み込み／圧縮に失敗しました）",
                        )
                    else:
                        if is_cancelled():
                            log("中断: 画像のアップロード前にキャンセルされました")
                            break
                        tmp_upload_path: Path | None = tmp_upload
                        try:
                            image_id, upload_succeeded = _upload_image_with_retries(
                                wc,
                                chat_uid_holder=chat_holder,
                                upload_src=upload_src,
                                file_name=file_path.name,
                                recreate_chat_fn=recreate_worker_chat,
                                log_emit=log,
                            )
                            if not image_id:
                                log(
                                    f"スキップ: アップロードに失敗しました（最大 "
                                    f"{UPLOAD_MAX_RETRIES} 回までリトライ）— "
                                    f"{file_path.name}"
                                )
                                append_upload_failure_row(
                                    worker_idx,
                                    file_path,
                                    prefer_pdf_section=False,
                                    message=(
                                        f"（画像アップロード失敗／{UPLOAD_MAX_RETRIES}回リトライまで）"
                                    ),
                                )
                            else:
                                if is_cancelled():
                                    log(
                                        "中断: 画像の解析要求前にキャンセルされました"
                                    )
                                    break
                                response = wc.send_message(
                                    chat_uid=chat_holder["chat_uid"],
                                    message=_build_check_message(file_path.name),
                                    image_ids=[image_id],
                                )
                                if is_cancelled():
                                    log(
                                        "中断: 画像の解析応答待ち後にキャンセルされました"
                                    )
                                    break
                                row = {
                                    "upload_ok": "〇" if upload_succeeded else "✖",
                                    "file_name": file_path.name,
                                    "resolved_path": str(file_path.resolve()),
                                    "analysis": response
                                    if response
                                    else "（解析結果を取得できませんでした）",
                                }
                                _enrich_with_match_scores(row, row["analysis"])
                                bucket_results[worker_idx].append(row)
                                emit_summary_row_md(row)
                                if on_row_completed is not None:
                                    on_row_completed(row)
                        finally:
                            if tmp_upload_path is not None:
                                tmp_upload_path.unlink(missing_ok=True)
                else:
                    try:
                        if is_cancelled():
                            log("中断: PDFアップロード前にキャンセルされました")
                            break
                        success, upload_succeeded = _upload_document_with_retries(
                            wc,
                            chat_uid_holder=chat_holder,
                            file_path=str(file_path),
                            file_name=file_path.name,
                            recreate_chat_fn=recreate_worker_chat,
                            log_emit=log,
                        )
                        if not success:
                            log(
                                f"スキップ: PDFのアップロードに失敗しました（最大 "
                                f"{UPLOAD_MAX_RETRIES} 回までリトライ）— "
                                f"{file_path.name}"
                            )
                            append_upload_failure_row(
                                worker_idx,
                                file_path,
                                prefer_pdf_section=True,
                                message=(
                                    f"（PDFアップロード失敗／{UPLOAD_MAX_RETRIES}回リトライまで）"
                                ),
                            )
                        else:
                            log(f"PDFがアップロードされました: {file_path.name}")
                            if is_cancelled():
                                log(
                                    "中断: PDFの解析要求前にキャンセルされました"
                                )
                                break
                            response = wc.send_message(
                                chat_uid=chat_holder["chat_uid"],
                                message=_build_pdf_check_message(file_path.name),
                            )
                            if is_cancelled():
                                log(
                                    "中断: PDFの解析応答待ち後にキャンセルされました"
                                )
                                break
                            row2 = {
                                "upload_ok": "〇" if upload_succeeded else "✖",
                                "file_name": file_path.name,
                                "resolved_path": str(file_path.resolve()),
                                "analysis": response
                                if response
                                else "（解析結果を取得できませんでした）",
                            }
                            _enrich_with_match_scores(
                                row2,
                                row2["analysis"],
                                prefer_kintai_section=True,
                            )
                            bucket_results[worker_idx].append(row2)
                            emit_summary_row_md(row2)
                            if on_row_completed is not None:
                                on_row_completed(row2)
                    except Exception as e:
                        log(
                            f"スキップ: PDFの処理に失敗しました — "
                            f"{file_path.name}: {e}"
                        )
            except BaseException as e:
                signal_fatal(e)
                break
            finally:
                notify_file_finished()

    threads = [
        threading.Thread(
            target=worker_loop,
            args=(wi, chat_uids[wi], buckets[wi]),
            name=f"kintai-worker-{wi}",
            daemon=False,
        )
        for wi in range(n_workers)
    ]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    results = []
    for wi in range(n_workers):
        results.extend(bucket_results[wi])

    def _row_kind_order(row: dict[str, str]) -> tuple[int, str]:
        fn = row.get("file_name") or ""
        ext = Path(fn).suffix.lower()
        if ext in IMAGE_SUFFIXES:
            kind_order = 0
        elif ext == PDF_SUFFIX:
            kind_order = 1
        elif ext in EXCEL_SUFFIXES:
            kind_order = 2
        else:
            kind_order = 9
        return (kind_order, fn)

    results.sort(key=_row_kind_order)

    fatal = next((e for e in worker_exc if e is not None), None)
    if fatal is not None:
        raise fatal

    if is_cancelled():
        log("中断: 結果ファイルの出力前にキャンセルされました")
        return results

    output_md = save_md_path if save_md_path is not None else Path.cwd() / "解析結果.md"
    if results:
        lines_content = "\n".join(_summary_table_md_lines(results))
    else:
        lines_content = "（アップロード・解析に成功した画像・PDFがありませんでした）"
    output_md.write_text(lines_content + "\n", encoding="utf-8")
    log(f"\n解析結果を保存しました: {output_md.resolve()}")
    return results
