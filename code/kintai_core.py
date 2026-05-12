"""勤務表画像・PDF解析（NewtonX）のコア処理。"""
from __future__ import annotations

from collections.abc import Callable

from newtonx_adk import NewtonXClient, ConfigManager
import json
import math
import re
import tempfile
import types
import unicodedata
from pathlib import Path

from PIL import Image, ImageOps, UnidentifiedImageError

# バイナリ換算の 1MiB。「未満」なのでこれ未満のサイズになるよう調整する
MIB = 1024 * 1024

IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff"}
PDF_SUFFIX = ".pdf"

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
画像ファイル名の会社名は勤務先会社名として抽出しないこと。
画像ファイル名（アップロードファイル名）として、次の名前のみを記載してください（サーバー側のIDや別名は使わないこと）:
{display_file_name}

勤務先会社名：
氏名：
合計勤務時間：（8:20・8時間20分・8.20・101_10H・86.17H 等。必要に応じて実データの表記のまま）
押印有無：〇/✖

"""


def _build_pdf_check_message(display_file_name: str) -> str:
    """PDF用。アップロード直後は send_message にドキュメントIDを渡さなくても参照できる想定。"""
    return f"""
同じPDFに勤怠（出退勤・打刻・勤務時間等）の情報と経費精算（領収書・立替等）の情報の両方が含まれている場合は、経費精算は無視し、必ず勤怠（勤務表）の情報だけを根拠に回答してください。経費側の社名・氏名・金額は採用しないでください。
PDFファイル名の会社名は勤務先会社名として抽出しないこと。
勤務表（勤怠）のPDFを解析し、１ファイルあたり適切な行数で以下の内容にあたるものをmd形式で表で出力してください。出力する勤務先・氏名・合計勤務時間・押印はすべて勤怠部分の記載に基づきます。

PDFファイル名（アップロードファイル名）として、次の名前のみを記載してください（サーバー側のIDや別名は使わないこと）:
{display_file_name}

勤務先会社名：
氏名：
合計勤務時間：（8:20・8時間20分・8.20・101_10H・86.17H 等。必要に応じて実データの表記のまま）
押印有無：〇/✖

"""


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
        "勤務先会社名",
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
            t in ("勤務先会社名", "勤務先", "勤務先会社", "就業先")
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
            "勤務先会社名",
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
    for kw in ("勤怠", "勤務表", "出退勤", "出退", "合計勤務", "合計勤務時間", "勤務先会社名"):
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


_SERAKU_MARK_RE = re.compile(r"セラク|seraku", re.IGNORECASE)


def _split_company_segments(text: str) -> list[str]:
    """1セル内の複数会社名を分割（区切りで分割）。"""
    t = (text or "").strip()
    if not t:
        return []
    parts = re.split(r"[/／、,，｜\|\n]+", t)
    return [p.strip() for p in parts if p.strip()]


def _contains_seraku_company_mark(s: str) -> bool:
    """セラク／英字 seraku（Seraku 等）を含む社名は当社関連として照合・一覧から除外する。
    誤表記の SELAC は当社名として扱わない。
    """
    t = unicodedata.normalize("NFKC", (s or "").strip())
    if not t:
        return False
    nospace = re.sub(r"\s+", "", t)
    return bool(_SERAKU_MARK_RE.search(nospace))


def _collect_company_candidates_from_document(text: str) -> list[tuple[int, str]]:
    """文面から 勤務先会社名 の候補を出現順に収集（重複は位置が別ならそのまま）。"""
    if not (text and str(text).strip()):
        return []
    raw = str(text)
    cands: list[tuple[int, str]] = []
    for m in re.finditer(
        r"勤務先会社名\s*[：:｜\|]?\s*([^\n\r\|]+?)(?=(\s*\||$|\n))",
        raw,
    ):
        v = m.group(1).strip()
        if v and v not in ("", "-", "—", "―", "N/A", "n/a", "要確認", "未記載", "null"):
            if not _is_table_headerish_cell(v):
                cands.append((m.start(1), v))
    for m in re.finditer(
        r"\|[^\n]*?勤務先会社名(?:\s*\|)?\s*([^\n\|]+)", raw, re.IGNORECASE
    ):
        v = m.group(1).strip()
        if v and v not in ("", "-", "—") and not _is_table_headerish_cell(v):
            cands.append((m.start(1), v))
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands


def _merge_company_names_for_display(
    ktab_company: str,
    doc_text: str,
) -> tuple[str, list[str]]:
    """
    セラクまたは英字 **seraku** を**含む**社名は一覧・照合から除外する（SELAC 表記は対象外）。
    戻り値: (表示用文字列 "A / B", 照合に使う当社以外の会社名リスト)
    """
    ordered_raw: list[str] = []
    if (ktab_company or "").strip():
        ordered_raw.append(ktab_company.strip())
    for _pos, c in _collect_company_candidates_from_document(doc_text):
        ordered_raw.append(c.strip())

    flat: list[str] = []
    seen_key: set[str] = set()
    for block in ordered_raw:
        for seg in _split_company_segments(block):
            if _contains_seraku_company_mark(seg):
                continue
            key = _normalize_mixed(seg)
            if key not in seen_key:
                seen_key.add(key)
                flat.append(seg)
    if not flat:
        return "不明", []
    display = " / ".join(flat)
    return display, flat


def _file_company_for_match(company_from_file: str) -> str:
    """ファイル名由来の会社。セラク／seraku を含むときは照合対象外（空）。"""
    s = (company_from_file or "").strip()
    if not s or _contains_seraku_company_mark(s):
        return ""
    return s


def _best_match_company_symbol(file_co: str, doc_companies: list[str]) -> str:
    """ファイル名会社と、文書の（当社・seraku 以外の）会社名リストの最良照合（〇>△>✖）。"""
    fp = _file_company_for_match(file_co)
    if not doc_companies:
        return "〇" if not fp else "✖"
    rank = {"〇": 3, "△": 2, "✖": 1}
    best = "✖"
    for d in doc_companies:
        ds = (d or "").strip()
        if not ds:
            continue
        sym = _compare_company(fp, ds) if fp else "✖"
        if rank[sym] > rank[best]:
            best = sym
        if best == "〇":
            break
    if not fp and doc_companies:
        return "△"
    return best


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


def _extract_company_name_from_document(text: str) -> str:
    """後方互換: 単一代表名（当社・seraku 以外を優先、複数は先頭のみ）。"""
    disp, lst = _merge_company_names_for_display("", text)
    if lst:
        return lst[0]
    return ""


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

    def _employee_no_from_file_name(file_name: str) -> str:
        stem = Path(file_name or "").stem
        if len(stem) < 7:
            return ""
        tail7 = stem[-7:]
        if tail7.isdigit():
            return tail7
        return "社員番号エラー"

    text_for_extraction = (
        _prefer_kintai_text_for_extraction(analysis) if prefer_kintai_section else analysis
    )
    fn = row.get("file_name", "")
    fp_co, fp_pe = _parse_filename_company_and_person(fn)
    row["employee_no"] = _employee_no_from_file_name(fn)
    ktab = _extract_kintai_from_markdown_table(text_for_extraction, fn) or {}
    k_co = (ktab.get("company") or "").strip()
    disp_co, doc_co_list = _merge_company_names_for_display(
        k_co, text_for_extraction
    )

    # 複数社名が取れた場合は「会社名1」「会社名2」に分けて保持
    name_company_1 = ""
    name_company_2 = ""
    if doc_co_list:
        name_company_1 = (doc_co_list[0] or "").strip()
        name_company_2 = (doc_co_list[1] or "").strip() if len(doc_co_list) >= 2 else ""
    if not name_company_1 and disp_co.strip() and disp_co != "不明":
        # 念のため display からも補完
        parts = [p.strip() for p in disp_co.split("/") if p.strip()]
        if parts:
            name_company_1 = parts[0]
            name_company_2 = parts[1] if len(parts) >= 2 else ""
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
    row["name_company_1"] = name_company_1 or "不明"
    row["name_company_2"] = name_company_2
    # 互換: 従来キーも残す
    row["name_company_from_doc"] = (disp_co or "").strip() if disp_co else ""
    row["name_person_from_doc"] = d_pe
    row["match_company"] = _best_match_company_symbol(fp_co, doc_co_list)
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


def _escape_md_table_cell(text: str) -> str:
    """Markdown表セル内用に | と改行をエスケープする。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("|", "\\|")
    return text.replace("\n", "<br>")


SUMMARY_MD_HEADER = (
    "| 画像ファイル名 | 会社名1 | 会社名2 | 氏名 | 社員番号 | 合計勤務時間（10進） | "
    "合計勤務時間（読取） | 押印有無 | 会社名比較（ファイル名✖文書） | ユーザ判断 |"
)


def _one_summary_data_line(r: dict[str, str]) -> str:
    """10列1行分（集計用）。ユーザ判断は初期値は match_company と同じ。"""
    fn = _escape_md_table_cell(r.get("file_name", ""))
    co1 = _escape_md_table_cell((r.get("name_company_1") or "").strip() or "不明")
    co2 = _escape_md_table_cell((r.get("name_company_2") or "").strip() or "")
    pe = _escape_md_table_cell(
        (r.get("name_person_from_doc") or "").strip() or "不明"
    )
    emp = _escape_md_table_cell((r.get("employee_no") or "").strip() or "")
    th = _escape_md_table_cell(
        _decimal_for_table_display((r.get("total_hours_decimal") or "").strip())
    )
    lr = _escape_md_table_cell(
        (r.get("total_hours_raw") or "").strip() or "（なし）"
    )
    se = _escape_md_table_cell(
        (r.get("seal_in_doc") or "").strip() or "不明"
    )
    mc = _escape_md_table_cell(r.get("match_company", "✖"))
    uj = (r.get("user_judgment_company") or "").strip() or (r.get("match_company") or "✖")
    uj = _escape_md_table_cell(uj)
    return f"| {fn} | {co1} | {co2} | {pe} | {emp} | {th} | {lr} | {se} | {mc} | {uj} |"


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
    fn = r.get("file_name", "") or ""
    co1 = (r.get("name_company_1") or "").strip() or "不明"
    co2 = (r.get("name_company_2") or "").strip() or ""
    pe = (r.get("name_person_from_doc") or "").strip() or "不明"
    emp = (r.get("employee_no") or "").strip() or ""
    th = _decimal_for_table_display((r.get("total_hours_decimal") or "").strip())
    lr = (r.get("total_hours_raw") or "").strip() or "（なし）"
    se = (r.get("seal_in_doc") or "").strip() or "不明"
    mc = r.get("match_company", "✖") or "✖"
    uj = (r.get("user_judgment_company") or "").strip() or mc
    return (fn, co1, co2, pe, emp, th, lr, se, mc, uj)


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
) -> list[dict[str, str]]:
    """
    指定フォルダ直下の画像と PDF を名前順で解析し、結果行のリストを返す。
    save_md_path が None のときは cwd に 解析結果.md を出力する。
    emit_progress_md_rows が False のとき、コンソール向け Markdown 行は on_log に流さない。
    on_file_progress: 処理が終わったファイル数を (実行済み数, 対象総数) で通知する（各ファイルの試行の末尾で1回）。
    on_row_completed: 解析結果を1件 results に載せた直後に呼ぶ（画像・PDF それぞれ成功時のみ）。
    cancel_event: threading.Event 互換。set() されたら以降の解析を中断する（処理中の1ファイルは止まらず、次の境界で止まる）。
    skip_file_names: ここに含まれるファイル名は解析処理自体をスキップする（progress は進める）。
    """
    log = on_log if on_log is not None else print
    results: list[dict[str, str]] = []

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
    if not image_files and not pdf_files:
        raise ValueError(f"画像またはPDFファイルが見つかりません: {data_dir}")

    total_files = len(image_files) + len(pdf_files)
    progress_done = 0

    def is_cancelled() -> bool:
        try:
            return bool(cancel_event is not None and cancel_event.is_set())
        except Exception:
            return False

    def notify_file_finished() -> None:
        nonlocal progress_done
        progress_done += 1
        if on_file_progress is not None:
            on_file_progress(progress_done, total_files)

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
        raw_id = matched.get("uid") if matched.get("uid") is not None else matched.get(
            "id"
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

    chat_id: str | None = None
    if image_files:
        chat_uid = client.create_chat_in_folder(
            assistant_uid=assistant_uid,
            folder_uid=folder_uid,
            title="業務課集計",
        )
        if not chat_uid:
            raise RuntimeError("画像用チャットの作成に失敗しました。")
        log(f"チャットが作成されました（画像用）: {chat_uid}")
        client.move_chat_to_folder(chat_uid=chat_uid, folder_uid=folder_uid)
        log(f"チャットが移動されました: {chat_uid}")
        chat_id = chat_uid

    summary_header_done = False

    def emit_summary_row_md(row_dict: dict[str, str]) -> None:
        nonlocal summary_header_done
        if not emit_progress_md_rows:
            return
        if not summary_header_done:
            log(SUMMARY_MD_HEADER)
            summary_header_done = True
        log(_one_summary_data_line(row_dict))

    for file_path in image_files:
        if is_cancelled():
            log("中断: 画像の解析を中断しました")
            break
        try:
            try:
                upload_src, tmp_upload = _resolve_upload_path(file_path)
            except (OSError, RuntimeError, ValueError, UnidentifiedImageError) as e:
                log(
                    f"スキップ: 画像の読み込み／圧縮に失敗しました — {file_path.name}: {e}"
                )
            else:
                if is_cancelled():
                    log("中断: 画像のアップロード前にキャンセルされました")
                    break
                tmp_upload_path: Path | None = tmp_upload
                try:
                    image_id = client.upload_image(
                        chat_uid=chat_id,
                        file_path=upload_src,
                        file_name=file_path.name,
                    )
                    if not image_id:
                        log(
                            f"スキップ: アップロードに失敗しました — {file_path.name}"
                        )
                    else:
                        if is_cancelled():
                            log("中断: 画像の解析要求前にキャンセルされました")
                            break
                        response = client.send_message(
                            chat_uid=chat_id,
                            message=_build_check_message(file_path.name),
                            image_ids=[image_id],
                        )
                        if is_cancelled():
                            log("中断: 画像の解析応答待ち後にキャンセルされました")
                            break
                        row = {
                            "file_name": file_path.name,
                            "resolved_path": str(file_path.resolve()),
                            "analysis": response
                            if response
                            else "（解析結果を取得できませんでした）",
                        }
                        _enrich_with_match_scores(row, row["analysis"])
                        results.append(row)
                        emit_summary_row_md(row)
                        if on_row_completed is not None:
                            on_row_completed(row)
                finally:
                    if tmp_upload_path is not None:
                        tmp_upload_path.unlink(missing_ok=True)
        finally:
            notify_file_finished()

    for file_path in pdf_files:
        if is_cancelled():
            log("中断: PDFの解析を中断しました")
            break
        try:
            try:
                if is_cancelled():
                    log("中断: PDFチャット作成前にキャンセルされました")
                    break
                pdf_chat_uid = client.create_chat_in_folder(
                    assistant_uid=assistant_uid,
                    folder_uid=folder_uid,
                    title=f"業務課集計 PDF: {file_path.name}",
                )
                if not pdf_chat_uid:
                    log(
                        f"スキップ: PDF用チャットの作成に失敗しました — {file_path.name}"
                    )
                else:
                    log(
                        f"PDF用チャットを作成しました: {pdf_chat_uid} ({file_path.name})"
                    )
                    client.move_chat_to_folder(
                        chat_uid=pdf_chat_uid, folder_uid=folder_uid
                    )

                    if is_cancelled():
                        log("中断: PDFアップロード前にキャンセルされました")
                        break
                    success = client.upload_document(
                        chat_uid=pdf_chat_uid,
                        file_path=str(file_path),
                        file_name=file_path.name,
                    )
                    if not success:
                        log(
                            f"スキップ: PDFのアップロードに失敗しました — {file_path.name}"
                        )
                    else:
                        log(f"PDFがアップロードされました: {file_path.name}")

                        if is_cancelled():
                            log("中断: PDFの解析要求前にキャンセルされました")
                            break
                        response = client.send_message(
                            chat_uid=pdf_chat_uid,
                            message=_build_pdf_check_message(file_path.name),
                        )
                        if is_cancelled():
                            log("中断: PDFの解析応答待ち後にキャンセルされました")
                            break
                        row2 = {
                            "file_name": file_path.name,
                            "resolved_path": str(file_path.resolve()),
                            "analysis": response
                            if response
                            else "（解析結果を取得できませんでした）",
                        }
                        _enrich_with_match_scores(
                            row2, row2["analysis"], prefer_kintai_section=True
                        )
                        results.append(row2)
                        emit_summary_row_md(row2)
                        if on_row_completed is not None:
                            on_row_completed(row2)
            except Exception as e:
                log(f"スキップ: PDFの処理に失敗しました — {file_path.name}: {e}")
        finally:
            notify_file_finished()

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
