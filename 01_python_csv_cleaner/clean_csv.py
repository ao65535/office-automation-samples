"""顧客名簿などのCSVを整形し、重複をチェックするツール。

使い方:
    python clean_csv.py 入力ファイル.csv [出力フォルダ]

出力:
    cleaned.csv     整形済みのデータ（重複を除いたもの）。Excelでそのまま開ける
    duplicates.csv  重複していた行（どの行と重複したかを記録）

標準ライブラリだけで動きます（追加のインストールは不要）。
"""

from __future__ import annotations

import csv
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

# ---- 列名の設定（実際のCSVに合わせて変更します） ----
COL_NAME = "氏名"
COL_EMAIL = "メールアドレス"
COL_PHONE = "電話番号"
COL_POSTAL = "郵便番号"
COL_DATE = "登録日"
COL_CHECK = "要確認"  # 出力時に追加する列

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


# ---- 1項目ずつの整形 ----
def to_halfwidth(text: str) -> str:
    """全角の英数字・記号を半角にする（NFKC正規化）。"""
    return unicodedata.normalize("NFKC", text)


def clean_name(value: str) -> str:
    """前後の空白を消し、姓と名の間の空白を半角1つにそろえる。"""
    value = value.replace("　", " ").strip()
    return re.sub(r" +", " ", value)


def clean_email(value: str) -> tuple[str, str | None]:
    value = to_halfwidth(value).strip().lower()
    if not value:
        return "", None
    if not EMAIL_RE.match(value):
        return value, "メールアドレスの形式"
    return value, None


MOBILE_PREFIXES = {"050", "070", "080", "090"}


def clean_phone(value: str) -> tuple[str, str | None]:
    """電話番号をハイフン区切りにそろえる。

    - 携帯・IP電話（050/070/080/090）は 11桁で 3-4-4
    - 東京(03)・大阪(06) は 2-4-4
    - それ以外の固定電話は市外局番の桁数が地域で異なるため、
      元の区切り位置を残す（区切りがなければ数字のまま）
    """
    half = to_halfwidth(value)
    digits = re.sub(r"\D", "", half)
    if not digits:
        return "", None

    is_mobile = digits[:3] in MOBILE_PREFIXES
    if (is_mobile and len(digits) != 11) or (not is_mobile and len(digits) != 10):
        return value.strip(), "電話番号の桁数"

    if is_mobile:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}", None
    if digits[:2] in {"03", "06"}:
        return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}", None
    groups = re.findall(r"\d+", half)
    return "-".join(groups), None


def clean_postal(value: str) -> tuple[str, str | None]:
    digits = re.sub(r"\D", "", to_halfwidth(value))
    if not digits:
        return "", None
    if len(digits) != 7:
        return value.strip(), "郵便番号の桁数"
    return f"{digits[:3]}-{digits[3:]}", None


DATE_PATTERNS = [
    re.compile(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$"),
    re.compile(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$"),
    re.compile(r"^(\d{4})(\d{2})(\d{2})$"),
]


def clean_date(value: str) -> tuple[str, str | None]:
    value = to_halfwidth(value).strip()
    if not value:
        return "", None
    for pattern in DATE_PATTERNS:
        m = pattern.match(value)
        if m:
            y, mo, d = (int(x) for x in m.groups())
            if 1 <= mo <= 12 and 1 <= d <= 31:
                return f"{y:04d}-{mo:02d}-{d:02d}", None
    return value, "日付の形式"


# ---- 1行の整形 ----
def clean_row(row: dict[str, str]) -> dict[str, str]:
    out = {k: (v or "").strip() for k, v in row.items()}
    problems: list[str] = []

    if COL_NAME in out:
        out[COL_NAME] = clean_name(out[COL_NAME])
    for col, func in (
        (COL_EMAIL, clean_email),
        (COL_PHONE, clean_phone),
        (COL_POSTAL, clean_postal),
        (COL_DATE, clean_date),
    ):
        if col in out:
            out[col], problem = func(out[col])
            if problem:
                problems.append(problem)

    out[COL_CHECK] = "、".join(problems)
    return out


def duplicate_key(row: dict[str, str]) -> str:
    """重複を判定するキー。メールアドレスがあればそれを、なければ氏名＋電話番号を使う。"""
    email = row.get(COL_EMAIL, "")
    if email:
        return "email:" + email
    name = row.get(COL_NAME, "").replace(" ", "")
    phone_digits = re.sub(r"\D", "", row.get(COL_PHONE, ""))
    return "name_phone:" + name + "|" + phone_digits


# ---- ファイル全体の処理 ----
@dataclass
class Result:
    cleaned: list[dict[str, str]] = field(default_factory=list)
    duplicates: list[dict[str, str]] = field(default_factory=list)
    fieldnames: list[str] = field(default_factory=list)


def process(rows: list[dict[str, str]], fieldnames: list[str]) -> Result:
    result = Result(fieldnames=[*fieldnames, COL_CHECK])
    first_seen: dict[str, int] = {}  # キー -> 元ファイルの行番号

    for line_no, raw in enumerate(rows, start=2):  # 1行目は見出し
        row = clean_row(raw)
        key = duplicate_key(row)
        if key in first_seen:
            dup = {"元の行番号": str(line_no), "重複先の行番号": str(first_seen[key]), **row}
            result.duplicates.append(dup)
            continue
        first_seen[key] = line_no
        result.cleaned.append(row)
    return result


def read_csv(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    """UTF-8（BOM付き含む）と Shift_JIS(cp932) の両方に対応して読み込む。"""
    for encoding in ("utf-8-sig", "cp932"):
        try:
            with path.open(encoding=encoding, newline="") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                return rows, list(reader.fieldnames or [])
        except UnicodeDecodeError:
            continue
    raise ValueError(f"文字コードを判定できませんでした: {path}")


def write_csv(path: Path, rows: list[dict[str, str]], fieldnames: list[str]) -> None:
    # Excelで文字化けしないよう BOM付きUTF-8 で書き出す
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 1
    src = Path(argv[1])
    out_dir = Path(argv[2]) if len(argv) > 2 else src.parent / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    rows, fieldnames = read_csv(src)
    result = process(rows, fieldnames)

    write_csv(out_dir / "cleaned.csv", result.cleaned, result.fieldnames)
    write_csv(
        out_dir / "duplicates.csv",
        result.duplicates,
        ["元の行番号", "重複先の行番号", *result.fieldnames],
    )

    need_check = sum(1 for r in result.cleaned if r[COL_CHECK])
    print(f"読み込み: {len(rows)}件")
    print(f"整形後  : {len(result.cleaned)}件 -> {out_dir / 'cleaned.csv'}")
    print(f"重複    : {len(result.duplicates)}件 -> {out_dir / 'duplicates.csv'}")
    print(f"要確認  : {need_check}件（「{COL_CHECK}」列に理由を記載）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
