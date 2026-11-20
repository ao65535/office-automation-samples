"""VBAサンプル用の売上データ（sample_sales.xlsx）を作り、正解の集計値を表示する。

実行: python make_sample_data.py
必要: pip install openpyxl
"""

import random
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

STORES = ["東京店", "大阪店", "名古屋店"]
PRODUCTS = {"コーヒー豆 200g": 1200, "ドリップバッグ 10袋": 980, "マグカップ": 1800, "ギフトセット": 3500}
START, DAYS = date(2026, 7, 1), 92  # 2026年7月〜9月


def main() -> None:
    rng = random.Random(42)  # 毎回同じデータになるよう固定
    wb = Workbook()
    wb.remove(wb.active)
    expected = defaultdict(int)  # (月, 店) -> 金額
    expected_products = defaultdict(lambda: [0, 0])

    for store in STORES:
        ws = wb.create_sheet(store)
        ws.append(["日付", "商品", "数量", "金額"])
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
        for _ in range(rng.randint(40, 60)):
            d = START + timedelta(days=rng.randrange(DAYS))
            product = rng.choice(list(PRODUCTS))
            qty = rng.randint(1, 5)
            amount = PRODUCTS[product] * qty
            ws.append([d, product, qty, amount])
            expected[(d.strftime("%Y-%m"), store)] += amount
            expected_products[product][0] += qty
            expected_products[product][1] += amount
        # 実際のデータにありがちな「集計できない行」をわざと入れる（マクロはスキップする）
        ws.append(["未定", "マグカップ", 1, 1800])
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 20
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
            row[0].number_format = "yyyy/mm/dd"
        for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
            row[0].number_format = "#,##0"

    out = Path(__file__).with_name("sample_sales.xlsx")
    wb.save(out)
    print(f"作成: {out}")

    print("\n正解の集計（月 × 店舗）")
    months = sorted({m for m, _ in expected})
    print("月," + ",".join(STORES) + ",合計")
    for m in months:
        vals = [expected[(m, s)] for s in STORES]
        print(m + "," + ",".join(map(str, vals)) + f",{sum(vals)}")
    totals = [sum(expected[(m, s)] for m in months) for s in STORES]
    print("合計," + ",".join(map(str, totals)) + f",{sum(totals)}")

    print("\n正解の集計（商品別・金額の大きい順）")
    for p, (q, a) in sorted(expected_products.items(), key=lambda x: -x[1][1]):
        print(f"{p},{q},{a}")


if __name__ == "__main__":
    main()
