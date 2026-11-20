"""clean_csv.py のテスト。 実行: python -m unittest test_clean_csv.py"""

import unittest

from clean_csv import (
    clean_date,
    clean_email,
    clean_name,
    clean_phone,
    clean_postal,
    process,
)


class TestCleaners(unittest.TestCase):
    def test_name(self):
        self.assertEqual(clean_name("　山田　　太郎 "), "山田 太郎")

    def test_email(self):
        self.assertEqual(clean_email(" Taro@Example.COM "), ("taro@example.com", None))
        self.assertEqual(clean_email("ｔａｒｏ＠ｅｘａｍｐｌｅ．ｃｏｍ"), ("taro@example.com", None))
        self.assertEqual(clean_email("taro@example")[1], "メールアドレスの形式")
        self.assertEqual(clean_email(""), ("", None))

    def test_phone(self):
        self.assertEqual(clean_phone("09012345678"), ("090-1234-5678", None))
        self.assertEqual(clean_phone("０９０（１２３４）５６７８"), ("090-1234-5678", None))
        self.assertEqual(clean_phone("0312345678"), ("03-1234-5678", None))
        # 市外局番の桁数が決められない番号は、元の区切りを残す
        self.assertEqual(clean_phone("0466-12-3456"), ("0466-12-3456", None))
        self.assertEqual(clean_phone("0466123456"), ("0466123456", None))
        self.assertEqual(clean_phone("12345")[1], "電話番号の桁数")
        self.assertEqual(clean_phone("090-5555-666")[1], "電話番号の桁数")

    def test_postal(self):
        self.assertEqual(clean_postal("２５１００５２"), ("251-0052", None))
        self.assertEqual(clean_postal("251-0052"), ("251-0052", None))
        self.assertEqual(clean_postal("2510")[1], "郵便番号の桁数")

    def test_date(self):
        for raw in ("2026/1/5", "2026-01-05", "2026年1月5日", "20260105", "２０２６／１／５"):
            self.assertEqual(clean_date(raw), ("2026-01-05", None), raw)
        self.assertEqual(clean_date("2026/13/5")[1], "日付の形式")
        self.assertEqual(clean_date("昨日")[1], "日付の形式")


class TestProcess(unittest.TestCase):
    def test_duplicates_by_email_and_name_phone(self):
        fields = ["氏名", "メールアドレス", "電話番号", "郵便番号", "登録日"]
        rows = [
            {"氏名": "山田 太郎", "メールアドレス": "taro@example.com", "電話番号": "09012345678", "郵便番号": "", "登録日": ""},
            {"氏名": "山田太郎", "メールアドレス": "TARO@example.com", "電話番号": "", "郵便番号": "", "登録日": ""},
            {"氏名": "鈴木 花子", "メールアドレス": "", "電話番号": "080-1111-2222", "郵便番号": "", "登録日": ""},
            {"氏名": "鈴木　花子", "メールアドレス": "", "電話番号": "08011112222", "郵便番号": "", "登録日": ""},
            {"氏名": "佐藤 次郎", "メールアドレス": "", "電話番号": "", "郵便番号": "", "登録日": ""},
        ]
        result = process(rows, fields)
        self.assertEqual(len(result.cleaned), 3)
        self.assertEqual(len(result.duplicates), 2)
        self.assertEqual(result.duplicates[0]["元の行番号"], "3")
        self.assertEqual(result.duplicates[0]["重複先の行番号"], "2")
        self.assertEqual(result.duplicates[1]["元の行番号"], "5")
        self.assertEqual(result.duplicates[1]["重複先の行番号"], "4")


if __name__ == "__main__":
    unittest.main()
