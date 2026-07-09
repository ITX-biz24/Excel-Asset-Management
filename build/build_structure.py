# -*- coding: utf-8 -*-
"""
構造ビルダ（openpyxl）
======================
個人資産管理システム v2 のブック構造（シート・数式・入力規則・条件付き書式・
グラフ・名前付き範囲・テーマ）を生成し、中間 .xlsx として保存する。

VBA は本スクリプトでは注入しない（openpyxl は VBA 非対応）。
VBA 注入と .xlsm 化は build/inject_vba.py が担当する。

設計は docs/SPEC.md を参照。口座はマスタ表で可変、残高は 2 本の SUMIFS で一般化する。
"""

import os
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import theme as T

# ------------------------------------------------------------
# パス
# ------------------------------------------------------------
BUILD_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BUILD_DIR)
OUT_XLSX = os.path.join(BUILD_DIR, "_structure.xlsx")

# ------------------------------------------------------------
# シート名（modConfig.bas と一致させること）
# ------------------------------------------------------------
SH = {
    "dash":  "Dashboard",
    "txn":   "取引履歴",
    "fixed": "固定収支",
    "sub":   "サブスク",
    "report":"レポート",
    "snap":  "スナップショット",
    "master":"マスタ",
    "config":"設定",
}

# テーブル displayName（ASCII で統一。列見出しは日本語）
TB = {
    "account":  "T_Account",
    "category": "T_Category",
    "pay":      "T_Pay",
    "txn":      "T_Txn",
    "fixed":    "T_Fixed",
    "sub":      "T_Sub",
    "snap":     "T_Snap",
}

# 既定の対象年月（今日 2026-07 に合わせる）
TARGET_YM = "2026-07"

# 取引テーブルの確保行数（テーブルは入力で自動拡張するが余裕を持って確保）
TXN_ROWS = 400

# データテーブル用のダーク系スタイル（明色テーブルより地色に馴染む）
TABLE_STYLE_DARK = "TableStyleDark9"


# ============================================================
# 汎用ヘルパ
# ============================================================
def paint(ws, max_row, max_col, color="bg"):
    """指定範囲をテーマ背景色で塗り、アプリらしい地色にする。"""
    f = T.fill(color)
    fnt = T.font()
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = f
            cell.font = fnt


def put(ws, coord, value=None, *, font_kw=None, fill_c=None, align=None,
        fmt=None, border=None):
    """1セルへ値とスタイルをまとめて設定する小ヘルパ。"""
    cell = ws[coord]
    if value is not None:
        cell.value = value
    if font_kw is not None:
        cell.font = T.font(**font_kw)
    if fill_c is not None:
        cell.fill = T.fill(fill_c)
    if align is not None:
        cell.alignment = align
    if fmt is not None:
        cell.number_format = fmt
    if border is not None:
        cell.border = border
    return cell


def title_block(ws, text, sub=""):
    """各シート上部の見出し（アイコン＋タイトル＋サブ）を描く。"""
    put(ws, "B2", text, font_kw=dict(size=20, bold=True, color="text"),
        align=T.LEFT)
    if sub:
        put(ws, "B3", sub, font_kw=dict(size=10, color="subtext"), align=T.LEFT)


def add_table(ws, name, ref, style=T.TABLE_STYLE):
    """テーブル（ListObject）を追加する。"""
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name=style, showRowStripes=True, showColumnStripes=False,
        showFirstColumn=False, showLastColumn=False)
    ws.add_table(tab)
    return tab


def add_name(wb, name, ref):
    """名前付き範囲を追加（テーブル列参照や絶対参照を受け付ける）。"""
    wb.defined_names.add(DefinedName(name=name, attr_text=ref))


def set_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def hide_grid(ws):
    ws.sheet_view.showGridLines = False


# ============================================================
# マスタシート
# ============================================================
def build_master(wb):
    ws = wb.create_sheet(SH["master"])
    hide_grid(ws)
    paint(ws, 60, 18)
    ws.sheet_properties.tabColor = T.COL["muted"]
    title_block(ws, "マスタ", "口座・カテゴリ・支払方法・種類の一元管理（ここを編集すれば全シートに反映）")
    set_widths(ws, {"A": 2, "B": 16, "C": 14, "D": 10, "E": 16,
                    "F": 2, "G": 16, "H": 12, "I": 2, "J": 16})

    # --- 口座テーブル（残高計算の中核。行を足すだけで口座を拡張できる）---
    hdr_row = 5
    acc_headers = ["口座名", "初期残高", "表示順", "現在残高"]
    for i, h in enumerate(acc_headers):
        put(ws, f"{get_column_letter(2+i)}{hdr_row}", h,
            font_kw=dict(bold=True, color="bg"), fill_c="accent", align=T.CENTER)
    accounts = [("現金", 100, 1), ("銀行", 100, 2), ("QR", 100, 3)]
    for j, (nm, init, order) in enumerate(accounts):
        r = hdr_row + 1 + j
        put(ws, f"B{r}", nm, fill_c="input", align=T.LEFT)
        put(ws, f"C{r}", init, fill_c="input", align=T.RIGHT, fmt=T.FMT_YEN0)
        put(ws, f"D{r}", order, fill_c="input", align=T.CENTER)
        # 現在残高 = 初期残高 + 入金先合計 - 出金元合計（口座に依存しない一般式）
        # 行内参照は openpyxl では [@..] が壊れるため明示セル参照を用いる（列全体参照は構造化のまま）
        put(ws, f"E{r}",
            f"=C{r}+SUMIFS({TB['txn']}[金額],{TB['txn']}[入金先],B{r})"
            f"-SUMIFS({TB['txn']}[金額],{TB['txn']}[出金元],B{r})",
            align=T.RIGHT, fmt=T.FMT_YEN)
    add_table(ws, TB["account"], f"B{hdr_row}:E{hdr_row+len(accounts)}")
    put(ws, f"B{hdr_row-1}", "■ 口座", font_kw=dict(bold=True, color="teal"), align=T.LEFT)

    # --- カテゴリテーブル（区分で収入/支出を分類）---
    cat_hdr = 5
    put(ws, f"G{cat_hdr-1}", "■ カテゴリ", font_kw=dict(bold=True, color="teal"), align=T.LEFT)
    put(ws, f"G{cat_hdr}", "カテゴリ", font_kw=dict(bold=True, color="bg"),
        fill_c="accent", align=T.CENTER)
    put(ws, f"H{cat_hdr}", "区分", font_kw=dict(bold=True, color="bg"),
        fill_c="accent", align=T.CENTER)
    categories = [
        ("給与", "収入"), ("賞与", "収入"), ("副業", "収入"), ("その他収入", "収入"),
        ("食費", "支出"), ("日用品", "支出"), ("住居", "支出"), ("水道光熱", "支出"),
        ("通信", "支出"), ("交通", "支出"), ("娯楽", "支出"), ("交際", "支出"),
        ("医療", "支出"), ("教育", "支出"), ("保険", "支出"), ("サブスク", "支出"),
        ("その他支出", "支出"),
    ]
    for j, (nm, kind) in enumerate(categories):
        r = cat_hdr + 1 + j
        put(ws, f"G{r}", nm, fill_c="input", align=T.LEFT)
        put(ws, f"H{r}", kind, fill_c="input", align=T.CENTER)
    add_table(ws, TB["category"], f"G{cat_hdr}:H{cat_hdr+len(categories)}")

    # --- 支払方法テーブル ---
    pay_hdr = 5
    put(ws, f"J{pay_hdr-1}", "■ 支払方法", font_kw=dict(bold=True, color="teal"), align=T.LEFT)
    put(ws, f"J{pay_hdr}", "支払方法", font_kw=dict(bold=True, color="bg"),
        fill_c="accent", align=T.CENTER)
    pays = ["現金", "クレジット", "銀行振込", "口座引落", "QR決済", "デビット", "電子マネー"]
    for j, nm in enumerate(pays):
        put(ws, f"J{pay_hdr+1+j}", nm, fill_c="input", align=T.LEFT)
    add_table(ws, TB["pay"], f"J{pay_hdr}:J{pay_hdr+len(pays)}")

    # --- 種類リスト（収入/支出/振替）: 静的名前付き範囲 ---
    kinds_col = "L"
    put(ws, f"{kinds_col}{pay_hdr-1}", "■ 種類", font_kw=dict(bold=True, color="teal"), align=T.LEFT)
    kinds = ["収入", "支出", "振替"]
    for j, k in enumerate(kinds):
        put(ws, f"{kinds_col}{pay_hdr+j}", k, fill_c="card2", align=T.CENTER)
    ws.column_dimensions[kinds_col].width = 10

    # 名前付き範囲（テーブル列参照は入力規則の元として利用）
    add_name(wb, "口座リスト",   f"{SH['master']}!{TB['account']}[口座名]")
    add_name(wb, "カテゴリリスト", f"{SH['master']}!{TB['category']}[カテゴリ]")
    add_name(wb, "支払方法リスト", f"{SH['master']}!{TB['pay']}[支払方法]")
    add_name(wb, "種類リスト",
             f"{SH['master']}!${kinds_col}${pay_hdr}:${kinds_col}${pay_hdr+len(kinds)-1}")
    return ws


# ============================================================
# 取引履歴シート（全集計の唯一の真実）
# ============================================================
# 列定義: (見出し, 列文字, 幅, 数値書式, 計算列か)
TXN_COLS = [
    ("日付",     "B", 12, T.FMT_DATE, False),
    ("種類",     "C", 8,  None,       False),
    ("出金元",   "D", 10, None,       False),
    ("入金先",   "E", 10, None,       False),
    ("金額",     "F", 13, T.FMT_YEN0, False),
    ("カテゴリ", "G", 11, None,       False),
    ("支払方法", "H", 11, None,       False),
    ("メモ",     "I", 26, None,       False),
    ("年月",     "J", 9,  T.FMT_YM,   True),
]

# デモデータ（利用者が削除して使い始める想定。2026-07）
TXN_DEMO = [
    # 日付,       種類,   出金元, 入金先, 金額,   カテゴリ,     支払方法,   メモ
    ("2026-07-05", "収入", "",     "銀行", 250000, "給与",       "銀行振込", "7月給与"),
    ("2026-07-02", "収入", "",     "QR",   8000,   "副業",       "QR決済",   "ポイント収入"),
    ("2026-07-01", "支出", "銀行", "",     80000,  "住居",       "口座引落", "家賃"),
    ("2026-07-03", "支出", "現金", "",     3200,   "食費",       "現金",     "スーパー"),
    ("2026-07-04", "支出", "QR",   "",     1500,   "食費",       "QR決済",   "コンビニ"),
    ("2026-07-06", "支出", "銀行", "",     6000,   "通信",       "口座引落", "携帯料金"),
    ("2026-07-06", "支出", "QR",   "",     4800,   "娯楽",       "QR決済",   "映画"),
    ("2026-07-08", "支出", "現金", "",     2500,   "食費",       "現金",     "ランチ"),
    ("2026-07-08", "支出", "銀行", "",     12000,  "交際",       "クレジット", "飲み会"),
    ("2026-07-07", "振替", "銀行", "現金", 30000,  "",           "銀行振込", "ATM出金"),
]


def build_txn(wb):
    ws = wb.create_sheet(SH["txn"])
    hide_grid(ws)
    paint(ws, TXN_ROWS + 10, 12)
    ws.sheet_properties.tabColor = T.COL["accent"]
    title_block(ws, "取引履歴", "すべての残高・集計・グラフの元データ。1行=1取引")
    set_widths(ws, {"A": 2})
    for name, col, w, fmt, _ in TXN_COLS:
        ws.column_dimensions[col].width = w

    hdr = 6
    # ヘッダ
    for name, col, w, fmt, _ in TXN_COLS:
        put(ws, f"{col}{hdr}", name, font_kw=dict(bold=True, color="text"), align=T.CENTER)
    # データ行（デモ＋空行）に数式・書式を敷設
    first, last = hdr + 1, hdr + TXN_ROWS
    for r in range(first, last + 1):
        idx = r - first
        demo = TXN_DEMO[idx] if idx < len(TXN_DEMO) else None
        for ci, (name, col, w, fmt, calc) in enumerate(TXN_COLS):
            cell = ws[f"{col}{r}"]
            if calc:  # 年月（計算列。行内参照は明示セル参照で堅牢化）
                cell.value = f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))'
                cell.fill = T.fill("card")  # 計算列は控えめな面で自動生成を示す
                cell.font = T.font(color="subtext")
                cell.alignment = T.CENTER
            elif demo is not None:
                v = demo[ci]
                if name == "日付" and v:
                    from datetime import datetime
                    cell.value = datetime.strptime(v, "%Y-%m-%d")
                else:
                    cell.value = v if v != "" else None
            if fmt:
                cell.number_format = fmt
            if name in ("金額",):
                cell.alignment = T.RIGHT
            elif name in ("種類", "出金元", "入金先", "支払方法"):
                cell.alignment = T.CENTER
            elif name in ("日付",):
                cell.alignment = T.CENTER

    ref = f"B{hdr}:J{last}"
    add_table(ws, TB["txn"], ref, style=TABLE_STYLE_DARK)
    ws.freeze_panes = f"B{hdr+1}"

    # --- 入力規則（一次防御。詳細な整合性チェックは VBA が担う）---
    data_rng = f"C{first}:C{last}"
    dv_kind = DataValidation(type="list", formula1="種類リスト", allow_blank=True,
                             showErrorMessage=True, errorTitle="種類",
                             error="収入 / 支出 / 振替 から選択してください")
    dv_kind.add(data_rng); ws.add_data_validation(dv_kind)

    dv_from = DataValidation(type="list", formula1="口座リスト", allow_blank=True)
    dv_from.add(f"D{first}:D{last}"); ws.add_data_validation(dv_from)
    dv_to = DataValidation(type="list", formula1="口座リスト", allow_blank=True)
    dv_to.add(f"E{first}:E{last}"); ws.add_data_validation(dv_to)

    dv_amt = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                            allow_blank=True, showErrorMessage=True,
                            errorTitle="金額", error="金額は0より大きい値を入力してください")
    dv_amt.add(f"F{first}:F{last}"); ws.add_data_validation(dv_amt)

    dv_cat = DataValidation(type="list", formula1="カテゴリリスト", allow_blank=True)
    dv_cat.add(f"G{first}:G{last}"); ws.add_data_validation(dv_cat)
    dv_pay = DataValidation(type="list", formula1="支払方法リスト", allow_blank=True)
    dv_pay.add(f"H{first}:H{last}"); ws.add_data_validation(dv_pay)

    dv_date = DataValidation(type="date", operator="greaterThanOrEqual",
                             formula1="DATE(2000,1,1)", allow_blank=True,
                             showErrorMessage=True, errorTitle="日付",
                             error="正しい日付を入力してください")
    dv_date.add(f"B{first}:B{last}"); ws.add_data_validation(dv_date)

    # --- 条件付き書式（種類で金額を色分け）---
    ws.conditional_formatting.add(
        f"C{first}:C{last}",
        FormulaRule(formula=[f'$C{first}="収入"'], font=Font(color=T.COL["green"], bold=True)))
    ws.conditional_formatting.add(
        f"C{first}:C{last}",
        FormulaRule(formula=[f'$C{first}="支出"'], font=Font(color=T.COL["red"], bold=True)))
    ws.conditional_formatting.add(
        f"C{first}:C{last}",
        FormulaRule(formula=[f'$C{first}="振替"'], font=Font(color=T.COL["teal"], bold=True)))
    ws.conditional_formatting.add(
        f"F{first}:F{last}",
        FormulaRule(formula=[f'$C{first}="収入"'], font=Font(color=T.COL["green"])))
    ws.conditional_formatting.add(
        f"F{first}:F{last}",
        FormulaRule(formula=[f'$C{first}="支出"'], font=Font(color=T.COL["red"])))
    return ws


# ============================================================
# 設定シート
# ============================================================
def build_config(wb):
    ws = wb.create_sheet(SH["config"])
    hide_grid(ws)
    paint(ws, 40, 12)
    ws.sheet_properties.tabColor = T.COL["muted"]
    title_block(ws, "設定", "システム全体の基準値。数式・VBA が参照する定数の置き場")
    set_widths(ws, {"A": 2, "B": 22, "C": 20, "D": 40})

    rows = [
        ("対象年月",       TARGET_YM, "ダッシュボード・当月集計の基準（yyyy-mm）"),
        ("集計開始年月",   "2026-01", "レポート月次サマリの開始月（yyyy-mm）"),
        ("残高不足チェック", "ON",      "出金時に残高不足を拒否するか（ON/OFF）"),
        ("自動バックアップ", "OFF",     "保存前に自動バックアップするか（ON/OFF）"),
        ("バックアップ先",  r"backup", "バックアップ保存フォルダ（相対 or 絶対パス）"),
        ("アプリ名",       "個人資産管理システム", ""),
        ("バージョン",     "v2.0",    ""),
    ]
    start = 5
    put(ws, f"B{start-1}", "■ 基本設定", font_kw=dict(bold=True, color="teal"), align=T.LEFT)
    for i, (k, v, note) in enumerate(rows):
        r = start + i
        put(ws, f"B{r}", k, font_kw=dict(bold=True, color="text"),
            fill_c="card", align=T.LEFT)
        put(ws, f"C{r}", v, fill_c="input", align=T.CENTER)
        put(ws, f"D{r}", note, font_kw=dict(size=9, color="subtext"), align=T.LEFT)

    # 名前付き範囲（数式から参照）
    add_name(wb, "対象年月",     f"{SH['config']}!$C${start}")
    add_name(wb, "集計開始年月", f"{SH['config']}!$C${start+1}")
    add_name(wb, "残高不足チェック", f"{SH['config']}!$C${start+2}")
    add_name(wb, "自動バックアップ", f"{SH['config']}!$C${start+3}")
    add_name(wb, "バックアップ先", f"{SH['config']}!$C${start+4}")
    return ws


# ============================================================
# メイン
# ============================================================
def build():
    wb = Workbook()
    # 既定シートは最後に順序調整するため一旦保持
    default = wb.active

    build_txn(wb)
    build_master(wb)
    build_config(wb)

    # 既定シートを削除（実シートを作り終えてから）
    wb.remove(default)

    # シート表示順（左から）: ダッシュボード→取引→固定→サブスク→レポート→スナップ→マスタ→設定
    order = [SH["dash"], SH["txn"], SH["fixed"], SH["sub"],
             SH["report"], SH["snap"], SH["master"], SH["config"]]
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

    wb.save(OUT_XLSX)
    print(f"[build_structure] saved: {OUT_XLSX}")
    print(f"[build_structure] sheets: {wb.sheetnames}")


if __name__ == "__main__":
    build()
