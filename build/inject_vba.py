# -*- coding: utf-8 -*-
"""
VBA 注入・仕上げビルダ（win32com）
==================================
build_structure.py が生成した中間 .xlsx を Excel COM で開き、
  1. vba/ の標準モジュール・クラスを CodeModule.AddFromString で注入
  2. ThisWorkbook / 取引履歴シートのイベントコードを注入
  3. 取引入力フォーム frmTransaction を生成（コントロール＋コード）
  4. 各シートにマクロ割当ボタンを配置
  5. グラフをダークテーマ化
  6. マクロ有効ブック(.xlsm) として保存
を行う。

前提: 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼」(AccessVBOM=1)。
"""

import os
import sys
import io
import win32com.client as win32

BUILD_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BUILD_DIR)
VBA_DIR = os.path.join(REPO_DIR, "vba")
SRC_XLSX = os.path.join(BUILD_DIR, "_structure.xlsx")
OUT_XLSM = os.path.join(REPO_DIR, "Excel-Asset-Management.xlsm")

# COM 定数
VBEXT_CT_STDMODULE = 1
VBEXT_CT_MSFORM = 3
XL_XLSM = 52          # xlOpenXMLWorkbookMacroEnabled
MSO_FALSE = 0
MSO_TRUE = -1
XL_CATEGORY = 1
XL_VALUE = 2

SH_TXN = "取引履歴"
SH_DASH = "Dashboard"
SH_FIXED = "固定収支"
SH_SUB = "サブスク"

# 注入する標準モジュール／クラス（ファイル名 → コンポーネント名）
STD_MODULES = [
    "modConfig", "modUtils", "modInputCheck", "modFixed",
    "modSnapshot", "modMonth", "modBackup", "modProtect", "modUpdate",
]

# テーマ色
C_CHART_BG = "313244"
C_TEXT = "CDD6F4"
C_GRID = "45475A"
C_SERIES = ["89B4FA", "A6E3A1", "F9E2AF", "F38BA8", "CBA6F7",
            "94E2D5", "FAB387", "89DCEB", "B4BEFE", "EBA0AC"]


def xlrgb(hex6):
    """'RRGGBB' → Excel の .RGB 用 Long（R + G*256 + B*65536）。"""
    r = int(hex6[0:2], 16); g = int(hex6[2:4], 16); b = int(hex6[4:6], 16)
    return r + g * 256 + b * 65536


def read_vba(name):
    # VBE の AddFromString は行末 CRLF を前提とする。LF のままだと
    # 行継続文字 `_` の直後で改行が誤解釈され、壊れたプロシージャになる。
    path = os.path.join(VBA_DIR, name)
    txt = io.open(path, encoding="utf-8").read()
    txt = txt.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    return txt


def inject_std_modules(wb):
    """標準モジュールを新規コンポーネントとして追加しコード注入。"""
    for name in STD_MODULES:
        comp = wb.VBProject.VBComponents.Add(VBEXT_CT_STDMODULE)
        comp.Name = name
        comp.CodeModule.AddFromString(read_vba(name + ".bas"))
    print(f"  標準モジュール {len(STD_MODULES)} 個を注入")


def inject_document_code(wb):
    """ThisWorkbook と 取引履歴シートのイベントコードを注入。"""
    # ThisWorkbook
    tw = wb.VBProject.VBComponents("ThisWorkbook")
    tw.CodeModule.AddFromString(read_vba("ThisWorkbook.cls"))

    # 取引履歴シート（CodeName でコンポーネント特定）
    ws = wb.Worksheets(SH_TXN)
    comp = wb.VBProject.VBComponents(ws.CodeName)
    comp.CodeModule.AddFromString(read_vba("Sheet_Txn.cls"))
    print("  ThisWorkbook / 取引履歴 のイベントコードを注入")


def build_userform(wb):
    """取引入力フォームを生成（コントロール＋コード）。"""
    comp = wb.VBProject.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmTransaction"
    for prop, val in (("Caption", "取引を追加"), ("Width", 300), ("Height", 340)):
        try:
            comp.Properties(prop).Value = val
        except Exception:
            pass
    frm = comp.Designer

    def add(progid, name, left, top, width, height, caption=None):
        c = frm.Controls.Add(progid, name, True)
        c.Left, c.Top, c.Width, c.Height = left, top, width, height
        if caption is not None:
            c.Caption = caption
        return c

    # ラベル＋入力を縦に並べる
    fields = [
        ("日付",     "txtDate",     "Forms.TextBox.1"),
        ("種類",     "cboKind",     "Forms.ComboBox.1"),
        ("出金元",   "cboFrom",     "Forms.ComboBox.1"),
        ("入金先",   "cboTo",       "Forms.ComboBox.1"),
        ("金額",     "txtAmount",   "Forms.TextBox.1"),
        ("カテゴリ", "cboCategory", "Forms.ComboBox.1"),
        ("支払方法", "cboPay",      "Forms.ComboBox.1"),
        ("メモ",     "txtMemo",     "Forms.TextBox.1"),
    ]
    y = 12
    for label, name, progid in fields:
        add("Forms.Label.1", "lbl_" + name, 12, y + 3, 70, 18, caption=label)
        add(progid, name, 90, y, 180, 20)
        y += 28

    add("Forms.CommandButton.1", "btnAdd", 60, y + 6, 90, 26, caption="追加")
    add("Forms.CommandButton.1", "btnClose", 165, y + 6, 90, 26, caption="閉じる")

    comp.CodeModule.AddFromString(read_vba("frmTransaction.cls"))
    print("  取引入力フォーム frmTransaction を生成")


def add_button(ws, cell, width, height, caption, macro, dy=1):
    rng = ws.Range(cell)
    b = ws.Buttons().Add(rng.Left, rng.Top + dy, width, height)
    b.Caption = caption
    b.OnAction = macro
    b.Font.Size = 10
    return b


def add_buttons(wb):
    """各シートにマクロ割当ボタンを配置。"""
    fixed = wb.Worksheets(SH_FIXED)
    add_button(fixed, "B4", 120, 22, "▶ 今月へ反映", "ReflectFixedItems")
    sub = wb.Worksheets(SH_SUB)
    add_button(sub, "B4", 120, 22, "▶ 今月へ反映", "ReflectSubscriptions")

    dash = wb.Worksheets(SH_DASH)
    base = dash.Range("F2")
    specs = [
        ("🔄 更新", "RefreshAll"),
        ("＋ 取引を追加", "ShowTransactionForm"),
        ("▶ 翌月へ", "RolloverToNextMonth"),
        ("💾 バックアップ", "BackupWorkbook"),
        ("📸 スナップ保存", "SaveMonthEndSnapshot"),
        ("🔒 数式保護", "ProtectFormulas"),
    ]
    bw, bh, gx, gy = 118, 24, 122, 28
    for i, (cap, macro) in enumerate(specs):
        col = i % 3
        row = i // 3
        b = dash.Buttons().Add(base.Left + col * gx, base.Top + row * gy, bw, bh)
        b.Caption = cap
        b.OnAction = macro
        b.Font.Size = 10
    print("  ボタンを配置（固定/サブスク/ダッシュボード）")


def theme_charts(wb):
    """ダッシュボードのグラフをダークテーマ化。"""
    dash = wb.Worksheets(SH_DASH)
    for co in dash.ChartObjects():
        try:
            ch = co.Chart
            ca = ch.ChartArea
            ca.Format.Fill.ForeColor.RGB = xlrgb(C_CHART_BG)
            ca.Format.Line.Visible = MSO_FALSE
            try:
                ch.PlotArea.Format.Fill.Visible = MSO_FALSE
            except Exception:
                pass
            ca.Font.Color = xlrgb(C_TEXT)
            ca.Font.Size = 9
            if ch.HasTitle:
                ch.ChartTitle.Font.Color = xlrgb(C_TEXT)
                ch.ChartTitle.Font.Size = 12
                ch.ChartTitle.Font.Bold = True
            # 軸
            for ax_type in (XL_CATEGORY, XL_VALUE):
                try:
                    ax = ch.Axes(ax_type)
                    ax.TickLabels.Font.Color = xlrgb(C_TEXT)
                    ax.Format.Line.ForeColor.RGB = xlrgb(C_GRID)
                    try:
                        ax.MajorGridlines.Border.Color = xlrgb(C_GRID)
                    except Exception:
                        pass
                except Exception:
                    pass
            if ch.HasLegend:
                ch.Legend.Font.Color = xlrgb(C_TEXT)
            # 系列色
            sc = ch.SeriesCollection()
            title = ch.ChartTitle.Text if ch.HasTitle else ""
            if "資産割合" in title:
                s = sc(1)
                for i in range(1, s.Points().Count + 1):
                    s.Points(i).Format.Fill.ForeColor.RGB = xlrgb(C_SERIES[(i - 1) % len(C_SERIES)])
            else:
                for si in range(1, sc.Count + 1):
                    col = xlrgb(C_SERIES[(si - 1) % len(C_SERIES)])
                    try:
                        sc(si).Format.Fill.ForeColor.RGB = col
                        sc(si).Format.Line.ForeColor.RGB = col
                    except Exception:
                        pass
        except Exception as e:
            print(f"  (グラフ書式スキップ: {e})")
    print("  グラフをダークテーマ化")


def main():
    if not os.path.exists(SRC_XLSX):
        print(f"ERROR: 中間ファイルがありません: {SRC_XLSX}")
        print("  先に build_structure.py を実行してください。")
        return 1
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    rc = 0
    try:
        # VBOM アクセス確認
        wb = xl.Workbooks.Open(SRC_XLSX)
        try:
            _ = wb.VBProject.VBComponents.Count
        except Exception:
            print("ERROR: VBAプロジェクトへアクセスできません（AccessVBOM を有効化してください）。")
            wb.Close(False)
            return 2

        print("[inject_vba] VBA を注入します")
        inject_std_modules(wb)
        inject_document_code(wb)
        build_userform(wb)
        add_buttons(wb)
        xl.CalculateFullRebuild()
        theme_charts(wb)

        # .xlsm として保存
        if os.path.exists(OUT_XLSM):
            os.remove(OUT_XLSM)
        wb.SaveAs(OUT_XLSM, FileFormat=XL_XLSM)
        wb.Close(True)
        print(f"[inject_vba] 保存完了: {OUT_XLSM}")
    except Exception as e:
        print(f"[inject_vba] ERROR: {e}")
        rc = 1
    finally:
        xl.Quit()
    return rc


if __name__ == "__main__":
    sys.exit(main())
