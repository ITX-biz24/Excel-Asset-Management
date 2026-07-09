# -*- coding: utf-8 -*-
"""
テーマ定義モジュール
--------------------
ダークモダンUI（Catppuccin Mocha ベース）の配色・フォント・共通スタイルを一元管理する。
色や書式を変えたいときはこのファイルだけを修正すれば全シートに反映される。
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# カラーパレット（ARGB 形式・先頭は透明度なしの "FF" を openpyxl が補完）
# ============================================================
COL = {
    "bg":       "1E1E2E",  # 基本背景（アプリの地色）
    "panel":    "181825",  # パネル（一段暗い背景）
    "card":     "313244",  # カード面
    "card2":    "45475A",  # カード面（濃淡違い）
    "input":    "494D64",  # 入力セル（編集可能を示す明るめの面）
    "text":     "CDD6F4",  # 通常テキスト（明色）
    "subtext":  "A6ADC8",  # 補助テキスト
    "muted":    "6C7086",  # 罫線・控えめな線
    "accent":   "89B4FA",  # アクセント（青）
    "green":    "A6E3A1",  # 収入・プラス
    "red":      "F38BA8",  # 支出・マイナス
    "yellow":   "F9E2AF",  # 注意・貯蓄率
    "teal":     "94E2D5",  # 振替・補助系
    "mauve":    "CBA6F7",  # 強調
    "peach":    "FAB387",  # 強調2
    "white":    "FFFFFF",
}

# グラフ用の系列カラー（口座/カテゴリの円・棒に使う）
CHART_SERIES = ["89B4FA", "A6E3A1", "F9E2AF", "F38BA8", "CBA6F7",
                "94E2D5", "FAB387", "89DCEB", "B4BEFE", "EBA0AC"]

# フォント（日本語対応のモダンフォント）
FONT_NAME = "Yu Gothic UI"

# ============================================================
# フォント生成ヘルパ
# ============================================================
def font(size=11, bold=False, color="text", italic=False):
    return Font(name=FONT_NAME, size=size, bold=bold,
                color=COL.get(color, color), italic=italic)

# ============================================================
# 塗りつぶし生成ヘルパ
# ============================================================
def fill(color):
    c = COL.get(color, color)
    return PatternFill(fill_type="solid", fgColor=c, bgColor=c)

# ============================================================
# 配置
# ============================================================
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=False)
RIGHT  = Alignment(horizontal="right",  vertical="center", wrap_text=False)
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ============================================================
# 罫線（薄いグレーの下線・カード境界など）
# ============================================================
def _side(color="muted"):
    return Side(style="thin", color=COL.get(color, color))

BORDER_BOTTOM = Border(bottom=_side("card2"))
BORDER_ALL    = Border(left=_side("card2"), right=_side("card2"),
                       top=_side("card2"), bottom=_side("card2"))

# ============================================================
# 数値書式
# ============================================================
FMT_YEN     = '¥#,##0;[Red]-¥#,##0'   # 通貨（マイナスは赤）
FMT_YEN0    = '¥#,##0'                 # 通貨（符号色なし）
FMT_PCT     = '0.0%'                   # パーセント（貯蓄率）
FMT_DATE    = 'yyyy/mm/dd'             # 日付
FMT_YM      = 'yyyy-mm'                # 年月キー
FMT_INT     = '#,##0'                  # 整数

# テーブルスタイル（openpyxl 既定スタイル名。ダーク面と馴染む中間色）
TABLE_STYLE = "TableStyleMedium15"
