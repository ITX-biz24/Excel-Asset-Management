Option Explicit
'==============================================================
' modConfig - システム全体で使う定数の一元管理
'--------------------------------------------------------------
' シート名・テーブル名・列番号・種類名などをここに集約する。
' レイアウトを変えたらこの定数だけ直せばよい（保守性のため）。
'==============================================================

' --- シート名（build_structure.py の SH と一致させること）---
Public Const SH_DASH As String = "Dashboard"
Public Const SH_TXN As String = "取引履歴"
Public Const SH_FIXED As String = "固定収支"
Public Const SH_SUB As String = "サブスク"
Public Const SH_REPORT As String = "レポート"
Public Const SH_SNAP As String = "スナップショット"
Public Const SH_MASTER As String = "マスタ"
Public Const SH_CONFIG As String = "設定"

' --- テーブル名（ListObject displayName）---
Public Const TB_TXN As String = "T_Txn"
Public Const TB_ACCOUNT As String = "T_Account"
Public Const TB_FIXED As String = "T_Fixed"
Public Const TB_SUB As String = "T_Sub"
Public Const TB_SNAP As String = "T_Snap"

' --- 取引テーブルの列番号（テーブル内の相対列・1始まり）---
Public Const COL_DATE As Long = 1
Public Const COL_KIND As Long = 2
Public Const COL_FROM As Long = 3
Public Const COL_TO As Long = 4
Public Const COL_AMOUNT As Long = 5
Public Const COL_CATEGORY As Long = 6
Public Const COL_PAY As Long = 7
Public Const COL_MEMO As Long = 8
Public Const COL_YM As Long = 9

' --- 種類 ---
Public Const KIND_INCOME As String = "収入"
Public Const KIND_EXPENSE As String = "支出"
Public Const KIND_TRANSFER As String = "振替"

' --- 反映済みマーク（固定収支・サブスクの二重反映防止キー）---
Public Const MARK_FIXED As String = "[固定]"
Public Const MARK_SUB As String = "[サブスク]"

' アプリ名（メッセージボックスのタイトル）
Public Const APP_TITLE As String = "個人資産管理システム"

' プログラムからの一括書込中は取引履歴の入力チェックを抑制するフラグ。
' Application.EnableEvents の低レベル操作は自動化実行時に不安定なため、
' このフラグで Worksheet_Change のチェックだけを黙らせる。
Public gSuppressTxnCheck As Boolean

' メッセージ抑制フラグ。自動テスト・自動バックアップ等、対話ダイアログを
' 出したくない文脈で True にすると Info/Warn の MsgBox を出さない。
Public gSilent As Boolean
