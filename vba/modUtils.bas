Option Explicit
'==============================================================
' modUtils - 汎用ヘルパ
'--------------------------------------------------------------
' 設定値の取得、テーブル取得、空行検索、行追加など、
' 各機能モジュールから共通利用する道具をまとめる。
'==============================================================

' 設定シートから名前で値を取得（B列=キー, C列=値）
Public Function GetConfig(ByVal key As String) As Variant
    Dim ws As Worksheet, f As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    Set f = ws.Columns("B").Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        GetConfig = ""
    Else
        GetConfig = f.Offset(0, 1).Value
    End If
End Function

' 設定が ON かどうか
Public Function ConfigIsOn(ByVal key As String) As Boolean
    ConfigIsOn = (UCase(CStr(GetConfig(key))) = "ON")
End Function

' 対象年月（yyyy-mm 文字列）
Public Function TargetYM() As String
    TargetYM = CStr(GetConfig("対象年月"))
End Function

' テーブル(ListObject)を名前で取得（全シート横断）
Public Function GetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' テーブル本体で、指定列が空の最初の行(Range)を返す。無ければ新規行を追加。
' 事前確保した空行（数式列のみ入る）を優先的に使い、末尾に散らからないようにする。
Public Function FirstEmptyRow(ByVal lo As ListObject, ByVal keyCol As Long) As Range
    Dim i As Long
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            If Len(Trim(CStr(lo.DataBodyRange.Cells(i, keyCol).Value))) = 0 Then
                Set FirstEmptyRow = lo.DataBodyRange.Rows(i)
                Exit Function
            End If
        Next i
    End If
    Set FirstEmptyRow = lo.ListRows.Add.Range
End Function

' 取引テーブルへ1件追加する（値検証は呼び出し側で済ませておくこと）。
' 書込中は gSuppressTxnCheck で Change イベントのチェックを抑制する
' （Application.EnableEvents の低レベル操作は自動化時に不安定なため使わない）。
Public Sub AddTransaction(ByVal dt As Date, ByVal kind As String, _
                          ByVal fromAcc As String, ByVal toAcc As String, _
                          ByVal amount As Double, ByVal category As String, _
                          ByVal pay As String, ByVal memo As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable(TB_TXN)
    Set r = FirstEmptyRow(lo, COL_DATE)
    gSuppressTxnCheck = True
    r.Cells(1, COL_DATE).Value = dt
    r.Cells(1, COL_KIND).Value = kind
    r.Cells(1, COL_FROM).Value = fromAcc
    r.Cells(1, COL_TO).Value = toAcc
    r.Cells(1, COL_AMOUNT).Value = amount
    r.Cells(1, COL_CATEGORY).Value = category
    r.Cells(1, COL_PAY).Value = pay
    r.Cells(1, COL_MEMO).Value = memo
    gSuppressTxnCheck = False
End Sub

' 口座の現在残高を取得（マスタ T_Account の現在残高列を参照）
Public Function AccountBalance(ByVal accName As String) As Double
    Dim lo As ListObject, i As Long
    Set lo = GetTable(TB_ACCOUNT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, 1).Value) = accName Then
            AccountBalance = Val(lo.DataBodyRange.Cells(i, 4).Value)
            Exit Function
        End If
    Next i
End Function

' 指定年月(yyyy-mm)の「支払日」から実際の日付を返す（月末超過は月末に丸める）
Public Function DateInMonth(ByVal ym As String, ByVal dayNum As Long) As Date
    Dim y As Long, m As Long, lastDay As Long
    y = CLng(Left$(ym, 4))
    m = CLng(Mid$(ym, 6, 2))
    lastDay = Day(DateSerial(y, m + 1, 0))  ' 翌月0日=当月末日
    If dayNum > lastDay Then dayNum = lastDay
    If dayNum < 1 Then dayNum = 1
    DateInMonth = DateSerial(y, m, dayNum)
End Function

' 情報メッセージ（gSilent 時は抑制）
Public Sub Info(ByVal msg As String)
    If gSilent Then Exit Sub
    MsgBox msg, vbInformation, APP_TITLE
End Sub

' 警告メッセージ（gSilent 時は抑制）
Public Sub Warn(ByVal msg As String)
    If gSilent Then Exit Sub
    MsgBox msg, vbExclamation, APP_TITLE
End Sub
