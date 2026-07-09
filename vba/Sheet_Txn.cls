Option Explicit
'==============================================================
' 取引履歴シート - 変更イベント
'--------------------------------------------------------------
' 取引テーブルへ入力があったとき、その行が「日付・種類・金額」まで
' 揃っていれば整合性・残高不足を検証して警告する（二次防御）。
'==============================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject, hit As Range, r As Range
    On Error GoTo cleanExit
    If gSuppressTxnCheck Then Exit Sub        ' 一括書込中はチェックしない
    If Me.ListObjects.Count = 0 Then Exit Sub
    Set lo = Me.ListObjects(1)               ' T_Txn
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set hit = Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub

    Application.EnableEvents = False
    For Each r In hit.Rows
        CheckTxnRow r.Cells(1, 1)
    Next r

cleanExit:
    Application.EnableEvents = True
End Sub
