Option Explicit
'==============================================================
' modProtect - 数式セルの保護
'--------------------------------------------------------------
' 数式セルだけをロックしてシートを保護する。入力セルは編集可のまま。
' UserInterfaceOnly でVBAからの書き込みは許可される。
' 手動トグル（ボタン）で ON/OFF する想定。
'==============================================================

Public Sub ProtectFormulas()
    Dim ws As Worksheet, fcells As Range
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect
        ws.Cells.Locked = False            ' まず全解除（入力セルは編集可に）
        Set fcells = Nothing
        Set fcells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Not fcells Is Nothing Then fcells.Locked = True   ' 数式だけロック
        ws.Protect Password:="", UserInterfaceOnly:=True, DrawingObjects:=False, _
                   Contents:=True, Scenarios:=False, _
                   AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True, AllowSorting:=True, _
                   AllowFiltering:=True, AllowUsingPivotTables:=True
        On Error GoTo 0
    Next ws
    Application.ScreenUpdating = True
    Info "数式セルを保護しました。入力セルはこれまで通り編集できます。"
End Sub

Public Sub UnprotectAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
    Next ws
    Info "すべてのシートの保護を解除しました。"
End Sub
