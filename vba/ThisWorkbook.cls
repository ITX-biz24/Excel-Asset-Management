Option Explicit
'==============================================================
' ThisWorkbook - ブックイベント
'--------------------------------------------------------------
' 起動時にダッシュボードを表示・再計算。保存前に（設定がONなら）
' 自動バックアップを実行する。
'==============================================================

Private Sub Workbook_Open()
    On Error Resume Next
    Application.CalculateFull
    UpdateDynamicCharts
    ThisWorkbook.Worksheets(SH_DASH).Activate
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If ConfigIsOn("自動バックアップ") Then BackupWorkbook True   ' silent
    On Error GoTo 0
End Sub
