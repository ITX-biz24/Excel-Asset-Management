Option Explicit
'==============================================================
' modUpdate - ワンクリック更新
'--------------------------------------------------------------
' 全再計算に加え、可変長テーブル（口座・スナップ）に合わせて
' 動的グラフ（資産割合・資産推移）の参照範囲を追従させる。
'==============================================================

Public Sub RefreshAll()
    Application.ScreenUpdating = False
    Application.CalculateFull
    UpdateDynamicCharts
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(SH_DASH).Activate
    Info "最新の状態に更新しました。"
End Sub

' 口座数・スナップ件数の増減にグラフを追従させる
Public Sub UpdateDynamicCharts()
    Dim dash As Worksheet, co As ChartObject
    Dim accSheet As Worksheet, snapSheet As Worksheet
    Dim accLo As ListObject, snapLo As ListObject
    Dim accFirst As Long, accLast As Long, snapFirst As Long, snapLast As Long
    Dim title As String, s As Series

    Set dash = ThisWorkbook.Worksheets(SH_DASH)
    Set accLo = GetTable(TB_ACCOUNT)
    Set snapLo = GetTable(TB_SNAP)
    If accLo Is Nothing Or snapLo Is Nothing Then Exit Sub
    Set accSheet = accLo.Parent
    Set snapSheet = snapLo.Parent

    accFirst = accLo.DataBodyRange.Row
    accLast = accFirst + accLo.ListRows.Count - 1
    snapFirst = snapLo.DataBodyRange.Row
    snapLast = snapFirst + snapLo.ListRows.Count - 1

    For Each co In dash.ChartObjects
        title = ""
        On Error Resume Next
        title = co.Chart.ChartTitle.Text
        On Error GoTo 0

        If InStr(title, "資産割合") > 0 Then
            If co.Chart.SeriesCollection.Count >= 1 Then
                Set s = co.Chart.SeriesCollection(1)
                s.Values = accSheet.Range("E" & accFirst & ":E" & accLast)
                s.XValues = accSheet.Range("B" & accFirst & ":B" & accLast)
            End If
        ElseIf InStr(title, "資産推移") > 0 Then
            If co.Chart.SeriesCollection.Count >= 1 Then
                Set s = co.Chart.SeriesCollection(1)
                s.Values = snapSheet.Range("C" & snapFirst & ":C" & snapLast)
                s.XValues = snapSheet.Range("B" & snapFirst & ":B" & snapLast)
            End If
        End If
    Next co
End Sub

' フォームを開く（ボタンから呼ぶ）
Public Sub ShowTransactionForm()
    frmTransaction.Show
End Sub
