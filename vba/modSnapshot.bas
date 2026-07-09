Option Explicit
'==============================================================
' modSnapshot - 資産スナップショット
'--------------------------------------------------------------
' 対象年月の資産残高を T_Snap へ「値」で保存する（履歴改変を防ぐ）。
' 資産推移グラフの元データになる。既存の同月・値行があれば上書き。
' 先頭のライブ行（年月が数式）は当月表示用なので上書きしない。
'==============================================================

' 総資産（全口座の現在残高合計）
Public Function TotalAssets() As Double
    Dim lo As ListObject, i As Long, s As Double
    Set lo = GetTable(TB_ACCOUNT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        s = s + Val(lo.DataBodyRange.Cells(i, 4).Value)
    Next i
    TotalAssets = s
End Function

' 対象年月（省略時=設定の対象年月）のスナップショットを値で保存
Public Sub SaveMonthEndSnapshot(Optional ByVal ym As String = "")
    Dim lo As ListObject, i As Long, target As Range
    If Len(ym) = 0 Then ym = TargetYM()
    If Len(ym) = 0 Then Warn "対象年月が未設定です。": Exit Sub
    Set lo = GetTable(TB_SNAP)
    If lo Is Nothing Then Exit Sub

    ' 同月の「値行」を探す（先頭のライブ数式行は除外）
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            Dim c As Range
            Set c = lo.DataBodyRange.Cells(i, 1)
            If Not c.HasFormula Then
                If CStr(c.Value) = ym Then Set target = lo.DataBodyRange.Rows(i): Exit For
            End If
        Next i
    End If
    If target Is Nothing Then Set target = lo.ListRows.Add.Range

    target.Cells(1, 1).Value = ym
    target.Cells(1, 2).Value = TotalAssets()
    target.Cells(1, 3).Value = AccountBalance("現金")
    target.Cells(1, 4).Value = AccountBalance("銀行")
    target.Cells(1, 5).Value = AccountBalance("QR")

    Info ym & " の資産スナップショットを保存しました（総資産 " & _
         Format(TotalAssets(), "#,##0") & " 円）。"
End Sub
