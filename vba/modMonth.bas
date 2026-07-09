Option Explicit
'==============================================================
' modMonth - 月次ロールオーバー（翌月へ）
'--------------------------------------------------------------
' 当月をスナップショットとして確定し、対象年月を翌月へ進める。
' 希望すれば固定収支・サブスクを新しい月へ自動反映する。
'==============================================================

' yyyy-mm を1か月進める
Private Function NextYM(ByVal ym As String) As String
    Dim y As Long, m As Long
    y = CLng(Left$(ym, 4)): m = CLng(Mid$(ym, 6, 2))
    m = m + 1
    If m > 12 Then m = 1: y = y + 1
    NextYM = Format(y, "0000") & "-" & Format(m, "00")
End Function

' 設定シートの対象年月を書き換える
Private Sub SetTargetYM(ByVal ym As String)
    Dim ws As Worksheet, f As Range
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    Set f = ws.Columns("B").Find(What:="対象年月", LookAt:=xlWhole)
    If Not f Is Nothing Then f.Offset(0, 1).Value = ym
End Sub

Public Sub RolloverToNextMonth()
    Dim ym As String, nxt As String, ans As VbMsgBoxResult

    ym = TargetYM()
    If Len(ym) = 0 Then Warn "対象年月が未設定です。": Exit Sub

    nxt = NextYM(ym)
    ans = MsgBox(ym & " を締めて " & nxt & " へ進めます。" & vbCrLf & vbCrLf & _
                 "・" & ym & " の資産スナップショットを保存します" & vbCrLf & _
                 "・対象年月を " & nxt & " に変更します" & vbCrLf & vbCrLf & _
                 "続行しますか？", vbQuestion + vbYesNo, APP_TITLE)
    If ans <> vbYes Then Exit Sub

    ' 1) 当月をスナップショットで確定
    SaveMonthEndSnapshot ym
    ' 2) 対象年月を翌月へ
    SetTargetYM nxt
    Application.CalculateFull

    ' 3) 固定収支・サブスクを新しい月へ反映するか確認
    ans = MsgBox(nxt & " に固定収支・サブスクを反映しますか？", _
                 vbQuestion + vbYesNo, APP_TITLE)
    If ans = vbYes Then
        ReflectFixedItems
        ReflectSubscriptions
    End If

    ThisWorkbook.Worksheets(SH_DASH).Activate
    Info "翌月へ進めました。現在の対象年月は " & nxt & " です。"
End Sub
