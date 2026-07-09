Option Explicit
'==============================================================
' modFixed - 固定収支・サブスクの当月反映
'--------------------------------------------------------------
' 有効な固定収支／サブスクを、対象年月の取引履歴へ一括追加する。
' メモに反映マークを書き、同じ月への二重反映を防ぐ。
'==============================================================

' 対象年月に、指定メモタグの取引が既に存在するか
Private Function AlreadyReflected(ByVal ym As String, ByVal tag As String) As Boolean
    Dim lo As ListObject, i As Long
    Set lo = GetTable(TB_TXN)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, COL_YM).Value) = ym Then
            If CStr(lo.DataBodyRange.Cells(i, COL_MEMO).Value) = tag Then
                AlreadyReflected = True
                Exit Function
            End If
        End If
    Next i
End Function

'--------------------------------------------------------------
' 固定収支を当月へ反映
'--------------------------------------------------------------
Public Sub ReflectFixedItems()
    Dim lo As ListObject, i As Long
    Dim ym As String, added As Long, skipped As Long
    Dim nm As String, kind As String, amt As Double, dayNum As Long
    Dim acc As String, cat As String, pay As String, tag As String
    Dim fromAcc As String, toAcc As String

    ym = TargetYM()
    If Len(ym) = 0 Then Warn "設定シートの対象年月が未設定です。": Exit Sub
    Set lo = GetTable(TB_FIXED)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    For i = 1 To lo.ListRows.Count
        nm = Trim(CStr(lo.DataBodyRange.Cells(i, 1).Value))
        If Len(nm) > 0 And UCase(CStr(lo.DataBodyRange.Cells(i, 8).Value)) = "ON" Then
            kind = CStr(lo.DataBodyRange.Cells(i, 2).Value)
            amt = Val(lo.DataBodyRange.Cells(i, 3).Value)
            dayNum = Val(lo.DataBodyRange.Cells(i, 4).Value)
            acc = CStr(lo.DataBodyRange.Cells(i, 5).Value)
            cat = CStr(lo.DataBodyRange.Cells(i, 6).Value)
            pay = CStr(lo.DataBodyRange.Cells(i, 7).Value)
            tag = MARK_FIXED & nm

            If amt > 0 And Not AlreadyReflected(ym, tag) Then
                If kind = KIND_INCOME Then
                    fromAcc = "": toAcc = acc
                Else
                    fromAcc = acc: toAcc = ""
                End If
                AddTransaction DateInMonth(ym, dayNum), kind, fromAcc, toAcc, amt, cat, pay, tag
                added = added + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    Application.CalculateFull

    Info "固定収支の反映が完了しました（" & ym & "）。" & vbCrLf & _
         "追加: " & added & " 件 / スキップ(反映済み等): " & skipped & " 件"
End Sub

'--------------------------------------------------------------
' サブスクを当月へ反映（支出として計上）
'--------------------------------------------------------------
Public Sub ReflectSubscriptions()
    Dim lo As ListObject, i As Long
    Dim ym As String, added As Long, skipped As Long
    Dim nm As String, amt As Double, dayNum As Long
    Dim acc As String, cat As String, pay As String, tag As String
    Dim startDt As Variant, monthEnd As Date

    ym = TargetYM()
    If Len(ym) = 0 Then Warn "設定シートの対象年月が未設定です。": Exit Sub
    monthEnd = DateInMonth(ym, 31)
    Set lo = GetTable(TB_SUB)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    For i = 1 To lo.ListRows.Count
        nm = Trim(CStr(lo.DataBodyRange.Cells(i, 1).Value))
        If Len(nm) > 0 And UCase(CStr(lo.DataBodyRange.Cells(i, 8).Value)) = "ON" Then
            startDt = lo.DataBodyRange.Cells(i, 2).Value
            dayNum = Val(lo.DataBodyRange.Cells(i, 3).Value)
            amt = Val(lo.DataBodyRange.Cells(i, 4).Value)
            pay = CStr(lo.DataBodyRange.Cells(i, 5).Value)
            acc = CStr(lo.DataBodyRange.Cells(i, 6).Value)
            cat = CStr(lo.DataBodyRange.Cells(i, 7).Value)
            tag = MARK_SUB & nm

            ' 開始日が当月末以前のサブスクのみ対象
            If amt > 0 And (Not IsDate(startDt) Or CDate(startDt) <= monthEnd) _
               And Not AlreadyReflected(ym, tag) Then
                AddTransaction DateInMonth(ym, dayNum), KIND_EXPENSE, acc, "", amt, cat, pay, tag
                added = added + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    Application.CalculateFull

    Info "サブスクの反映が完了しました（" & ym & "）。" & vbCrLf & _
         "追加: " & added & " 件 / スキップ(反映済み等): " & skipped & " 件"
End Sub
