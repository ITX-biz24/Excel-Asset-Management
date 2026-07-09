Option Explicit
'==============================================================
' modInputCheck - 入力チェック（整合性・残高不足・必須）
'--------------------------------------------------------------
' 取引1件の妥当性を検証する。フォーム入力・シート直接入力の
' 両方から呼ばれる二次防御。問題があれば理由を文字列で返す（""=OK）。
'==============================================================

' 取引1件を検証。戻り値 "" なら妥当、そうでなければエラー理由。
' checkBalance:=True のとき残高不足も判定する。
Public Function ValidateTxn(ByVal dtVal As Variant, ByVal kind As String, _
                            ByVal fromAcc As String, ByVal toAcc As String, _
                            ByVal amountVal As Variant, ByVal category As String, _
                            Optional ByVal checkBalance As Boolean = True) As String
    Dim amt As Double

    ' 必須: 日付
    If Not IsDate(dtVal) Then
        ValidateTxn = "日付を正しく入力してください。": Exit Function
    End If

    ' 必須: 金額 > 0
    If Not IsNumeric(amountVal) Then
        ValidateTxn = "金額を数値で入力してください。": Exit Function
    End If
    amt = CDbl(amountVal)
    If amt <= 0 Then
        ValidateTxn = "金額は0より大きい値を入力してください。": Exit Function
    End If

    ' 種類
    Select Case kind
        Case KIND_INCOME
            If Len(toAcc) = 0 Then ValidateTxn = "収入は入金先を指定してください。": Exit Function
            If Len(fromAcc) > 0 Then ValidateTxn = "収入では出金元を空欄にしてください。": Exit Function
            If Len(category) = 0 Then ValidateTxn = "カテゴリを選択してください。": Exit Function
        Case KIND_EXPENSE
            If Len(fromAcc) = 0 Then ValidateTxn = "支出は出金元を指定してください。": Exit Function
            If Len(toAcc) > 0 Then ValidateTxn = "支出では入金先を空欄にしてください。": Exit Function
            If Len(category) = 0 Then ValidateTxn = "カテゴリを選択してください。": Exit Function
        Case KIND_TRANSFER
            If Len(fromAcc) = 0 Or Len(toAcc) = 0 Then
                ValidateTxn = "振替は出金元・入金先の両方を指定してください。": Exit Function
            End If
            If fromAcc = toAcc Then
                ValidateTxn = "振替の出金元と入金先を同じ口座にはできません。": Exit Function
            End If
        Case Else
            ValidateTxn = "種類は 収入 / 支出 / 振替 から選択してください。": Exit Function
    End Select

    ' 残高不足（出金が発生する種類のみ）
    If checkBalance And Len(fromAcc) > 0 Then
        If AccountBalance(fromAcc) < amt Then
            ValidateTxn = fromAcc & " の残高が不足しています（残高 " & _
                          Format(AccountBalance(fromAcc), "#,##0") & " 円 / 出金 " & _
                          Format(amt, "#,##0") & " 円）。"
            Exit Function
        End If
    End If

    ValidateTxn = ""
End Function

' 取引履歴シートで完成した行を検証する（Worksheet_Change から呼ばれる）。
' 行が「日付・種類・金額」まで揃ったときだけ検証し、問題があれば警告する。
Public Sub CheckTxnRow(ByVal r As Range)
    Dim lo As ListObject
    Set lo = GetTable(TB_TXN)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub

    Dim dtVal As Variant, kind As String, fromAcc As String, toAcc As String
    Dim amountVal As Variant, category As String
    Dim base As Range
    Set base = lo.DataBodyRange.Rows(r.Row - lo.DataBodyRange.Row + 1)

    dtVal = base.Cells(1, COL_DATE).Value
    kind = CStr(base.Cells(1, COL_KIND).Value)
    fromAcc = CStr(base.Cells(1, COL_FROM).Value)
    toAcc = CStr(base.Cells(1, COL_TO).Value)
    amountVal = base.Cells(1, COL_AMOUNT).Value
    category = CStr(base.Cells(1, COL_CATEGORY).Value)

    ' 未完成の行はチェックしない（入力途中を邪魔しない）
    If Not IsDate(dtVal) Then Exit Sub
    If Len(kind) = 0 Then Exit Sub
    If Not IsNumeric(amountVal) Then Exit Sub

    Dim msg As String
    msg = ValidateTxn(dtVal, kind, fromAcc, toAcc, amountVal, category, _
                      ConfigIsOn("残高不足チェック"))
    If Len(msg) > 0 Then
        Warn "入力内容を確認してください:" & vbCrLf & vbCrLf & msg
    End If
End Sub
