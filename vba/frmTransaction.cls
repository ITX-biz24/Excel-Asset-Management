Option Explicit
'==============================================================
' frmTransaction - 取引入力フォーム（コード）
'--------------------------------------------------------------
' コントロールはビルド時(inject_vba.py)に生成される。ここは動作コード。
' 追加前に modInputCheck.ValidateTxn で検証し、不正なら追加しない（一次防御）。
'==============================================================

Private Sub UserForm_Initialize()
    FillCombo cboKind, "種類リスト"
    FillCombo cboFrom, "口座リスト"
    FillCombo cboTo, "口座リスト"
    FillCombo cboCategory, "カテゴリリスト"
    FillCombo cboPay, "支払方法リスト"
    txtDate.Value = Format(Date, "yyyy/mm/dd")
    cboKind.Value = KIND_EXPENSE
End Sub

' 名前付き範囲からコンボの選択肢を読み込む
Private Sub FillCombo(ByRef cbo As MSForms.ComboBox, ByVal rangeName As String)
    Dim c As Range
    cbo.Clear
    On Error Resume Next
    For Each c In Application.Range(rangeName)
        If Len(CStr(c.Value)) > 0 Then cbo.AddItem CStr(c.Value)
    Next c
    On Error GoTo 0
End Sub

Private Sub btnAdd_Click()
    Dim msg As String
    If Not IsDate(txtDate.Value) Then
        MsgBox "日付を正しく入力してください。", vbExclamation, APP_TITLE: Exit Sub
    End If
    If Not IsNumeric(txtAmount.Value) Then
        MsgBox "金額を数値で入力してください。", vbExclamation, APP_TITLE: Exit Sub
    End If

    msg = ValidateTxn(CDate(txtDate.Value), CStr(cboKind.Value), CStr(cboFrom.Value), _
                      CStr(cboTo.Value), txtAmount.Value, CStr(cboCategory.Value), _
                      ConfigIsOn("残高不足チェック"))
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, APP_TITLE: Exit Sub
    End If

    AddTransaction CDate(txtDate.Value), CStr(cboKind.Value), CStr(cboFrom.Value), _
             CStr(cboTo.Value), CDbl(txtAmount.Value), CStr(cboCategory.Value), _
             CStr(cboPay.Value), CStr(txtMemo.Value)
    Application.CalculateFull

    ' 連続入力のため金額・メモのみクリア
    txtAmount.Value = ""
    txtMemo.Value = ""
    txtAmount.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
