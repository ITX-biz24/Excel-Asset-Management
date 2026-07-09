Option Explicit
'==============================================================
' modPreset - 口座残高の強制プリセット
'--------------------------------------------------------------
' マスタの「初期残高」を実残高として強制適用する。取引履歴と
' スナップショットを全消去し、現在残高＝初期残高 にリセットする。
' 過去の履歴を一切参照しない「強制セット」。デモ値から実運用へ
' 切り替えるときや、残高を実額に合わせ直すときに使う。
'==============================================================

Public Sub ForceSetBalances()
    Dim ans As VbMsgBoxResult

    ' 破壊的操作なので確認（自動テスト等 gSilent 時は確認を省略）
    If Not gSilent Then
        ans = MsgBox( _
            "マスタの初期残高を実残高として強制適用します。" & vbCrLf & _
            "取引履歴とスナップショットをすべて削除し、" & vbCrLf & _
            "現在残高＝初期残高（プリセット値）にリセットします。" & vbCrLf & vbCrLf & _
            "この操作は元に戻せません。実行しますか？", _
            vbExclamation + vbYesNo + vbDefaultButton2, APP_TITLE)
        If ans <> vbYes Then Exit Sub
    End If

    ClearTxnData
    ResetSnapshots
    Application.CalculateFull

    Info "初期残高を強制適用しました。" & vbCrLf & _
         "現在残高＝初期残高（履歴なし）です。総資産 " & _
         Format(TotalAssets(), "#,##0") & " 円。"
End Sub

' 取引履歴を全消去（入力列のみ。年月の計算列は数式を保持）
Private Sub ClearTxnData()
    Dim lo As ListObject
    Set lo = GetTable(TB_TXN)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    gSuppressTxnCheck = True
    ' 1〜8列（日付〜メモ）を消し、9列（年月＝数式）は残す
    lo.DataBodyRange.Columns(COL_DATE).Resize(, COL_MEMO).ClearContents
    gSuppressTxnCheck = False
End Sub

' スナップショットを初期状態（当月ライブ行のみ）へ戻す
Private Sub ResetSnapshots()
    Dim lo As ListObject
    Set lo = GetTable(TB_SNAP)
    If lo Is Nothing Then Exit Sub
    Do While lo.ListRows.Count > 1
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
    If lo.ListRows.Count = 0 Then lo.ListRows.Add
    With lo.DataBodyRange
        .Cells(1, 1).Formula = "=対象年月"
        .Cells(1, 2).Formula = "=SUM(T_Account[現在残高])"
        .Cells(1, 3).Formula = "=IFERROR(INDEX(T_Account[現在残高],MATCH(""現金"",T_Account[口座名],0)),0)"
        .Cells(1, 4).Formula = "=IFERROR(INDEX(T_Account[現在残高],MATCH(""銀行"",T_Account[口座名],0)),0)"
        .Cells(1, 5).Formula = "=IFERROR(INDEX(T_Account[現在残高],MATCH(""QR"",T_Account[口座名],0)),0)"
    End With
End Sub
