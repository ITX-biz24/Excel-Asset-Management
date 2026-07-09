Option Explicit
'==============================================================
' modBackup - バックアップ
'--------------------------------------------------------------
' ブックをタイムスタンプ付きで複製保存する。保存先は設定シートの
' 「バックアップ先」。相対パスならブックと同じ場所を基準にする。
'==============================================================

Public Sub BackupWorkbook(Optional ByVal silent As Boolean = False)
    Dim folder As String, baseName As String, ext As String
    Dim stamp As String, dest As String

    folder = Trim(CStr(GetConfig("バックアップ先")))
    If Len(folder) = 0 Then folder = "backup"

    ' 相対パスならブックの場所を基準に解決
    If Not (Mid$(folder, 2, 1) = ":" Or Left$(folder, 2) = "\\") Then
        If Len(ThisWorkbook.Path) = 0 Then
            Warn "先にブックを保存してからバックアップしてください。": Exit Sub
        End If
        folder = ThisWorkbook.Path & Application.PathSeparator & folder
    End If

    ' フォルダが無ければ作成
    If Dir(folder, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folder
        On Error GoTo 0
    End If

    ' ファイル名（拡張子は元ブックに合わせる）
    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then
        ext = Mid$(baseName, InStrRev(baseName, "."))
        baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    Else
        ext = ".xlsm"
    End If
    stamp = Format(Now, "yyyymmdd_hhnnss")
    dest = folder & Application.PathSeparator & baseName & "_" & stamp & ext

    On Error GoTo failed
    ThisWorkbook.SaveCopyAs dest
    If Not silent Then Info "バックアップを保存しました。" & vbCrLf & dest
    Exit Sub
failed:
    If Not silent Then Warn "バックアップに失敗しました: " & Err.Description
End Sub
