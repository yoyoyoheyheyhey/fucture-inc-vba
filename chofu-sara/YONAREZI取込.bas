Option Explicit

Public Sub YonareziImport()
    Dim fd As FileDialog, srcFile As String
    Dim wbSrc As Workbook
    Dim wsSrcNippo As Worksheet
    Dim wsDst As Worksheet, wsWrite As Worksheet
    Dim scr As Boolean, ev As Boolean, al As Boolean

    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    ev = Application.EnableEvents:    Application.EnableEvents = False
    al = Application.DisplayAlerts:   Application.DisplayAlerts = False

    ' 1) ファイル選択
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "取り込むソースファイルを選択してください"
        .Filters.Clear: .Filters.Add "Excelファイル", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo Tidy
        srcFile = .SelectedItems(1)
    End With

    ' 2) 開く
    Set wbSrc = Workbooks.Open(srcFile, ReadOnly:=True)

    ' 3) 日報：EnsureWritable → 書き込み
    On Error Resume Next
    Dim nippoSheetName As String
    nippoSheetName = GetSheetName(NIPPO_UID, wbSrc) 
    Set wsSrcNippo = wbSrc.Sheets(nippoSheetName)
    On Error GoTo 0
    If wsSrcNippo Is Nothing Then
        MsgBox "『" & SHEET_NIPPO & "』の書き込むソースが見つかりません。", vbExclamation
        GoTo CloseSrc
    End If

    Set wsDst = ThisWorkbook.Worksheets(SHEET_NIPPO)
    If EnsureWritable(wsDst, wsWrite) Then
        Call WriteNippo(wsWrite, wsSrcNippo)
    Else
        MsgBox "『" & SHEET_NIPPO & "』の書込先を準備できませんでした。", vbExclamation
    End If

    ' 4) 売上日報：EnsureWritable → 書き込み（マッピング完成後に有効化）
    'Dim wsSrcUriage As Worksheet
    'On Error Resume Next
    'Set wsSrcUriage = wbSrc.Sheets("（売上日報の元シート名）")
    'On Error GoTo 0
    'If Not wsSrcUriage Is Nothing Then
    '    Set wsDst = ThisWorkbook.Worksheets(SHEET_URIAGE)
    '    If EnsureWritable(wsDst, wsWrite) Then
    '        Call WriteUriage(wsWrite, wsSrcUriage)
    '    End If
    'End If

CloseSrc:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    MsgBox "取り込み完了。", vbInformation

Tidy:
    Application.DisplayAlerts = al
    Application.EnableEvents = ev
    Application.ScreenUpdating = scr
End Sub
