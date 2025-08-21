Option Explicit

Public Sub YonareziImport()
    Dim fd As FileDialog, srcFile As String
    Dim wbSrc As Workbook
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

    DoWriteNippo wbSrc
    DoWriteUriage wbSrc
    DoWriteDanshiHibarai wbSrc

CloseSrc:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    MsgBox "取り込み完了。", vbInformation

Tidy:
    Application.DisplayAlerts = al
    Application.EnableEvents = ev
    Application.ScreenUpdating = scr
End Sub
