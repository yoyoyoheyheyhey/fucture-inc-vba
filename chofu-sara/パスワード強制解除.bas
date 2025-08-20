Option Explicit

Sub UnprotectAllSheetsBrute()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For Each ws In ActiveWorkbook.Worksheets
        If ws.ProtectContents Then
            BruteUnprotectSheet ws
        End If
    Next ws

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub BruteUnprotectSheet(ByVal ws As Worksheet)
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, h As Long
    Dim s As String

    On Error Resume Next

    ' ← 引数なしの Unprotect は使わない（ダイアログになるため）
    ws.Unprotect Password:=""                  '空パスで外れるケース
    If ws.ProtectContents = False Then Exit Sub

    For a = 65 To 66
    For b = 65 To 66
    For c = 65 To 66
    For d = 65 To 66
    For e = 65 To 66
    For f = 65 To 66
    For g = 65 To 66
    For h = 65 To 66
        s = Chr(a) & Chr(b) & Chr(c) & Chr(d) & Chr(e) & Chr(f) & Chr(g) & Chr(h)
        ws.Unprotect Password:=s               '常にパラメータ付きで呼ぶ
        If ws.ProtectContents = False Then
            Debug.Print "解除成功パスワード: " & s
            ' MsgBox "解除成功パスワード: " & s
            Exit Sub
        End If
    Next h: Next g: Next f: Next e: Next d: Next c: Next b: Next a

    On Error GoTo 0
End Sub

