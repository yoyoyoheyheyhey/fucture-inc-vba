Option Explicit

Public Sub DoWriteUriage(ByVal wbSrc As Workbook)
    Dim wsSrcUriage As Worksheet
    Dim wsDst As Worksheet, wsWrite As Worksheet
    Dim uriageSheetName As Variant

    ' ---- 1) ソースシート確定 ----
    uriageSheetName = GetSheetName(URIAGE_UID, wbSrc)
    If IsError(uriageSheetName) Then
        MsgBox "『" & SHEET_URIAGE & "』のソースUIDがメタ情報で見つかりません。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsSrcUriage = wbSrc.Sheets(CStr(Trim$(uriageSheetName)))
    On Error GoTo 0
    If wsSrcUriage Is Nothing Then
        MsgBox "ソースブックにシート '" & CStr(uriageSheetName) & "' がありません。", vbExclamation
        Exit Sub
    End If

    ' ---- 2) 出力先シート ----
    Set wsDst = ThisWorkbook.Worksheets(SHEET_URIAGE)
    If Not EnsureWritable(wsDst, wsWrite) Then
        MsgBox "『" & SHEET_URIAGE & "』の書込先を準備できませんでした。", vbExclamation
        Exit Sub
    End If

    ' ---- 3) ソースデータ（ヘッダー1行 + データ1行想定） ----
    Dim srcRow As Long: srcRow = 2 ' データは2行目
    Dim valUriage As Variant, valCard As Variant, valShakoPay As Variant
    Dim valFood As Variant, valDanshiPay As Variant

    valUriage = wsSrcUriage.Cells(srcRow, "A").Value    ' 売上
    valCard = wsSrcUriage.Cells(srcRow, "C").Value      ' カード売上
    valShakoPay = wsSrcUriage.Cells(srcRow, "D").Value  ' 社交日払い
    valDanshiPay = wsSrcUriage.Cells(srcRow, "H").Value ' 男子日払い

    ' ---- 4) 出力先に書き込み ----
    wsWrite.Range("C4").Value = valUriage
    wsWrite.Range("C6").Value = valCard
    wsWrite.Range("C7").Value = valShakoPay
    wsWrite.Range("G5").Value = valDanshiPay
End Sub
