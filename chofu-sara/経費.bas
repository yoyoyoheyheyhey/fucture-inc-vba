Option Explicit

Public Sub DoWriteKeihi(ByVal wbSrc As Workbook)
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet, wsWrite As Worksheet
    Dim lastRow As Long, r As Long, dstRow As Long
    Dim mainCat As Variant, subCat As Variant, amount As Variant
    Dim written As Long, skipped As Long, cap As Long

    ' 1) ソースシート（経費）を確定（metaは使わない）
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(SHEET_KEIHI)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "ソースブックにシート『" & SHEET_KEIHI & "』が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 2) 出力先（売上日報）を確定・書き込み可能化
    Set wsDst = ThisWorkbook.Worksheets(SHEET_URIAGE)
    If Not EnsureWritable(wsDst, wsWrite) Then
        MsgBox "『" & SHEET_URIAGE & "』の書込先を準備できませんでした。", vbExclamation
        Exit Sub
    End If

    ' 3) 転記準備
    ' wsWrite.Range("B9:D100").ClearContents ' 既存クリア
    cap = 100 ' 最終行
    dstRow = 9
    written = 0: skipped = 0

    ' ソース最終行（主分類B列ベース）
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        ' データなし
        Exit Sub
    End If

    ' 4) 行ごとに転記（B=主分類, C=副分類, F=金額）
    For r = 2 To lastRow
        mainCat = wsSrc.Cells(r, "B").Value
        subCat = wsSrc.Cells(r, "C").Value
        amount = wsSrc.Cells(r, "F").Value

        ' 空行スキップ（主/副/金額すべて空 or 金額=0/空 の場合は好みで調整）
        If IsEmpty(mainCat) And IsEmpty(subCat) And (IsEmpty(amount) Or amount = 0) Then
            ' 何もしない
        Else
            If dstRow > cap Then
                skipped = skipped + 1
            Else
                wsWrite.Cells(dstRow, "B").Value = mainCat   ' 主分類
                wsWrite.Cells(dstRow, "C").Value = subCat    ' 副分類
                wsWrite.Cells(dstRow, "D").Value = amount    ' 金額
                written = written + 1
                dstRow = dstRow + 1
            End If
        End If
    Next r

    ' 5) 結果通知（任意）
    If skipped > 0 Then
        MsgBox "経費の転記を完了しました。" & vbCrLf & _
               "書込み: " & written & " 行 / 省略: " & skipped & " 行（B9:D100の容量超過）", vbInformation
    End If
End Sub
