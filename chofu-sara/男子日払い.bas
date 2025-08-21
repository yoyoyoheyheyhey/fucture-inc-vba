Option Explicit

Public Sub DoWriteDanshiHibarai(ByVal wbSrc As Workbook)
    Dim wsDst As Worksheet, wsWrite As Worksheet, wsSrc As Worksheet
    Dim tgtDate As Date

    ' 1) ファイル名から終了日（対象日）
    If Not TryParseEndDateFromFileName(wbSrc.Name, tgtDate) Then
        MsgBox "ソース名から終了日が取得できません: " & wbSrc.Name, vbExclamation
        Exit Sub
    End If
    
    ' 出力先は一度だけ準備
    Set wsDst = ThisWorkbook.Worksheets(SHEET_DANSHI_HIBARAI)
    If Not EnsureWritable(wsDst, wsWrite) Then Exit Sub
    
    ' 男子
    If ResolveSrcAndDst(wbSrc, DANSHI_HIBARAI_DANSHI_UID, SHEET_DANSHI_HIBARAI, wsSrc, wsWrite, wsWrite, True) Then
        WriteDanshiHibarai wsSrc, wsWrite, tgtDate
    End If
    
    ' アルバイト
    If ResolveSrcAndDst(wbSrc, DANSHI_HIBARAI_PART_UID, SHEET_DANSHI_HIBARAI, wsSrc, wsWrite, wsWrite, True) Then
        WriteDanshiHibarai wsSrc, wsWrite, tgtDate
    End If
End Sub

' 指定ソースシート（A:名前, B:日払い額, 1行目ヘッダ）を
' 出力シート E4:AN4（ヘッダ名）× B5:B35（日）へ反映（同一人は加算）
Private Sub WriteDanshiHibarai( _
    ByVal wsSrc As Worksheet, _
    ByVal wsWrite As Worksheet, _
    ByVal tgtDate As Date)

    Dim lastRow As Long, r As Long
    Dim nm As String, amt As Double
    Dim col As Long, row As Long
    Dim cur As Variant

    ' 書き込み先の行（対象日）
    row = 4 + Day(tgtDate) ' B5が1日目 → 4 + day
    If row < 5 Or row > 35 Then Exit Sub ' 念のため

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow ' 1行目はヘッダ
        nm = Trim$(CStr(wsSrc.Cells(r, "A").Value))
        If Len(nm) = 0 Then GoTo ContinueLoop

        If IsNumeric(wsSrc.Cells(r, "B").Value) Then
            amt = CDbl(wsSrc.Cells(r, "B").Value)
        Else
            amt = 0
        End If
        If amt = 0 Then GoTo ContinueLoop

        ' ヘッダ E4:AN4 を完全一致検索（必ず wsWrite を親に）
        col = FindHeaderColumn(wsWrite, nm, wsWrite.Range("E4"), wsWrite.Range("AN4"))
        If col = 0 Then GoTo ContinueLoop ' 見つからなければスキップ（拡張可）

        ' 既存値に加算
        cur = wsWrite.Cells(row, col).Value
        If IsNumeric(cur) Then
            wsWrite.Cells(row, col).Value = CDbl(cur) + amt
        ElseIf IsEmpty(cur) Or Len(CStr(cur)) = 0 Then
            wsWrite.Cells(row, col).Value = amt
        Else
            wsWrite.Cells(row, col).Value = amt
        End If

ContinueLoop:
    Next r
End Sub