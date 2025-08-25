Option Explicit

Private Const SRC_START_ROW As Long = 2 ' ソースのデータ開始行
Private Const DST_START_ROW As Long = 6 ' 出力先の書き込み開始行
Private Const TOTAL_FIND_COL As Long = 1 ' 合計行を探す列（A列=1）

Public Sub DoWriteNippo(wbSrc As Workbook)
    Dim wsSrc As Worksheet, wsWrite As Worksheet

    If Not ResolveSrcAndDst(wbSrc, NIPPO_UID, SHEET_NIPPO, wsSrc, wsWrite) Then Exit Sub

    WriteNippo wsWrite, wsSrc
End Sub

' ソース（社交 | 日報）→ 日報に書き込み
Private Sub WriteNippo(ByVal wsWrite As Worksheet, ByVal wsSrc As Worksheet)
    Dim lastRowSrc As Long, r As Long, dstRow As Long
    Dim totalRow As Long, maxRow As Long

    ' ソースの最終行（ユーザー=A列基準）
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    ' 合計行の上までに制限（合計を壊さない）
    totalRow = FindTotalRow(wsWrite)
    If totalRow > 0 Then
        maxRow = totalRow - 1
    Else
        maxRow = wsWrite.Rows.Count - 1
    End If

    dstRow = DST_START_ROW
    For r = SRC_START_ROW To lastRowSrc
        If dstRow > maxRow Then Exit For

        ' ソースのデータを変数に格納
        Dim name As Variant, startAt As Variant, endAt As Variant
        Dim honshiP As Variant, jonaiP As Variant, extP As Variant, dohanP As Variant
        Dim honshiBack As Variant, jonaiBack As Variant, dohanBack As Variant
        Dim drinkBack As Variant, adpPay As Variant, sendingFee As Variant
        Dim uniformFee As Variant, penalty As Variant, honshiSales As Variant
        Dim empNo As Variant
        name = wsSrc.Cells(r, "A").Value ' 名前
        startAt = wsSrc.Cells(r, "B").Value ' 出勤
        endAt = wsSrc.Cells(r, "D").Value ' 退勤
        honshiP = wsSrc.Cells(r, "F").Value ' 本指P
        jonaiP = wsSrc.Cells(r, "G").Value ' 場内P
        extP = wsSrc.Cells(r, "H").Value ' 延長P
        dohanP = wsSrc.Cells(r, "I").Value ' 同伴P
        honshiBack = wsSrc.Cells(r, "J").Value ' 本指ﾊﾞｯｸ
        jonaiBack = wsSrc.Cells(r, "K").Value ' 場内ﾊﾞｯｸ
        dohanBack = wsSrc.Cells(r, "L").Value ' 同伴ﾊﾞｯｸ
        drinkBack = wsSrc.Cells(r, "M").Value ' ドリンクバック
        adpPay = wsSrc.Cells(r, "N").Value ' 日払い
        sendingFee = wsSrc.Cells(r, "O").Value ' 送り
        uniformFee = wsSrc.Cells(r, "P").Value ' 制服
        penalty = wsSrc.Cells(r, "Q").Value ' ﾍﾟﾅﾙﾃｨ
        honshiSales = wsSrc.Cells(r, "R").Value ' 本指売上

        ' --- 空/0判定（全て空や0ならスキップ） ---
        Dim arr As Variant
        arr = Array(startAt, endAt, honshiP, jonaiP, extP, dohanP, _
                    honshiBack, jonaiBack, dohanBack, drinkBack, adpPay, sendingFee, _
                    uniformFee, penalty, honshiSales)
        Dim allBlank As Boolean: allBlank = True
        Dim v As Variant
        For Each v In arr
            If Not IsEmpty(v) Then
                If VarType(v) = vbString Then
                    If Len(Trim$(v)) > 0 Then allBlank = False: Exit For
                ElseIf IsNumeric(v) Then
                    If v <> 0 Then allBlank = False: Exit For
                Else
                    ' その他型（Dateなど）は値があれば有効とみなす
                    If Len(CStr(v)) > 0 Then allBlank = False: Exit For
                End If
            End If
        Next

        If allBlank Then
            ' --- この行は全部空/0だったのでスキップ ---
        Else
            empNo = GetEmployeeNumber("社交", name)
            ' --- 書き込み ---
            wsWrite.Cells(dstRow, "B").Value = empNo
            wsWrite.Cells(dstRow, "D").Value = startAt
            wsWrite.Cells(dstRow, "F").Value = endAt
            wsWrite.Cells(dstRow, "H").Value = honshiP
            wsWrite.Cells(dstRow, "I").Value = jonaiP
            wsWrite.Cells(dstRow, "J").Value = extP
            wsWrite.Cells(dstRow, "K").Value = dohanP
            wsWrite.Cells(dstRow, "L").Value = honshiBack
            wsWrite.Cells(dstRow, "M").Value = jonaiBack
            wsWrite.Cells(dstRow, "N").Value = dohanBack
            wsWrite.Cells(dstRow, "Q").Value = drinkBack
            wsWrite.Cells(dstRow, "R").Value = adpPay
            wsWrite.Cells(dstRow, "S").Value = sendingFee
            wsWrite.Cells(dstRow, "T").Value = uniformFee
            wsWrite.Cells(dstRow, "U").Value = penalty
            wsWrite.Cells(dstRow, "V").Value = honshiSales

            dstRow = dstRow + 1
        End If
    Next r
End Sub

' A列から「合」を含むセルを探して、その行番号を返す（見つからなければ 0）
Private Function FindTotalRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Columns(TOTAL_FIND_COL).Find(What:="合", LookAt:=xlPart, LookIn:=xlValues, _
                                            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not f Is Nothing Then
        FindTotalRow = f.Row
    Else
        FindTotalRow = 0
    End If
End Function