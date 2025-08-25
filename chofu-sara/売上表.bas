Option Explicit

Private Const SRC_START_ROW As Long = 2 ' ソースのデータ開始行
Private Const DST_START_ROW As Long = 5 ' 出力先の書き込み開始行
Private Const DST_LAST_ROW As Long = 35 ' 出力先の書き込み終了行（月ごとに31行分確保している想定）

Public Sub DoWriteUriageHyo(ByVal wbSrc As Workbook)
    Dim wsSrc As Worksheet, wsWrite As Worksheet
    Dim tgtDate As Date

    ' ファイル名から終了日（対象日）
    If Not TryParseEndDateFromFileName(wbSrc.Name, tgtDate) Then
        MsgBox "ソース名から終了日が取得できません: " & wbSrc.Name, vbExclamation
        Exit Sub
    End If

    If Not ResolveSrcAndDst(wbSrc, URIAGE_NIPPO_UID, SHEET_URIAGE_HYO, wsSrc, wsWrite) Then Exit Sub

    WriteUriageHyo wsSrc, wsWrite, tgtDate
End Sub

Private Sub WriteUriageHyo( _
    ByVal wsSrc As Worksheet, _
    ByVal wsWrite As Worksheet, _
    ByVal tgtDate As Date)

    Dim lastRowSrc As Long, r As Long
    Dim nm As String, amt As Double
    Dim col As Long, row As Long
    Dim cur As Variant

    ' 書き込み先の行（対象日）
    row = DST_START_ROW + Day(tgtDate) - 1 ' 開始行が1日目, e.g. 対象日が1/5なら1行目からスタートして5日目なので、開始行 + 5 - 1
    If row < DST_START_ROW Or row > DST_LAST_ROW Then Exit Sub ' 念のため

    ' 出力先のセル番地
    '  - Ｂ本数: L<row>
    '  - Ｂ売上: M<row>
    '  - 社交数: N<row>
    '  - 社時: O<row>
    '  - 組数: P<row>
    '  - 客数: Q<row>
    '  - 客単: R<row>
    '  - Ｔ.Ａ: S<row>
    '  - 柏屋8％: T<row>
    '  - 柏屋10％: U<row>
    srcCols = Array("K", "L", "M", "N", "O", "P", "Q", "R", "S", "T")
    ' ソースのセル番地
    '  - 売上: A2
    '  - 現金売上: B2
    '  - カード売上: C2
    '  - 売掛: D2
    '  - 食材仕入: E2
    '  - 仕入: F2
    '  - 仕入合計: G2
    '  - 社交日払い: H2
    '  - 男子日払い: I2
    '  - 入金額: J2
    '  - Ｂ本数: K2
    '  - Ｂ売上: L2
    '  - 社交数: M2
    '  - 社時: N2
    '  - 組数: O2
    '  - 客数: P2
    '  - 客単: Q2
    '  - T.A: R2
    '  - 仕入/酒代(柏屋)/8%: S2
    '  - 仕入/酒代(柏屋)/10%: T2
    dstCols = Array("L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")

    ' コピー処理
    For i = LBound(srcCols) To UBound(srcCols)
        wsWrite.Cells(rowDst, Columns(dstCols(i)).Column).Value = _
            wsSrc.Cells(2, Columns(srcCols(i)).Column).Value
    Next i
End Sub