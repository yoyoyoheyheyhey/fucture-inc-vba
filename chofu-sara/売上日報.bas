Option Explicit

Public Sub DoWriteUriage(ByVal wbSrc As Workbook)
    Dim wsSrc As Worksheet, wsWrite As Worksheet

    If Not ResolveSrcAndDst(wbSrc, URIAGE_UID, SHEET_URIAGE, wsSrc, wsWrite) Then Exit Sub

    Dim r As Long: r = 2 ' データは2行目
    Dim valUriage As Variant, valCard As Variant, valShakoPay As Variant
    Dim valFood As Variant, valDanshiPay As Variant

    With wsSrc
        valUriage   = .Cells(r, "A").Value   ' 売上
        valCard     = .Cells(r, "C").Value   ' カード売上
        valShakoPay = .Cells(r, "D").Value   ' 社交日払い
        valDanshiPay = .Cells(r, "H").Value  ' 男子日払い
    End With

    With wsWrite
        .Range("C4").Value = valUriage
        .Range("C6").Value = valCard
        .Range("C7").Value = valShakoPay
        .Range("G5").Value = valDanshiPay
    End With
End Sub
