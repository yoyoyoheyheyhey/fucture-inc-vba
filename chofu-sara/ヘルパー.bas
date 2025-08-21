Option Explicit

' 指定シートを「書き込み可能」にする。
' 成功時 True → wsWrite に対象シートを返す。
' 解除できない／書き込みテスト失敗なら False。
Public Function EnsureWritable(ByVal ws As Worksheet, ByRef wsWrite As Worksheet) As Boolean
    Dim pw As String, v As Variant

    Set wsWrite = Nothing
    pw = GetSheetPassword(ws.Name)   ' パスワードが分かっていれば返す（不明なら ""）

    On Error Resume Next

    ' 0) すでに無保護
    If Not ws.ProtectContents Then
        Set wsWrite = ws
        EnsureWritable = True
        Exit Function
    End If

    ' 1) 既知パスで解除を試す
    If Len(pw) > 0 Then
        ws.Unprotect Password:=pw
        If Not ws.ProtectContents Then
            ' 再保護しつつ VBA から書けるように
            ws.Protect Password:=pw, UserInterfaceOnly:=True
            Set wsWrite = ws
            EnsureWritable = True
            Exit Function
        End If
    End If

    ' 2) UIOnly だけ付けて書けるかテスト
    ws.Protect Password:=pw, UserInterfaceOnly:=True
    v = ws.Range("XFD1048576").Value
    Err.Clear
    ws.Range("XFD1048576").Value = v
    If Err.Number = 0 Then
        Set wsWrite = ws
        EnsureWritable = True
        Exit Function
    End If

    ' 3) ここまで来たら失敗
    On Error GoTo 0
    EnsureWritable = False
End Function

' ワークシートから呼べる関数: =GetEmployeeNumber("社交","みずき")
Public Function GetEmployeeNumber(ByVal role As String, ByVal name As String) As Variant
    Dim ws As Worksheet
    Dim numCol As Long, nameCol As Long, startRow As Long
    Dim lastRow As Long, r As Long
    Dim targetName As String, cellName As String

    On Error GoTo ErrHandler

    ' 対象シート
    Set ws = ThisWorkbook.Worksheets("名簿")

    ' ロール別の列・開始行
    Select Case CleanName(role)
        Case "社交"
            numCol = Columns("B").Column   ' 社員番号
            nameCol = Columns("C").Column  ' 名前
            startRow = 3
        Case "男子"
            numCol = Columns("K").Column   ' 社員番号
            nameCol = Columns("L").Column  ' 名前
            startRow = 3
        Case "アルバイト"
            numCol = Columns("K").Column   ' 社員番号
            nameCol = Columns("L").Column  ' 名前
            startRow = 16
        Case Else
            GetEmployeeNumber = CVErr(xlErrValue)      ' 不正なrole
            Exit Function
    End Select

    ' 検索範囲の最終行
    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    If lastRow < startRow Then
        GetEmployeeNumber = CVErr(xlErrNA)             ' データなし
        Exit Function
    End If

    ' 入力名の整形（前後スペース/全角スペースを吸収）
    targetName = CleanName(name)

    ' 1行ずつ一致チェック（厳密一致、大小区別なし・前後スペース無視）
    For r = startRow To lastRow
        cellName = CleanName(CStr(ws.Cells(r, nameCol).Value))
        If StrComp(cellName, targetName, vbTextCompare) = 0 Then
            GetEmployeeNumber = ws.Cells(r, numCol).Value
            Exit Function
        End If
    Next r

    ' 見つからない
    GetEmployeeNumber = CVErr(xlErrNA)
    Exit Function

ErrHandler:
    GetEmployeeNumber = CVErr(xlErrNA)
End Function

' UIDからシート名を返す関数
' =GetSheetName("任意のUID") のように呼び出す
Public Function GetSheetName(ByVal uid As String, Optional ByVal wb As Workbook = Nothing) As Variant
    Dim ws As Worksheet
    Dim f As Range
    Dim target As String

    On Error GoTo ErrHandler

    ' 検索対象のブックを決定（未指定ならThisWorkbookを使うフォールバック）
    If wb Is Nothing Then Set wb = ThisWorkbook

    ' ____meta____ シート取得
    Set ws = wb.Worksheets("____meta____")

    ' UID正規化（前後スペース除去）
    target = Trim$(CStr(uid))

    ' A列から完全一致で検索
    With ws.Columns(1)
        Set f = .Find(What:=target, LookAt:=xlWhole, LookIn:=xlValues, _
                      SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With

    If Not f Is Nothing Then
        GetSheetName = ws.Cells(f.Row, 2).Value  ' 同じ行のB列（シート名）
    Else
        GetSheetName = CVErr(xlErrNA)            ' 見つからない
    End If
    Exit Function

ErrHandler:
    GetSheetName = CVErr(xlErrValue)
End Function

' ソースUIDと出力先シート名（定数）を受けて、
' - ソースブック(wbSrc)から該当シートを取得（____meta____参照）
' - 出力先ThisWorkbookのシートをEnsureWritableで取得
' - wsWriteReuse を渡せば出力先の準備はスキップ（再利用）
' - silentOnMissing=True なら、ソース未存在時にメッセージなしで False を返す
' - 成功時 True を返し、wsSrc / wsWrite にセット
Public Function ResolveSrcAndDst( _
    ByVal wbSrc As Workbook, _
    ByVal uid As String, _
    ByVal dstSheetName As String, _
    ByRef wsSrc As Worksheet, _
    ByRef wsWrite As Worksheet, _
    Optional ByVal wsWriteReuse As Worksheet = Nothing, _
    Optional ByVal silentOnMissing As Boolean = False _
) As Boolean

    Dim wsDst As Worksheet
    Set wsSrc = Nothing

    ' 1) ソース（静かに取得）
    Set wsSrc = ResolveSrcSheet(wbSrc, uid)
    If wsSrc Is Nothing Then
        If Not silentOnMissing Then
            Dim shName As Variant: shName = GetSheetName(uid, wbSrc)
            If IsError(shName) Then
                MsgBox "メタ情報(____meta____)に UID が見つかりません。" & vbCrLf & "UID: " & uid, vbExclamation
            Else
                MsgBox "ソースブックにシート '" & CStr(shName) & "' が見つかりません。", vbExclamation
            End If
        End If
        Exit Function
    End If

    ' 2) 出力先（再利用が来ていればそれを使う）
    If Not wsWriteReuse Is Nothing Then
        Set wsWrite = wsWriteReuse
    Else
        Set wsDst = ThisWorkbook.Worksheets(dstSheetName)
        If Not EnsureWritable(wsDst, wsWrite) Then
            If Not silentOnMissing Then _
                MsgBox "出力先シート『" & dstSheetName & "』を書き込み可能にできませんでした。", vbExclamation
            Exit Function
        End If
    End If

    ResolveSrcAndDst = True
End Function


' ファイル名「<title> YYYY-MM-DD ~ YYYY-MM-DD」から後ろ側の日付を拾う
' 成功時 True / tgt に日付
Public Function TryParseEndDateFromFileName(ByVal fileName As String, ByRef tgt As Date) As Boolean
    Dim re As Object, mc As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "(\d{4}-\d{2}-\d{2})"
        .Global = True
        .IgnoreCase = True
    End With
    Set mc = re.Execute(fileName)
    If mc.Count >= 2 Then
        Set m = mc(mc.Count - 1) ' 最後の一致（終了日）
        On Error Resume Next
        tgt = CDate(m.Value)
        On Error GoTo 0
        TryParseEndDateFromFileName = (tgt > 0)
    Else
        TryParseEndDateFromFileName = False
    End If
End Function

' ヘッダー名から列番号を返す
Public Function FindHeaderColumn(ByVal ws As Worksheet, _
                                  ByVal headerName As String, _
                                  ByVal leftTop As Range, _
                                  ByVal rightTop As Range) As Long
    Dim rng As Range, f As Range
    Set rng = ws.Range(leftTop, rightTop)
    Set f = rng.Find(What:=headerName, LookAt:=xlWhole, LookIn:=xlValues)

    If f Is Nothing Then
        FindHeaderColumn = 0
    Else
        FindHeaderColumn = f.Column
    End If
End Function

' 名前の簡易正規化：前後スペース削除＋全角スペース→半角
Private Function CleanName(ByVal s As String) As String
    Dim zsp As String
    zsp = ChrW(&H3000) ' 全角スペース
    s = Replace$(s, zsp, " ")
    s = Trim$(s)
    CleanName = s
End Function

' UID→ソースシート取得（見つからなければ Nothing）
Private Function ResolveSrcSheet(ByVal wbSrc As Workbook, ByVal uid As String) As Worksheet
    Dim shName As Variant
    On Error GoTo Ender

    shName = GetSheetName(uid, wbSrc)
    If IsError(shName) Then Exit Function

    Set ResolveSrcSheet = wbSrc.Worksheets(CStr(Trim$(shName)))
    Exit Function
Ender:
    Set ResolveSrcSheet = Nothing
End Function
