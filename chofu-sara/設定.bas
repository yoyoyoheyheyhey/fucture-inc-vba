Option Explicit

' 出力先シート名
Public Const SHEET_NIPPO As String = "日報"
Public Const SHEET_URIAGE_NIPPO As String = "売上日報"
Public Const SHEET_DANSHI_HIBARAI As String = "男子日払い"


' ソースの対象シート検索用UID(カスタムレポート)
Public Const NIPPO_UID As String = "Njk4X2Nob2Z1" ' 日報 > 社交 | 日報
Public Const URIAGE_NIPPO_UID As String = "MjUyX2Nob2Z1" ' 日報 > 売上日報
Public Const DANSHI_HIBARAI_DANSHI_UID As String = "NzEwX2Nob2Z1" ' 男子日払い > 男子
Public Const DANSHI_HIBARAI_PART_UID As String = "ODAzX2Nob2Z1" ' 男子日払い > アルバイト

' ソースの対象シート名（経費はリソースWBが別）
Public Const SHEET_KEIHI As String = "経費"

' 既知のパスワードを返す（不明なら空文字）
Public Function GetSheetPassword(ByVal sheetName As String) As String
    Select Case sheetName
        Case SHEET_NIPPO
            GetSheetPassword = ""          ' わかっていればここに設定
        Case SHEET_URIAGE_NIPPO
            GetSheetPassword = ""          ' 同上
        Case SHEET_DANSHI_HIBARAI
            GetSheetPassword = ""          ' 同上
        Case Else
            GetSheetPassword = ""
    End Select
End Function
