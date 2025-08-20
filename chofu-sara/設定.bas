Option Explicit

' 出力先シート名
Public Const SHEET_NIPPO As String = "日報"
Public Const SHEET_URIAGE As String = "売上日報"

' ソースの対象シート検索用UID
Public Const NIPPO_UID As String = "Njk4X2Nob2Z1"
Public Const URIAGE_UID As String = "ODA0X2Nob2Z1"
Public Const DANSHI_KYU_DANSHI_UID As String = "NzA5X2Nob2Z1"
Public Const DANSHI_KYU_PART_UID As String = "ODAyX2Nob2Z1"
Public Const DANSHI_HIBARAI_DANSHI_UID As String = "NzEwX2Nob2Z1"
Public Const DANSHI_HIBARAI_PART_UID As String = "ODAzX2Nob2Z1"


' 既知のパスワードを返す（不明なら空文字）
Public Function GetSheetPassword(ByVal sheetName As String) As String
    Select Case sheetName
        Case SHEET_NIPPO
            GetSheetPassword = ""          ' わかっていればここに設定
        Case SHEET_URIAGE
            GetSheetPassword = ""          ' 同上
        Case Else
            GetSheetPassword = ""
    End Select
End Function
