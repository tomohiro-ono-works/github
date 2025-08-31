' ==========================
' Module: credentials.bas
' ==========================
Option Explicit

Public Const BACKLOG_SPACE_URL As String = "**********"  ' 例: https://example.backlog.com
Public Const BACKLOG_PROJECT_KEY As String = "**********" ' 例: APP
Public Const BACKLOG_API_KEY As String = ""              ' あなたのBacklog APIキー

Public Const TASK_FOLDER_NAME As String = "********"     ' Outlookタスク配下のサブフォルダ
Public Const SUMMARY_MAX As Long = 20                    ' 件名丸め（20文字）

Public Function DoneStatusNames() As Variant
    DoneStatusNames = Array("完了", "Closed", "Done", "Resolved")
End Function