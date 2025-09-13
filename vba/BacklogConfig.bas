' ==========================
' Module: BacklogConfig.bas
' ==========================
Option Explicit

'============================
' Backlog 接続設定
'============================
Public Const API_KEY As String = "YOUR_API_KEY"
Public Const SPACE_KEY As String = "your_space"
Public Const PROJECT_ID As String = "123456"   '対象プロジェクトID

'============================
' API URL生成
'============================
Public Function BuildIssuesUrl() As String
    '完了(4) と 未着手(1) 以外 → 処理中(2), 処理済み(3) を取得
    Dim baseUrl As String
    baseUrl = "https://" & SPACE_KEY & ".backlog.jp/api/v2/issues"

    BuildIssuesUrl = baseUrl & "?apiKey=" & API_KEY _
                     & "&projectId[]=" & PROJECT_ID _
                     & "&statusId[]=2&statusId[]=3"
End Function

