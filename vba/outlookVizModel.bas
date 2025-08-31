' ==========================
' Module: outlookVizModel.bas
' ==========================
Option Explicit

' 旧UI互換の CommandBars を使って、エクスプローラにボタンを追加します。
' リボンの“新しいタブ”を完全自動追加することはVBA単体では不可のため、
' 手動カスタマイズと併用してください。（この関数は補助的な簡易ボタン追加です）

Private Const msoControlButton As Long = 1
Private Const msoButtonIconAndCaption As Long = 3

Public Sub InstallToolbarButton()
    On Error GoTo EH
    Dim exp As Outlook.Explorer
    Set exp = Application.ActiveExplorer
    If exp Is Nothing Then
        MsgBox "Outlook のメインウィンドウを開いてから実行してください。", vbExclamation
        Exit Sub
    End If
    Dim cbs As Office.CommandBars
    Set cbs = exp.CommandBars
    Dim cb As Office.CommandBar
    Set cb = cbs("Standard") ' 見つからない場合はClassic UIが無い可能性

    Dim btn As Office.CommandBarButton
    Set btn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Backlog 同期"
    btn.Style = msoButtonIconAndCaption
    btn.OnAction = "issueController.SyncBacklogToOutlook"
    MsgBox "ツールバーにボタンを追加しました。", vbInformation
    Exit Sub
EH:
    MsgBox "自動追加に失敗しました。Outlookの[リボンのユーザー設定]から、" & _
           "マクロ 'issueController.SyncBacklogToOutlook' を手動で追加してください。", vbExclamation
End Sub


' ==========================
' ThisOutlookSession（参考：イベントに紐づけたい場合）
' ==========================
'Option Explicit
'
'Private Sub Application_Startup()
'    ' 起動時にボタン追加を試みる（失敗しても無視）
'    On Error Resume Next
'    outlookVizModel.InstallToolbarButton
'    On Error GoTo 0
'End Sub
