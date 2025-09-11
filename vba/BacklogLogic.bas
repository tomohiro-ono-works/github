' ==========================
' Module: BacklogLogic.bas
' ==========================
Option Explicit

'==========================================
' メイン処理: Backlogのチケットをタスクに登録
'==========================================
Public Sub FetchBacklogTasks()

    Dim url As String
    url = BuildIssuesUrl()   '設定モジュールの関数を呼び出す

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "APIエラー: " & http.Status & " " & http.statusText, vbCritical
        Exit Sub
    End If

    Dim issues As Object
    Set issues = JsonConverter.ParseJson(http.responseText)

    Dim issue As Variant
    For Each issue In issues
        Dim key As String
        key = issue("issueKey")

        If Not TaskExists(key) Then
            Dim t As Outlook.TaskItem
            Set t = Application.CreateItem(olTaskItem)
            t.Subject = key & " " & issue("summary")
            If Not issue("description") Is Nothing Then
                t.Body = issue("description")
            End If
            If Not issue("dueDate") Is Nothing Then
                t.DueDate = CDate(issue("dueDate"))
            End If
            t.Save
        End If
    Next

    MsgBox "Backlogのチケットをタスクに登録しました。", vbInformation
End Sub

'==========================================
' issueKeyを含むタスクが既存か確認
'==========================================
Private Function TaskExists(issueKey As String) As Boolean
    Dim ns As Outlook.Namespace
    Set ns = Application.GetNamespace("MAPI")

    Dim items As Outlook.Items
    Set items = ns.GetDefaultFolder(olFolderTasks).Items

    Dim itm As Object
    For Each itm In items
        If InStr(itm.Subject, issueKey) > 0 Then
            TaskExists = True
            Exit Function
        End If
    Next
    TaskExists = False
End Function

