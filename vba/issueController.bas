' ==========================
' Module: issueController.bas
' ==========================
Option Explicit

Public Sub SyncBacklogToOutlook()
    On Error GoTo EH
    Dim issues As Collection
    Dim created As Long, skipped As Long

    If BACKLOG_API_KEY = "" Then Err.Raise vbObjectError + 1, , "BACKLOG_API_KEY を設定してください。"

    Set issues = ListIssuesParentOnlyNotDone(BACKLOG_SPACE_URL, BACKLOG_PROJECT_KEY, BACKLOG_API_KEY)

    Dim olApp As Outlook.Application
    Dim ns As Outlook.NameSpace
    Dim tasksRoot As Outlook.MAPIFolder
    Dim targetFolder As Outlook.MAPIFolder

    Set olApp = Outlook.Application
    Set ns = olApp.Session
    Set tasksRoot = ns.GetDefaultFolder(olFolderTasks)
    Set targetFolder = GetOrCreateTaskSubFolder(tasksRoot, TASK_FOLDER_NAME)

    Dim existing As Scripting.Dictionary
    Set existing = BuildExistingIndex(targetFolder)

    EnsureCategoriesExist issues

    Dim i As Long
    For i = 1 To issues.Count
        Dim it As TIssue
        it = issues(i)

        Dim subject As String
        subject = it.IssueKey & " " & Left$(it.Summary, SUMMARY_MAX)

        Dim issueUrl As String
        issueUrl = BuildIssueUrl(BACKLOG_SPACE_URL, it.IssueKey)

        Dim bodyHtml As String
        bodyHtml = "<p><a href='" & issueUrl & "'>" & issueUrl & "</a></p>"

        If HasDuplicate(existing, subject, issueUrl) Then
            skipped = skipped + 1
        Else
            CreateTask targetFolder, subject, bodyHtml, it.DueDate, it.IssueTypeName
            If Not existing.Exists(subject) Then existing.Add subject, New Collection
            existing(subject).Add bodyHtml
            created = created + 1
        End If
    Next

    MsgBox "作成: " & created & vbCrLf & _
           "スキップ(重複): " & skipped, vbInformation
    Exit Sub
EH:
    MsgBox "エラー: " & Err.Number & vbCrLf & Err.Description, vbExclamation
End Sub