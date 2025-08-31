' ==========================
' Module: todoModel.bas
' ==========================
Option Explicit

Public Function GetOrCreateTaskSubFolder(ByVal parent As Outlook.MAPIFolder, ByVal name As String) As Outlook.MAPIFolder
    Dim f As Outlook.MAPIFolder
    For Each f In parent.Folders
        If f.Name = name Then Set GetOrCreateTaskSubFolder = f: Exit Function
    Next
    Set GetOrCreateTaskSubFolder = parent.Folders.Add(name, olFolderTasks)
End Function

Public Function BuildExistingIndex(ByVal folder As Outlook.MAPIFolder) As Scripting.Dictionary
    Dim idx As New Scripting.Dictionary
    Dim it As Object, subj As String, body As String
    For Each it In folder.Items
        On Error Resume Next
        subj = NzStr(it.Subject)
        body = NzStr(it.HTMLBody)
        If body = "" Then body = NzStr(it.Body)
        On Error GoTo 0
        If subj <> "" Then
            If Not idx.Exists(subj) Then idx.Add subj, New Collection
            idx(subj).Add body
        End If
    Next
    Set BuildExistingIndex = idx
End Function

Public Sub EnsureCategoriesExist(ByVal issues As Collection)
    On Error Resume Next
    Dim cats As Outlook.Categories
    Set cats = Application.Session.Categories
    Dim exist As New Scripting.Dictionary

    Dim i As Long
    For i = 1 To cats.Count
        exist(cats.Item(i).Name) = True
    Next

    Dim it As TIssue
    For i = 1 To issues.Count
        it = issues(i)
        If it.IssueTypeName <> "" Then
            If Not exist.Exists(it.IssueTypeName) Then
                cats.Add it.IssueTypeName ' 既定色
                exist(it.IssueTypeName) = True
            End If
        End If
    Next
    On Error GoTo 0
End Sub

Public Function HasDuplicate(ByVal idx As Scripting.Dictionary, ByVal subject As String, ByVal issueUrl As String) As Boolean
    If Not idx.Exists(subject) Then Exit Function
    Dim bodies As Collection: Set bodies = idx(subject)
    Dim b As Variant
    For Each b In bodies
        If InStr(1, CStr(b), issueUrl, vbTextCompare) > 0 Then HasDuplicate = True: Exit Function
    Next
End Function

Public Sub CreateTask( _
    ByVal folder As Outlook.MAPIFolder, _
    ByVal subject As String, _
    ByVal bodyHtml As String, _
    ByVal dueYmd As String, _
    ByVal categoryName As String)

    Dim task As Outlook.TaskItem
    Set task = folder.Items.Add(olTaskItem)
    task.Subject = subject
    If bodyHtml <> "" Then task.HTMLBody = bodyHtml
    If dueYmd <> "" Then On Error Resume Next: task.DueDate = CDate(dueYmd): On Error GoTo 0
    If categoryName <> "" Then task.Categories = categoryName
    task.Save
End Sub