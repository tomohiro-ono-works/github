' ==========================
' Module: backlogModel.bas
' ==========================
Option Explicit

Public Type TIssue
    IssueKey As String
    Summary As String
    DueDate As String         ' "YYYY-MM-DD" or ""
    IssueTypeName As String
End Type

Public Function ListIssuesParentOnlyNotDone( _
    ByVal spaceUrl As String, _
    ByVal projectKey As String, _
    ByVal apiKey As String _
) As Collection
    Dim pid As Long: pid = GetProjectId(spaceUrl, projectKey, apiKey)
    Dim notDoneIds As Collection: Set notDoneIds = GetNotDoneStatusIds(spaceUrl, projectKey, apiKey)

    Dim result As New Collection
    Dim offset As Long: offset = 0
    Dim pageSize As Long: pageSize = 100

    Do
        Dim qs As String
        Dim idVar As Variant
        qs = "?apiKey=" & apiKey & _
             "&projectId[]=" & pid & _
             "&parentChild=1" & _
             "&count=" & pageSize & _
             "&offset=" & offset
        For Each idVar In notDoneIds
            qs = qs & "&statusId[]=" & CLng(idVar)
        Next
        
        Dim url As String
        url = spaceUrl & "/api/v2/issues" & qs

        Dim arr As Variant
        arr = HttpGetJson(url)
        If TypeName(arr) <> "Collection" Then Exit Do
        If arr.Count = 0 Then Exit Do

        Dim item As Variant
        For Each item In arr
            Dim it As TIssue
            it.IssueKey = NzStr(item("issueKey"))
            it.Summary = NzStr(item("summary"))
            it.DueDate = NzStr(item("dueDate"))
            If Not item("issueType") Is Nothing Then
                it.IssueTypeName = NzStr(item("issueType")("name"))
            Else
                it.IssueTypeName = ""
            End If
            result.Add it
        Next

        offset = offset + pageSize
        If arr.Count < pageSize Then Exit Do
    Loop

    Set ListIssuesParentOnlyNotDone = result
End Function

Private Function GetProjectId(ByVal spaceUrl As String, ByVal projectKey As String, ByVal apiKey As String) As Long
    Dim url As String
    url = spaceUrl & "/api/v2/projects/" & projectKey & "?apiKey=" & apiKey
    Dim obj As Variant
    obj = HttpGetJson(url)
    GetProjectId = CLng(obj("id"))
End Function

Private Function GetNotDoneStatusIds(ByVal spaceUrl As String, ByVal projectKey As String, ByVal apiKey As String) As Collection
    Dim url As String
    url = spaceUrl & "/api/v2/projects/" & projectKey & "/statuses?apiKey=" & apiKey
    Dim arr As Variant
    arr = HttpGetJson(url)

    Dim col As New Collection
    Dim it As Variant, nm As String, v As Variant, isDone As Boolean
    For Each it In arr
        nm = NzStr(it("name"))
        isDone = False
        For Each v In DoneStatusNames()
            If StrComp(CStr(v), nm, vbTextCompare) = 0 Then isDone = True: Exit For
        Next
        If Not isDone Then
            If InStr(1, nm, "完了", vbTextCompare) > 0 Then
                isDone = True
            End If
        End If
        If Not isDone Then col.Add CLng(it("id"))
    Next
    Set GetNotDoneStatusIds = col
End Function

Public Function BuildIssueUrl(ByVal spaceUrl As String, ByVal issueKey As String) As String
    BuildIssueUrl = spaceUrl & "/view/" & issueKey
End Function

' --- HTTP / JSON ---
Public Function HttpGetJson(ByVal url As String) As Variant
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", url, False
    xhr.setRequestHeader "Accept", "application/json"
    xhr.send
    If xhr.Status < 200 Or xhr.Status >= 300 Then
        Err.Raise vbObjectError + 2000, , "HTTP " & xhr.Status & ": " & xhr.responseText
    End If
    HttpGetJson = JsonConverter.ParseJson(xhr.responseText)
End Function

Public Function NzStr(ByVal v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then NzStr = "" Else NzStr = CStr(v)
    ElseIf IsEmpty(v) Or IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function