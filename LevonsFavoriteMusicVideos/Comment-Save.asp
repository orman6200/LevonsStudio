<%
Option Explicit

Response.ContentType = "application/json"
Response.AddHeader "Cache-Control", "no-store"

Function CsvEscape(value)
    If IsNull(value) Then value = ""
    CsvEscape = Replace(CStr(value), """", """""")
End Function

Dim name, text, timestamp, comment, filePath, fso, fileObj, existingContent, csvLine
name = Trim(Request.Form("name"))
text = Trim(Request.Form("text"))
timestamp = Trim(Request.Form("timestamp"))
comment = Trim(Request.Form("comment"))

If timestamp = "" Then
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & "  " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
End If

If name = "" Or text = "" Then
    Response.Write "{""success"":false,""error"":""Name and Comment are required.""}"
    Response.End
End If

csvLine = """" & CsvEscape(name) & """,""" & CsvEscape(text) & """,""" & CsvEscape(timestamp) & """"

filePath = Server.MapPath("Comments.csv")
Set fso = Server.CreateObject("Scripting.FileSystemObject")

If fso.FileExists(filePath) Then
    Set fileObj = fso.OpenTextFile(filePath, 1, True)
    If Not fileObj.AtEndOfStream Then
        existingContent = fileObj.ReadAll
    Else
        existingContent = ""
    End If
    fileObj.Close
Else
    existingContent = ""
End If

' Prepend the new line to the existing content instead of appending it
existingContent = csvLine & vbCrLf & existingContent

Set fileObj = fso.OpenTextFile(filePath, 2, True)
fileObj.Write existingContent
fileObj.Close

Response.Write "{""success"":true}"
%>