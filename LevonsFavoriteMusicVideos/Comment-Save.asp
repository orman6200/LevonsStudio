<%
Option Explicit
Response.ContentType = "application/json"

Dim comment, filePath, fso, file
comment = Request.Form("comment")
If Len(Trim(comment)) = 0 Then
    Response.Write "{""success"":false,""error"":""Comment is required.""}"
    Response.End
End If

comment = Replace(comment, Chr(13), " ")
comment = Replace(comment, Chr(10), " ")
comment = Replace(comment, """", """")

filePath = Server.MapPath("/Levon'sFavoriteMusicVideos/Comments.csv")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(filePath) Then
    Set file = fso.CreateTextFile(filePath, True, False)
    file.Close
End If

Set file = fso.OpenTextFile(filePath, 8, True, 0)
file.WriteLine """" & comment & """"
file.Close

Set file = Nothing
Set fso = Nothing

Response.Write "{""success"":true}"
%>