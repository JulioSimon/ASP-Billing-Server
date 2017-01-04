<%
Dim objHttp
Dim ServerCheck_SV1
Dim ServerCheck_DB1


''''''''''
On Error Resume Next
Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")

objHttp.open "POST", "http://127.0.0.1/hostcheck.asp", false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send

If Err.Number <> 0 Then
    If Err.Number >= 400 Then
        ServerCheck_SV1 = -101
    Else
        ServerCheck_SV1 = -102
    End If
Else
	ServerCheck_SV1 = 0
End If

Set objHttp = Nothing



''''''''''
On Error Resume Next
%>
<!-- #include file = "DBConnect.asp" -->
<%
If Err.Number <> 0 Then
    If Err.Number >= 400 Then
        ServerCheck_DB1 = -201
    Else
        ServerCheck_DB1 = -202
    End If
Else
	ServerCheck_DB1 = 0
End If

DBConnA.Close
Set DBConnA = Nothing



''''''''''
Dim ServerResult

ServerResult = Int(ServerCheck_SV1) + int(ServerCheck_DB1)

If ServerResult = 0 Then
	Response.Write "1"
Else
	Response.Write ServerResult
End If
%>