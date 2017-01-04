<!-- #include file = "DBConnect.asp" -->
<!-- #include file = "Function.asp" -->
<!-- #include file = "Class_MD5.asp" -->
<%

On Error Resume Next

' Error
If Err.Number <> 0 Then
	Response.Write "-1"
	Response.End
End If

''' ACL access only from BillingServer
Dim IP
Dim BillingServer
BillingServer = "--private--"
IP = Request.ServerVariables("REMOTE_ADDR")

Dim sdSQL
Dim sdRS
Dim ReturnValue

sdSQL = "SET NOCOUNT ON  EXEC CGI.CGI_WebGetTotalCurrency "

Set sdRS = DBConnA.Execute(sdSQL)
ReturnValue = sdRS(0)

sdRS.Close
Set sdRS = Nothing
DBConnA.Close
Set DBConnA = Nothing

' Error
If Err.Number <> 0 Then
	Response.Write "-1"
	Response.End
End If
Response.Write ReturnValue
Response.End
%>

