<!-- #include file = "DBConnect.asp" -->
<!-- #include file = "Function.asp" -->
<!-- #include file = "Class_MD5.asp" -->
<%

On Error Resume Next

' Error
If Err.Number <> 0 Then
	Response.Write "ERROR"
	Response.End
End If



''' ACL access only from BillingServer
Dim IP
Dim BillingServer
BillingServer = "--private--"
IP = Request.ServerVariables("REMOTE_ADDR")
If(IP <> BillingServer) Then
	Response.Write "ACCESSDENY"
	Response.End
End If

''''''' Check MD5 secretstring 
Dim OrderNo
Dim Username
Dim Package
Dim Numcurrency
Dim ECash
Dim Secret

OrderNo = Trim(Request("OrderNo"))
OrderNo = FilterReqXSS(OrderNo)

Username = Trim(Request("Username"))
Username = FilterReqXSS(Username)

Package = Trim(Request("Package"))
Package = FilterReqXSS(Package)

Numcurrency = Trim(Request("Numcurrency"))
Numcurrency = FilterReqXSS(Numcurrency)

ECash = Trim(Request("ECash"))
ECash = FilterReqXSS(ECash)

SecretStr = Trim(Request("SecretStr"))
SecretStr = FilterReqXSS(SecretStr)

' Error - Parameter
Dim ParaResult
ParaResult = "OK"

If OrderNo = "" OR IsNull(OrderNo) OR IsEmpty(OrderNo) Then
	ParaResult = "PARA01"
End If
If Username = "" OR IsNull(Username) OR IsEmpty(Username) Then
	ParaResult = "PARA02"
End If
If Package = "" OR IsNull(Package) OR IsEmpty(Package) Then
	ParaResult = "PARA03"
End If
If Numcurrency = "" OR IsNull(Numcurrency) OR IsEmpty(Numcurrency) Then
	ParaResult = "PARA04"
End If
If ECash = "" OR IsNull(ECash) OR IsEmpty(ECash) Then
	ParaResult = "PARA05"
End If
If SecretStr = "" OR IsNull(SecretStr) OR IsEmpty(SecretStr) Then
	ParaResult = "PARA06"
End If

If ParaResult <> "OK" Then
	DBConnA.Close
	Set DBConnA = Nothing
	Response.Write ParaResult
	Response.End
End If

''''''CHECK MD5 Valid key String
Dim KeyString
Dim objMD5
Dim Confirm_Valid_Key

KeyString = "ENT-N2E-CGI"

Set objMD5 = New MD5
objMD5.Text = OrderNo & Username & Package & Numcurrency & ECash & IP & KeyString
Confirm_Valid_Key = objMD5.HEXMD5
' Error
If Err.Number <> 0 Then
	DBConnA.Close
	Set DBConnA = Nothing
	Response.Write "ERROR"
	Response.End
End If

If Trim(SecretStr) <> Trim(Confirm_Valid_Key) Then
	DBConnA.Close
	Set DBConnA = Nothing
	Response.Write "INVALID"
	Response.End
End If

'''''' Check User Exist
Dim sdSQL
Dim sdRS
Dim ReturnValue
sdSQL = "SET NOCOUNT ON  EXEC CGI.CGI_WebPurchaseCurrency '" & OrderNo & "'," & Username & "," & Package & "," & Numcurrency & "," & ECash & " "
' Response.Write sdSQL
Set sdRS = DBConnA.Execute(sdSQL)
ReturnValue = sdRS(0)
sdRS.Close
Set sdRS = Nothing
DBConnA.Close
Set DBConnA = Nothing
' Error
If Err.Number <> 0 Then
	Response.Write "ERROR"
	Response.End
End If
Response.Write ReturnValue
Response.End
%>