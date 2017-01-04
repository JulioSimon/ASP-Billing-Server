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



'''''''''' Req
Dim UserJID

UserJID = Trim(Request("JID"))

UserJID = FilterReqXSS(UserJID)


' Error
If UserJID = "" OR IsNull(UserJID) OR IsEmpty(UserJID) Then
	DBConnA.Close
	Set DBConnA = Nothing

	Response.Write "-2"
	Response.End
End If



''''''''''
Dim sdSQL
Dim sdRS
Dim ReturnValue
Dim currencyOwn
Dim currencyGift
Dim Mileage

sdSQL = "DECLARE @ReturnValue int "
sdSQL = sdSQL & "DECLARE @currencyOwn int "
sdSQL = sdSQL & "DECLARE @currencyGift int "
sdSQL = sdSQL & "DECLARE @Mileage int "
sdSQL = sdSQL & "EXEC @ReturnValue = _GetcurrencyDataForAppServer "& UserJID &", @currencyOwn OUTPUT, @currencyGift OUTPUT, @Mileage OUTPUT "
sdSQL = sdSQL & "SELECT @ReturnValue, @currencyOwn, @currencyGift, @Mileage"
Set sdRS = DBConnA.Execute(sdSQL)

ReturnValue = sdRS(0)
currencyOwn = sdRS(1)
currencyGift = sdRS(2)
Mileage = sdRS(3)

sdRS.Close
Set sdRS = Nothing
DBConnA.Close
Set DBConnA = Nothing

' Error
If Err.Number <> 0 Then
	Response.Write "-3"
	Response.End
End If

' return
If Cint(ReturnValue) <> 0 Then
	Response.Write "-4"
	Response.End
Else
	If currencyOwn = "" OR IsNull(currencyOwn) OR IsEmpty(currencyOwn) Then currencyOwn = 0
	If currencyGift = "" OR IsNull(currencyGift) OR IsEmpty(currencyGift) Then currencyGift = 0
	If Mileage = "" OR IsNull(Mileage) OR IsEmpty(Mileage) Then Mileage = 0
End If



''''''''''
Dim KeyString
Dim objMD5
Dim Valid_Key

KeyString = "--private--"

Set objMD5 = New MD5
objMD5.Text = UserJID & "." & currencyOwn & "." & currencyGift & "." & Mileage & "." & KeyString
Valid_Key = objMD5.HEXMD5

' Error
If Err.Number <> 0 Then
	Response.Write "-5"
	Response.End
End If



''''''''''
Response.Write "1:"& currencyOwn &","& currencyGift &","& Mileage &","& Valid_Key
Response.End
%>