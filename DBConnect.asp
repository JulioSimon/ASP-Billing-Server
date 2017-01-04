<%
Dim DBConnA, strConnectA
Set DBConnA = Server.CreateObject("ADODB.Connection")
strConnectA = "Provider=SQLOLEDB;Data Source=;Initial Catalog=;user ID=;password=;"
DBConnA.Open strConnectA

%>