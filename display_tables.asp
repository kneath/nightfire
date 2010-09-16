<html>
<head>
<title>Database Explorer</title>
</head>
<body>
<!--#Include File="adovbs.inc"-->
<%
Dim Conn, TableRS, ColumnRS, Database
If Request.Form("Database") = "" Then
	Database = "mydb.mdb"
Else
	Database = Request.Form("Database")
End If

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString="Provider=SQLOLEDB;User ID=sa;Password=Kt6261;Initial Catalog=warpspire"
Conn.Open

Set TableRS = Conn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))

Do While Not TableRS.EOF
	Response.Write "--><b>" & TableRS("Table_Name").Value & "</b><br>"
	Set ColumnRS = Conn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TableRS("Table_Name").Value))
	Do While Not ColumnRS.EOF
		Response.Write "---->" &  ColumnRS("Column_Name") & " (" & ColumnRS("Data_Type") & ")<br>"
		ColumnRS.MoveNext
	Loop
	Response.Write "<br><br><br>"
	Set ColumnRS = Nothing
	TableRS.MoveNext
Loop

Set TableRS = Nothing
Set Conn = Nothing
%>
<form action = "tables.asp" method = "post">
<input type="text" name = "Database" value="<%=Databse%>">
<input type="submit">
</form>
	
