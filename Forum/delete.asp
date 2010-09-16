<%Option Explicit%>
<%Response.Buffer = TRUE%>
<!--#Include File = "include.asp"-->
<html>
	<head>
		<title>Welcome To <%=MyTitle%></title>
		<link rel="stylesheet" type="text/css" href="forum.css">
	</head>

	<body>
		

		<center>
			<img src = "banner.gif" align = "center">
<%

Dim TableName, i, x

Set TableName = Request.QueryString("table")

Dim Thread, NumThread, DeleteThread, MsgID
NumThread = Request.Form("DeleteNum")
NumThread = NumThread
MsgID = Request.Form("MessageID")

Dim SQL



If Request.Form("DeleteMsg") = "Yes" Then
   SQL = "Delete * FROM "&TableName&" WHERE MsgID = "&Request.Form("PostID")
   Conn.Execute(SQL)
Else
		SQL = "DELETE * FROM "&TableName&" WHERE SubjectID = "
For x = 0 to NumThread
	Thread = Request.Form("Delete" & x)
	If Thread <> "" Then
		i = i + 1
		If i = 1 Then
			SQL = SQL & Thread
		Else
			SQL = SQL & " OR SubjectID = " & Thread
		End If		
	End If
Next
Conn.Execute(SQL)
End If
Response.Write "<p>" & MyDelete & "</p>"
%>
			<a href = "forum.asp?table=<%=TableName%>&admin=<%=Request.Querystring("admin")%>"> Return to Forums </a>
		</center>
	</body>
</html>