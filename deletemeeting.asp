<%Option Explicit%>
<%Response.Buffer = TRUE%>
<!--#Include File = "include.asp"-->
<html>
	<head>
		<title>Welcome To <%=MyTitle%></title>
		<link rel="stylesheet" type="text/css" href="game.css">
	</head>
	<body bgcolor="#000000" text="#ffffff" link="#ff6633" vlink="#ff6633" alink="ffcc33">

		

		<center>
			<img src = "banner.gif" align = "center">
<%

Dim TableName, i, x

Set TableName = Request.QueryString("table")

Dim Thread, NumThread, DeleteThread
NumThread = Request.Form("DeleteNum")
NumThread = NumThread

Dim SQL
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

Response.Write "<p>" & MyDelete & "</p>"
%>
			<a href = "meeting.asp?table=<%=TableName%>
<%
If Request.Querystring("admin") = AdminPass Then
%>
&admin=<%=Request.Querystring("admin")%>"
<%
End If
%>
> Return to Sytstem Meeting </a>
		</center>
	</body>
</html>