<% Option Explicit %>

<!--#Include File = "adovbs.inc"-->
<html>
<head>
<title> Database Manager </title>
<script language="javascript">
<!--//

	if (navigator.appName == "Netscape") {
		document.write('<link rel=stylesheet href="gameNS.css" type="text/css">');
	} else {
		document.write('<link rel=stylesheet href="game.css" type="text/css">');
	}
	
//-->	
</script>
</head>
<body bgcolor="#000000" text="#ffffff" link="#ffffff" vlink="#ffffff" alink="ffcc33">
<center>

<%
If Request.QueryString("pass") <> "BrakIsGod" Then
   Response.Write "Dumbass... can't even guess the password... jeez."
Else

If Request("execute") = 1 Then

Dim Conn, OpenConn
Sub OpenConnection()
	If OpenConn <> TRUE Then
		OpenConn = TRUE
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.ConnectionString="Provider=SQLOLEDB;User ID=sa;Password=Kt6261;Initial Catalog=warpspire"
		Conn.Open
	End If
End Sub

Call OpenConnection()

Dim SQL
'SQL = "CREATE TABLE Gen"
'SQL = SQL & "(Poster TEXT(30), Subject TEXT(30), Message TEXT(255), Original INTEGER,  DatePosted DATE, SubjectID INTEGER, Posts INTEGER)"

SQL = Request.Form("SQL")
Conn.Execute(SQL)

Response.Write "<p class = 'normal'>SQL Query Completed</p>"

If Request("rs") = "yes" Then

Response.Write "<table><tr bgcolor='#ee5900'>"
Dim RS

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, Conn, adOpenStatic, adLockPessimistic

Dim i, x

i = RS.Fields.Count
x = 0

Do While i > 0

Response.Write "<td><div>" & RS(x).Name & "</div></td>" 
i = i - 1
x = x + 1
Loop

Response.Write "</tr>"

i = RS.Fields.Count
x = 0
Dim y
y = 1
If RS.EOF = FALSE Then
Do While RS.EOF = FALSE
If y = 1 Then
Response.Write "<tr bgcolor = '#000066'>"
y =2
Else
Response.Write "<tr bgcolor = '#000033'>"
y = 1
End If

i = RS.Fields.Count
i = i - 1
x = 0

Do While i > -1

Response.Write "<td>" & RS(x) & "</td>"

i = i - 1
x = x + 1

Loop
Response.Write "</tr>"
RS.MoveNext

Loop



End If


Response.Write "</table>"
End If

Set Conn = nothing
Else
End If



If Request.Form("execute") = 2 Then
Dim SelectedTable
SelectedTable = Request.Form("table")
End If
%>

<form action = "sql.asp?pass=<%=Request.Querystring("pass")%>" method = "POST">
<p>SQL Query</p><br><textarea name = "SQL" cols = "35" rows = "6" class = "numbar"><%=Request("SQL")%></textarea><br>
<input type = "hidden" value = "1" name = "execute">
<p>Return RS: <input type = "radio" name = "rs" value = "yes" class = "button"> &nbsp;&nbsp;&nbsp;&nbsp; <input type = "submit" value = "Execute SQL" class = "button"></p>
</form>
<br>
<br>
<br>
<p>Database Editor</p>
<br>
<form action = "sql.asp" method=post>
<input type=hidden name=execute value=2>
<table width="40%">
<tr>
<td>Table Name:</td>
<td><input type=text name="table" value="<%=SelectedTable%>" size=15 class="numbar"></td>
</tr>
</table>
<%
If Request.Form("execute") = 2 Then
Call OpenConnection()
Dim RSDE, FieldName
Set RSDE = Server.CreateObject("ADODB.Recordset")
RSDE.Open SelectedTable, Conn, adOpenStatic, adLockPessimistic
If Request.Form("addrecord") = "TRUE" Then
i = RSDE.Fields.Count
x = 0
RSDE.AddNew

Do While i > 0
FieldName = RSDE(x).Name
RSDE(x).Value = Request.Form(FieldName)

i = i - 1
x = x + 1

Loop
RSDE.Update
End If

Response.Write "<table><tr  bgcolor='#ee5900'>"
i = RSDE.Fields.Count
x = 0

Do While i > 0

Response.Write "<td><div align=center>" & RSDE(x).Name & "</div></td>" 
i = i - 1
x = x + 1
Loop
Response.Write "</tr><tr>"

i = RSDE.Fields.Count
x = 0

Do While i > 0
Response.Write "<td><input type=text class='numbar' name='"&RSDE(x).Name&"'></td>"
i = i - 1
x = x + 1
Loop
Response.Write "</tr></table><input type=hidden name='addrecord' value='TRUE'>"
End If
End If
%>
<table width="40%">
<tr>
<td><input type=button value = "Edit Record" class="button"></td>
<td><input type=submit value = "Add Record" class="button"></td>
</tr>
</table>
<br><a href = "index.html">Go Back</a>
</center>
</body>
</html>

