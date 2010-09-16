<%Option Explicit%>
<%Response.Buffer = TRUE%>
<!--#Include File = "include.asp"-->
<%
Dim TableName, ThreadID

Set TableName = Request.QueryString("table")
Set ThreadID = Request.QueryString("id")

Dim TheCookieName

Dim UserName

UserName = Request.Cookies("NightFire")("UserName")


Call OpenConnection()

Dim RSGI

Set RSGI = Server.CreateObject("ADODB.Recordset")
RSGI.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserName&"'", Conn

Dim System, Minister
Minister = RSGI("Minister")
System = RSGI("SystemNumber")


Dim SQL, SQL2, RS, RS2

SQL = "SELECT * FROM MEETING WHERE SubjectID = "&ThreadID&" AND Original = 2 AND System = "&System&" ORDER BY DatePosted"
SQL2 = "SELECT * FROM MEETING WHERE SubjectID = "&ThreadID&" AND Original = 1 AND System = "&System
'Response.Write SQL
Set RS = Server.CreateObject("ADODB.Recordset")
Set RS2 = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic
RS2.Open SQL2, Conn, 3, 3
Dim BorderNum, BorderColor, BackColor

Dim Subject, Poster, Message, TimePosted

Subject = RS2("Subject")

TimePosted = RS2("DatePosted")
Poster = RS2("Poster")
Message = RS2("Message")
Message = Replace(Message, vbCrLf, "<br>" & vbCrLf)


RS2.Close
If BorderNum = 2 Then
	BorderColor = "black"
	BackColor = "#000066"
	BorderNum = 1
Else
	BorderColor = "white"
	BackColor = "#000033"
	BorderNum = 2
End If
%>
<html>
<head>
<title>NightFire Meeting: <%=Subject%></title>
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
<body bgcolor="#000000" text="#ffffff" link="#ffffff" vlink="#ffffff" alink="ffcc33" MARGINWIDTH=0 MARGINHEIGHT=0 TOPMARGIN=0 BOTTOMMARGIN=0 LEFTMARGIN=0 RIGHTMARGIN=0>
<center>
<img src = "banner.gif" align = "center">
<p class = "large"><%=MyTitle%>:<%=Subject%></p></center>
<table border="1" cellpadding="0" cellspacing="0" align = "center">
<center>
<tr bordercolor = "<%=BorderColor%>" bgcolor = "<%=BackColor%>">
<td><p class = "small">Last Post at <%=TimePosted%> </p><br>
<p><%=Message%></p><br>
<p>-<%=Poster%></p>
</td></tr>
<%
If RS.EOF = FALSE Then
RS.MoveFirst
Do While Not RS.EOF
TimePosted = RS("DatePosted")
Poster = RS("Poster")
Message = RS("Message")
Message = Replace(Message, vbCrLf, "<br>" & vbCrLf)



If BorderNum = 2 Then
	BorderColor = "black"
	BackColor = "#000066"
	BorderNum = 1
Else
	BorderColor = "white"
	BackColor = "#000033"
	BorderNum = 2
End If
%>
<tr bordercolor = "<%=BorderColor%>" bgcolor = "<%=BackColor%>">
<td><p class = "small">Posted at <%=TimePosted%> </p><br>
<p><%=Message%></p><br>
<p>-<%=Poster%></p>
</td></tr>
<%
RS.MoveNext
Loop
End If
RS.Close
Conn.Close

%>
</table>

</table>
<a href = "meeting.asp?table=<%=TableName%>
<%
If Request.Querystring("admin") = AdminPass Then
%>
&admin=<%=Request.Querystring("admin")%>
<%
End If
%>
">Back to <%=MyTitle%></a>
			<form action="postmeeting.asp?table=<%=TableName%>
<%
If Request.Querystring("admin") = AdminPass Then
%>
&admin=<%=Request.Querystring("admin")%>
<%
End If
%>
" method="post">
			<table border="0" cellpadding="3" cellspacing="0" width="270" align = "center">
				<caption><p>Post a New Message</p></caption>
				<tr>
					<td align="left" width="40%"><p>Dictator:</p></td>
					<td align="left" width="60%"><%=Minister%></td>
				</tr>
<%
If AuthOn = TRUE Then
%>
				<tr>
					<td align="left" width="40%"><p>Password:</p></td>
					<td align="left" width="60%"><p><input name="Password" size="10" maxlength="20" type="password" value = "<%=TheCookiePass%>"  class = "numbar"> &nbsp;<input type = "checkbox" name = "savepass" value = "yes" class = "button"> Save</p></td>
				</tr>
<%
End If
%>
				<tr>
					<td align="left" width="40%"><p>Subject:</p></td>
					<td align="left" width="60%"><p><%=Subject%></p></td>
				</tr>
			</table>
   <div align = "center">
			<textarea name = "message" cols = "35" rows = "6"  class = "numbar"></textarea><br>
			<input type = "hidden" name = "Original" value = "2">
			<input type = "hidden" name = "Subject" value = "<%=Subject%>">
			<input type = "hidden" name = "ThreadID" value = "<%=ThreadID%>">
			<input type = "submit" value = "Post" class = "button" align = "center">
   </div>
			</p>
			</form>
</center>
</html>