<%Option Explicit%>
<%Response.Buffer = TRUE%>
<!--#Include File = "include.asp"-->
<html>
	<head>
		<title>Welcome To <%=MyTitle%></title>
		<link rel="stylesheet" type="text/css" href="forum.css">
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

	<body bgcolor="#000000" text="#ffffff" link="#ff6633" vlink="#ff6633" alink="ffcc33">
		

		<center>
	<!--#Include File = "ad.asp"-->
	<!--#Include File = "statbar.asp"-->
			<table border="0" cellpadding="3" cellspacing="0" width="100%">
				<caption bgcolor="000080">
					<p class="large"><%=MyTitle%><br></p>
				
<%
Dim TheCookieName, TheCookiePass, UserName

UserName = Request.Cookies("NightFire")("UserName")


Dim TableName

TableName = "MEETING"
Response.Write "<form action = 'deletemeeting.asp?table="&TableName&"&admin="&Request.Querystring("admin")&"' method = 'post'>"

Dim SQL, RS, RSInfo, System

'Just a reminder here: TableNames = Announcements, Gen, Strategy, Politics, BugsNErrors, Suggestions. Fields in the Tables: Poster, Subject, Message, Original, DatePosted, SubjectID, Posts.
Set RSInfo = Server.CreateObject("ADODB.RecordSet")

RSInfo.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
'Now all connections have been made

If RSInfo("King") = "Yes" and Request.QueryString("admin") <> AdminPass Then
	Response.Redirect "meeting.asp?admin=NightFireKingShip"
End If
Dim Ruler
Ruler = RSInfo("Minister")
Ruler = Ruler & " of " &RSInfo("Nation")
System = RSInfo("SystemNumber")

SQL = "SELECT * FROM "&TableName&" WHERE Original = 1 AND System = "&System&" ORDER BY DatePosted DESC"
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.Open SQL, Conn, adOpenStatic, adLockPessimistic
'Now all connections have been made

Dim WelcomeMessage
If RS.EOF Then
	WelcomeMessage = "<p align = 'center'>" & MyWelcomeNew & "</p>"
	Response.Write WelcomeMessage
	Response.Write "</caption><tr bgcolor='#000066'>"
Else
	WelcomeMessage = "<p align = 'center'>" & MyWelcomeOld & "</p>"
	Response.Write WelcomeMessage

Response.Write "</caption><tr bgcolor = '#0000CC'>"
Dim admin
admin = Request.Querystring("admin")
If admin = AdminPass Then
admin = "yes"
Response.Write "<td><div>Delete</div></td>"
End If
%>
					<td width="40%">
						<div>
							Subject</div>
					</td>
					<td width="10%">
						<div>
							Posts</div>
					</td>
					<td width="20%">
						<div>
							Poster</div>
					</td>
					<td width="30%">
						<div>
							Last Post</div>
					</td>
				</tr>
				<%
	Dim Poster, Subject, Posts, TimePosted, ColorNum, Color, ThreadID, pos
	
	Color = 1
	ColorNum = 0

	Do While Not RS.EOF
		Poster = RS("Poster")
		pos = InStr(Poster, " of ")
		If pos = 0 Then
			pos = 100
		End If
		Poster = Mid(Poster, 1, pos)
		Subject = RS("Subject")
		Posts = RS("Posts")
		TimePosted = RS("DatePosted")
		ThreadID = RS("SubjectID")
		
		If ColorNum/2 = ColorNum\2 Then
			Color = "somecolor here"
		Else
			Color = "someother color here"
		End If
		Response.Write "<tr>"
		If admin = "yes" Then
		Response.Write "<td bgcolor = '#000066'><div><input type = 'checkbox' value = '"&ThreadID&"' name = 'Delete" & ColorNum & "' class = 'button'></td>"
		End If
		Response.Write "<td bgcolor = '#000033'><div><a href = 'threadmeeting.asp?table="&TableName&"&id="&ThreadID&"&Posts="&Posts
		If Request.QueryString("admin") = AdminPass Then
			Response.Write "&admin="&Request.Querystring("admin")
		End If
		Response.Write "'>" & VbCrlf &Subject&"</a></div></td>" & VbCrlf
		Response.Write "<td bgcolor = '#000066'><div>"&Posts&"</div></td>" & VbCrlf
		Response.Write "<td bgcolor = '#000033'><div>"&Poster&"</div></td>" & VbCrlf
		Response.Write "<td bgcolor = '#000066'><div>"&TimePosted&"</div></td></tr>" & VbCrlf
		
		ColorNum = ColorNum + 1
		RS.MoveNext
	Loop
End If
		
Response.Write "</table>"
If admin = "yes" then
Response.Write "<input type = 'hidden' name = 'deletenum' value = '" & ColorNum & "'>" & vbCrlf & "<input type = 'submit' value = 'Delete Checked Threads' class = 'button'>"
End If
	%>
			
			</form>
			<form action="postmeeting.asp?table=<%=TableName%>
<%
If Request.Querystring("admin") = AdminPass Then
%>
&admin=<%=Request.Querystring("admin")%>
<%
End If
%>
" method="post">
			<table border="0" cellpadding="3" cellspacing="0" width="190">
				<caption><p>Post a New Message</p></caption>
				<tr>
					<td align="left" width="50%"><p>Dictator:</p></td>
					<td align="left" width="50%"><input type = "hidden" name = "Ruler" value = "<%=Ruler%>"><%=Ruler%></td>
				</tr>
<%
If AuthOn = TRUE Then
%>
				<tr>
					<td align="left" width="50%"><p>Password:</p></td>
					<td align="left" width="50%"><p><input name="Password" size="10" maxlength="20" type="password" value = "<%=TheCookiePass%>"  class = "numbar"> &nbsp;<input type = "checkbox" name = "savepass" value = "yes" class = "button"> Save</p></td>
				</tr>
<%
End If
%>
				<tr>
					<td align="left" width="50%"><p>Subject:</p></td>
					<td align="left" width="50%"><input name="Subject" size="20" maxlength="30" type="text"  class = "numbar"></td>
				</tr>
			</table>
			<textarea name = "message" cols = "35" rows = "6" class = "numbar"></textarea><br>
			<input type = "hidden" name = "Original" value = "1">
			<input type = "submit" value = "Post" class = "button">
			</form>
			</p>
		</center>
		
		<p>
		
	</body>

</html>
<%
Set Conn = nothing
Set RS = nothing
%>