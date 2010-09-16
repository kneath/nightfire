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
			<p>
				<img src = "banner.gif" align = "center">
			<table border="0" cellpadding="3" cellspacing="0" width="100%">
				<caption bgcolor="000080">
					<p class="large"><%=MyTitle%><br></p>
				
<%
Dim TheCookieName, TheCookiePass

TheCookieName = Request.Cookies("BrakForum")
If AuthOn = TRUE Then
	TheCookiePass = Request.Cookies("BrakForumPass")
End If


Dim TableName

Set TableName = Request.QueryString("table")

If TableName = "" Then
TableName = "gen"
End If
Response.Write "<form action = 'delete.asp?table="&TableName& "&admin="&Request.Querystring("admin")&"' method = 'post'>"

Dim SQL, RS

'Just a reminder here: TableNames = Announcements, Gen, Strategy, Politics, BugsNErrors, Suggestions. Fields in the Tables: Poster, Subject, Message, Original, DatePosted, SubjectID, Posts.
SQL = "SELECT * FROM "&TableName&" WHERE Original = 1 ORDER BY DatePosted DESC"
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

Response.Write "</caption><tr class = 'titlebar'>"
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
	Dim Poster, Subject, Posts, TimePosted, ColorNum, Color, ThreadID
	
	Color = 1
	ColorNum = 0

	Do While Not RS.EOF
		Set Poster = RS("Poster")
		Set Subject = RS("Subject")
		Set Posts = RS("Posts")
		Set TimePosted = RS("DatePosted")
		Set ThreadID = RS("SubjectID")
		
		If ColorNum/2 = ColorNum\2 Then
			Color = "000066"
		Else
			Color = "000033"
		End If
		Response.Write "<tr>"
		If admin = "yes" Then
		Response.Write "<td bgcolor='#000044'><div><input type = 'checkbox' value = '"&ThreadID&"' name = 'Delete" & ColorNum & "' class = 'button'></td>"
		End If
		Response.Write "<td bgcolor='#000066'><div><a href = 'thread.asp?table="&TableName&"&id="&ThreadID&"&Posts="&Posts
		If Request.QueryString("admin") = AdminPass Then
			Response.Write "&admin="&Request.Querystring("admin")
		End If
		Response.Write "'>" & VbCrlf &Subject&"</a></div></td>" & VbCrlf
		Response.Write "<td bgcolor='#000044'><div>"&Posts&"</div></td>" & VbCrlf
		Response.Write "<td bgcolor='#000066'><div>"&Poster&"</div></td>" & VbCrlf
		Response.Write "<td bgcolor='#000044'><div>"&TimePosted&"</div></td></tr>" & VbCrlf
		
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
			<form action="post.asp?table=<%=TableName%>
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
					<td align="left" width="50%"><p>Poster:</p></td>
					<td align="left" width="50%"><input name="Poster" size="10" maxlength="13" type="text" value = "<%=TheCookieName%>"  class = "numbar"></td>
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