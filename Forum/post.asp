<%Option Explicit%>
<%Response.Buffer = TRUE%>
<!--#Include File = "include.asp"-->
<%

Dim TableName

Set TableName = Request.QueryString("table")

'Let's get all the info from the forms now
Dim Subject, Poster, Message, Time, Original, ThreadID, Posts

Set Poster = Request.Form("Poster")
Set Subject = Request.Form("Subject")
Set Message = Request.Form("Message")
ThreadID = Request.Form("ThreadID")
If Request.QueryString("admin") <> AdminPass Then
	Message = Server.HTMLEncode(Message)
	Poster = Server.HTMLEncode(Poster)
	Subject = Server.HTMLEncode(Subject)
	Subject = Replace(Subject, "'", "&#184;")
End If
Time = Now
Set Original = Request.Form("Original")
'Posts and ThreadID comes later
If Poster = "" OR Subject = "" OR Message = "" Then
%>
<html>
<head>
<title><%=MyTitle%>:Posting...</title>
<link rel="stylesheet" type="text/css" href="forum.css">
</head>
<body>
<center>
<p >I'm sorry, but you did not fill out one of the required fields<br></p>
<a href = "forum.asp?table=<%=TableName%>
<%
If Request.Querystring("admin") = AdminPass Then
%>
&admin=<%=Request.Querystring("admin")%>
<%
End If
%>
">Go Back to NightFire Forums</a>
</center>
</body>
</html>

<%
Else
Response.Cookies("BrakForum") = Poster
Response.Cookies("BrakForum").Expires = DateAdd("d", 30, Now())
Dim SQL, RS

'Just a reminder here: TableNames = Announcements, General, Strategy, Politics, BugsNErrors, Suggestions. Fields in the Tables: Poster, Subject, Message, Original, DatePosted, SubjectID, Posts.
If Request.Form("Original") = 1 Then
	SQL = "SELECT * FROM "&TableName
Else
	SQL = "SELECT * FROM "&TableName&" WHERE SubjectID = "&ThreadID
End If
'Response.Write SQL
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, Conn, 3, 3
'Now all connections have been made


If Original = 1 Then
	Posts = 1
Else
	Posts = RS.RecordCount
End If

If RS.EOF Then
	ThreadID = 1
Else
	Select Case Original

	Case "1"
		Dim RSTemp, SQLTemp

		SQLTemp = "SELECT SubjectID FROM "&TableName&" WHERE Original = 1 ORDER BY SubjectID DESC"
		Set RSTemp = Server.CreateObject("ADODB.Recordset")
		'Response.Write SQLTemp
		RSTemp.Open SQLTemp, Conn, 3, 3
		RSTemp.MoveFirst
		Set ThreadID = RSTemp("SubjectID")
		ThreadID = ThreadID + 1
		Set RSTemp = nothing
	Case "2"
		ThreadID = Request.Form("ThreadID")
		Set RSTemp = Server.CreateObject("ADODB.Recordset")
		RSTemp.Open "SELECT * FROM "&TableName&" WHERE SubjectID = "&ThreadID, Conn
		Subject = RSTemp("Subject")
		RSTemp.Close
		Set RSTemp = nothing
	End Select
End If


RS.AddNew

RS("Poster") = Poster
RS("Subject") = Subject
RS("Message") = Replace(Message, vbCrLf, "<br>")
RS("Original") = Original
RS("DatePosted") = Time
RS("SubjectID") = ThreadID
RS("Posts") = Posts
RS.Update
RS.Close

If Original = 2 Then
RS.Open "SELECT Posts, DatePosted FROM "&TableName&" WHERE SubjectID = "&ThreadID&" AND Original = 1", Conn
RS("Posts") = Posts + 1
RS("DatePosted") = Now()
RS.Update
End If
 'haha! all done w/ the ASP Code! There's no damn html! doh!
%>
<html>
<head>
<title>Posted Successfully!</title>
<link rel="stylesheet" type="text/css" href="forum.css">
</head>
<body>
<center>
<img src="banner.gif" align = "center">
<p>Your Message has been posted successfully!<br>
<a href = "forum.asp?table=<%=TableName%>
<%
If Request.Querystring("admin") = AdminPass Then
%>
&admin=<%=Request.Querystring("admin")%>
<%
End If
%>
">Go Back to <%=MyTitle%></a></p>
<%
End If
Set Conn = nothing
Set RS = nothing
%>