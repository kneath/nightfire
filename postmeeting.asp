<%Option Explicit%>
<%Response.Buffer = TRUE%>
<!--#Include File = "include.asp"-->
<%

Dim TableName

TableName = Request.QueryString("table")

'Let's get all the info from the forms now
Dim Subject, Poster, Message, Time, Original, ThreadID, Posts

Subject = Request.Form("Subject")
Message = Request.Form("Message")
If Request.QueryString("admin") <> AdminPass Then
	Message = Server.HTMLEncode(Message)
	Poster = Server.HTMLEncode(Poster)
	Subject = Server.HTMLEncode(Subject)
End If
Time = Now
Set Original = Request.Form("Original")
'Posts and ThreadID comes later
If Subject = "" OR Message = "" Then
%>
<html>
<head>
<title><%=MyTitle%>:Posting...</title>
<link rel="stylesheet" type="text/css" href="game.css">
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
<p >I'm sorry, but you did not fill out one of the required fields<br></p>
<a href = "meeting.asp?table=<%=TableName%>
<%
If Request.Querystring("admin") = AdminPass Then
%>
&admin=<%=Request.Querystring("admin")%>
<%
End If
%>
">Go Back to System Meeting</a>
</center>
</body>
</html>
<%
Else
Dim SQL, RS

'Just a reminder here: TableNames = Announcements, General, Strategy, Politics, BugsNErrors, Suggestions. Fields in the Tables: Poster, Subject, Message, Original, DatePosted, SubjectID, Posts.
If Request.Form("Original") = 1 Then
	SQL = "SELECT * FROM "&TableName
Else
	SQL = "SELECT * FROM "&TableName&" WHERE SubjectID = "&Request.Form("ThreadID")
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
	End Select
End If


Dim UserName
UserName = Request.Cookies("NightFire")("UserName")

Dim RSGI
Set RSGI = Server.CreateObject("ADODB.Recordset")

RSGI.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserName&"'", Conn, adOpenForwardOnly, adLockPessimistic

Dim System, Minister, Minister2
Minister = RSGI("Minister")
Minister2 = RSGI("Nation")

System = RSGI("SystemNumber")

Poster = Minister + " of " + Minister2

RS.AddNew

RS("Poster") = Poster
RS("Subject") = Subject
RS("Message") = Replace(Message, vbCrLf, "<br>")
RS("Original") = Original
RS("DatePosted") = Time
RS("SubjectID") = ThreadID
RS("Posts") = Posts
RS("System") = System
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
<link rel="stylesheet" type="text/css" href="game.css">
</head>
<body bgcolor="#000000" text="#ffffff" link="#ff6633" vlink="#ff6633" alink="ffcc33">
<center>
<img src="banner.gif" align = "center">
<p>Your Message has been posted successfully!<br>
<a href = "meeting.asp?table=<%=TableName%>
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