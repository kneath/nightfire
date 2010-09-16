<%@ Language=VBScript %>
<% Option Explicit %>
<%Response.Buffer=TRUE%>
<!--#Include File = "include.asp"-->
<%
Dim UserName
UserName = Request.Cookies("NightFire")("UserName")

Call OpenConnection()

Dim RSTemp, RSM
Set RSTemp = Server.CreateObject("ADODB.Recordset")
Set RSM = Server.CreateObject("ADODB.Recordset")
RSTemp.Open "SELECT UserName, SystemNumber FROM GENERALINFO WHERE UserName = '" & UserName & "'", Conn, adOpenForwardOnly

Dim SystemNumber
SystemNumber = RSTemp("SystemNumber")
RSTemp.Close

Dim Date1, Date2
RSM.Open "SELECT * FROM MESSAGES WHERE Type = 'SystemNews"& SystemNumber & "' ORDER BY Time DESC", Conn, adOpenStatic
If RSM.EOF = FALSE Then
	Date2 = FormatDateTime(RSM("Time"), vbShortDate)
	RSM.MoveLast
	Date1 = FormatDateTime(RSM("Time"), vbShortDate)
	RSM.MoveFirst
End If
%>
<html>
	<head>
		<title>NightFire</title>
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
		<!--#Include File = "ad.asp"-->
		<!--#Include File = "statbar.asp"-->
		<br>
		<br>
		<p class = "large" align = "center">System News</p>
		<table width = "85%" align = "center">
			<tr bgcolor = "#0000CC">
				<td colspan=2 align=center>News For <%=Date1%> - <%=Date2%></td>
			</tr>
<%
If RSM.EOF = FALSE Then
RSM.MoveFirst
If RSM.EOF = FALSE Then
	Dim strTime, Message
	Do While Not RSM.EOF
		strTime = RSM("Time")
		Message = RSM("Message")
%>
			<tr>
				<td bgcolor = "#000066"><%=strTime%></td>
				<td bgcolor = "#000033"><%=Message%></td>
			</tr>
<%
		RSM.MoveNext
	Loop
End If
End If
RSM.Close
Conn.Close
%>
	</body>
</html>
