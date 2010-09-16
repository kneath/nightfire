<!--#Include File = "include.asp"-->
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
<SCRIPT SRC="fieldcheck.js"></SCRIPT>
<p align = "center" class = "large">Top 100 Nations By Land</p>
<br><br><br>
<table width = "75%" align = "center">
	<tr bgcolor = "#0000CC">
		<td>Rank</td>
		<td>Nation</td>
		<td>System</td>
		<td>Land</td>
	</tr>
<%
Dim Nation, Land, NW, System

Dim RS 
Set RS = Server.CreateObject("ADODB.Recordset")

RS.Open "SELECT a.*, b.* FROM GENERALINFO a, GENERALSTATS b WHERE a.UserName = b.UserName ORDER BY Land DESC", Conn

Dim i
For i = 1 to 100

	If RS.EOF = TRUE Then
		Nation = "-"
		Land = "-"
		NW = "-"
		System = "-"
	Else
		Nation = RS("Nation")
		Land = FormatNumber(RS("Land"), 0)
		NW = RS("Networth")
		System = RS("SystemNumber")
		RS.MoveNext
	End If
%>
	<tr>
		<td bgcolor = "#000033"><%=i%></td>
		<td bgcolor = "#000066"><%=Nation%></td>
		<td bgcolor = "#000033"><%=System%></td>
		<td bgcolor = "#000066"><%=Land%></td>
	</tr>
<%

Next

%>
</table>
	

