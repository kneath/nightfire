<%Option Explicit%>
<!--#Include File = "include.asp"-->
<%
Dim x, y, z, r, x5, x6, x7, x8, x9, x10, final
Function MAX(x, y, z, r, x5, x6, x7, x8, x9,x10)
	If x>y AND x>z AND x>r Then
		MAX = x
	ElseIf y>x AND y>z AND y>r AND y>x5 AND y>x6 AND y>x7 AND y>x8 AND y>x9 AND y>x10 Then
		MAX = y
	ElseIf z>x AND z>y AND z>r AND z>x5 AND z>x6 AND z>x7 AND z>x8 AND z>x9 AND z>x10 Then
		MAX = z
	ElseIf r>x AND r>y AND r>z AND r>x5 AND r>x6 AND r>x7 AND r>x8 AND r>x9 AND r>x10 Then
		MAX = r
	ElseIf x5>x AND x5>y AND x5>r AND x5>x6 AND x5>x7 AND x5>x8 AND x5>x9 AND x5>x10 Then
		MAX = x5
	ElseIf x6>x AND x6>y AND x6>r AND x6>x5 AND x6>x7 AND x6>x8 AND x6>x9 AND x6>x10 Then
		MAX = x6
	ElseIf x7>x AND x7>y AND x7>r AND x7>x6 AND x7>x5 AND x7>x8 AND x7>x9 AND x7>x10 Then
		MAX = x7
	ElseIf x8>x AND x8>y AND x8>r AND x8>x6 AND x8>x7 AND x8>x5 AND x8>x9 AND x8>x10 Then
		MAX = x8
	ElseIf x9>x AND x9>y AND x9>r AND x9>x6 AND x9>x7 AND x9>x8 AND x9>x5 AND x9>x10 Then
		MAX = x9
	ElseIf x10>x AND x10>y AND x10>r AND x10>x6 AND x10>x7 AND x10>x8 AND x10>x9 AND x10>x5 Then
		MAX = x10	
	Else
		MAX = "NoMax"
	End If
End Function

Dim UserName, Message
UserName = Request.Cookies("NightFire")("UserName")

Dim RSGI, RSGII, RSGIII, i
Set RSGI = Server.CreateObject("ADODB.Recordset")
Set RSGII = Server.CreateObject("ADODB.Recordset")
Set RSGIII = Server.CreateObject("ADODB.Recordset")
RSGI.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
RSGII.Open "SELECT * FROM GENERALINFO WHERE SystemNumber = "&RSGI("SystemNumber"), Conn, adOpenStatic, adLockPessimistic

If Request.Form("Nation") <> "" Then
	RSGI("KingChoice") = Request.Form("Nation")
	RSGI.Update
	Message = "<p class = 'normal'>Your vote has been changed</p>"
End If

Dim RSSS
Set RSSS = Server.CreateObject("ADODB.Recordset")
RSSS.Open "SELECT * FROM SYSTEMS WHERE SystemNumber="&RSGI("SystemNumber"), Conn, adOpenStatic, adLockPessimistic
If Request.Form("SystemName") <>  "" Then
	RSSS("SystemName") = Request.Form("SystemName")
	RSSS.Update
End If
Dim SystemName
SystemName = RSSS("SystemName")
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
<SCRIPT SRC="fieldcheck.js"></SCRIPT>


<!--#Include File = "ad.asp"-->
<!--#Include File = "statbar.asp"-->

<form action = "vote.asp" method = "POST">
<center>
<%=Message%>
<br><br>
<%
If RSGI("King") = "Yes" Then
%>
<p align=left>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;System Name: <input type="text" size="30" class="numbar" name="SystemName" maxlength=30 value = "<%=SystemName%>"> &nbsp;&nbsp;&nbsp;&nbsp;<input type=submit value = "Change" class = "button"><br><br><br><br><br><br></p>
<%
End If
%>
</form>
<form action = "vote.asp" method=post>
<table width = "90%">
<tr bgcolor = "#0000CC">
<td colspan=3 class = "large" align = "center">Voting Registry</td>
</tr>
<tr bgcolor = "#000066">
<td align = "center">Nation</td>
<td align = "center">Votes</td>
<td align = "center">Vote</td>
</tr>
<%
Dim arrNation(9), arrVotes(9), arrVote(9), arrKing(9)

For i = 0 to 9
	If RSGII.EOF = False Then
		arrNation(i) = RSGII("Nation")
		arrVote(i) = RSGII("KingChoice")
		RSGII.MoveNext
	Else
		arrNation(i) = "--"
		arrVote(i) = "-"
	End If
Next

For i = 0 to 9
	RSGII.MoveFirst
	RSGII.Filter = "KingChoice = '"&arrNation(i)&"'"
	If RSGII.EOF = TRUE Then
	arrVotes(i) = 0
	Else
	arrVotes(i) = RSGII.Recordcount
	End If
	RSGII.Filter = adFilterNone
Next

Dim King, x1
For i = 0 to 9
If arrVotes(i) = MAX(arrVotes(0), arrVotes(1), arrVotes(2), arrVotes(3), arrVotes(4), arrVotes(5), arrVotes(6), arrVotes(7), arrVotes(8), arrVotes(9)) Then
	King = arrNation(i)
	x1 = i
End If
Next

If King = "" Then
	King = "NoMax"
End If

If King <> "NoMax" Then
	RSGII.Filter = adFilterNone
	If RSGII.EOF = FALSE Then
	Do While RSGII.EOF = FALSE
		RSGII("King") = "No"
		RSGII.Update
		RSGII.MoveNext
	Loop
	RSGII.Movefirst
	'RSGII.Filter = "Nation = '"&King&"'"
	'RSGII.Update
	'RSGII("King") = "Yes"
	'RSGII.Update
	Conn.Execute("UPDATE GENERALINFO SET King = 'Yes' WHERE Nation = '"&King&"'")	
	End If
End If

arrKing(x1) = " <span class = 'alert'>*ESR*</span>"

For i = 0 to 9
	Response.Write "<tr>"
	Response.Write "<td bgcolor = '#000055'>"&arrNation(i) & arrKing(i)&"</td>"
	Response.Write "<td bgcolor = '#000044'>"&arrVotes(i)&"</td>"
	Response.Write "<td bgcolor = '#000055' class = 'normal'>"&arrVote(i)&"</td>"
	Response.Write "</tr>"
Next
%>
</table>
<br>
<select class = "numbar" name = "Nation">
<option value = "" selected>Choose a Nation to Vote For</option>
<%

Dim NationT
RSGIII.Open "SELECT * FROM GENERALINFO WHERE SystemNumber = " &RSGI("SystemNumber"), Conn
If RSGIII.EOF = False Then
	RSGIII.MoveFirst
	Do While Not RSGIII.EOF
		NationT = RSGIII("Nation")
		Response.Write "<option value = '"&NationT&"'>"&NationT&"</option>"
		RSGIII.MoveNext
	Loop
End If

RSGI.Close
RSGII.Close
RSGIII.Close
Set RSGI = nothing
Set RSGII = nothing
Set RSGIII = nothing

Conn.Close
Set Conn = nothing

%>
</select>
  &nbsp; &nbsp; &nbsp;<input type = "submit" class = "button" value = "Vote">
</form>
</body>
</html>