<%@ Language=VBScript %>
<% Option Explicit %>

<!--#Include File = "include.asp"-->
<%
Response.Buffer =  TRUE

Dim UserName

UserName = Request.Cookies("NightFire")("UserName")

Dim SystemNumberRS, SystemNumber, PlayerSystem

Set SystemNumberRS = Conn.Execute("SELECT SystemNumber, UserName FROM GENERALINFO WHERE UserName = '"&UserName&"'")

PlayerSystem = SystemNumberRS("SystemNumber")

If Request.Form("System") <> "" Then
	Response.Cookies("NightFire")("SelectedSystem") = Request.Form("System")
End If
If Request.Cookies("NightFire")("SelectedSystem") = ""  Then
	Response.Cookies("NightFire")("SelectedSystem") = PlayerSystem
End If
	'they have some other system in there that needs to stay.
SystemNumber = Request.Cookies("NightFire")("SelectedSystem")



Set SystemNumberRS = nothing

Dim RS, RST

Set RS = Server.CreateObject("ADODB.Recordset")
Set RST = Server.CreateObject("ADODB.Recordset")
RS.Open "SELECT a.*, b.* FROM GENERALINFO a, GENERALSTATS b WHERE a.UserName = b.UserName AND a.SystemNumber = "&SystemNumber&" ORDER BY Land DESC", Conn, adOpenForwardOnly, adLockPessimistic
RST.Open "SELECT * FROM SYSTEMS WHERE SystemNumber="&SystemNumber, Conn
 Dim Nation, Land, Networth, StarType
StarType = RST("SystemType")
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

<body bgcolor="#000000" text="#ffffff" link="#ffffff" vlink="#ffffff" alink="ffcc33" >

<!--#Include File = "ad.asp"-->
<!--#Include File = "statbar.asp"-->
<br>

<table cellspacing=2 cellpadding=2 colspan=140 align = "center">
	<tr>
		<td colspan=140 bgcolor="#0000CC">
			<div class = "large" align = "center">System</div>
		</td>
	</tr>
<tr>	<td colspan=140 bgcolor="#000066" align = "center"><form action = "system.asp" method = "post">Change System: &nbsp;&nbsp; <input name = "System" value = "<%=SystemNumber%>" class = "numbar" size = "2"> &nbsp;&nbsp; <input type = "submit" class = "button" value = "Change System"></form></td</tr>
	<tr>	<td colspan=140 bgcolor="#000044" align = "center">
			<div class = "normal" align = "center"><%=RST("SystemName")%> (#<%=SystemNumber%>) - <%=StarType%></div>
		</td>
	</tr>
	<tr bgcolor="#000066">
		<td width="45%"><div align = "center">Nation</div></td>
		<td width="15%"><div align = "center">Race</div></td>
		<td width="20%"><div align = "center">Land</div></td>
		<td width="20%"><div align = "center">Networth</div></td>
	</tr>
<%
Dim y, Race, UserNamee
y = 1

If RS.EOF = TRUE Then
	Response.Write "	<tr>"
	Response.Write "		<td colspan=140 bgcolor = '#000044'><div align = 'center'>No Nations Found</div></td>"
	Response.Write "	</tr>"
Else
	Do While RS.EOF = FALSE
		Nation = RS("Nation")
		If RS("King") = "Yes" Then
			Nation = Nation & " <span class='alert'>*ESR*</span>"
		End If
		Land = RS("Land")
		Networth = RS("Networth")
		Race = RS("Race")
		UserNamee = RS("UserName")
		If y = 1 Then
			Response.Write "	<tr bgcolor = '#000033'>"
			y = 2
		Else
			Response.Write "	<tr bgcolor = '#000044'>"
			y = 1
		End If
		If UserNamee = UserName Then
			Response.Write "		<td width='45%'><div align = 'left' class = 'normal'>"&Nation&"</div></td>"
			Response.Write "		<td width='15%'><div align = 'center' class = 'normal'>"&Race&"</div></td>"
			Response.Write "		<td width='20%'><div align = 'center' class = 'normal'>"&FormatNumber(Land, 0)&"</div></td>"
			Response.Write "		<td width='20%'><div align = 'center' class = 'normal'>"&FormatNumber(Networth, 0)&"</div></td>"
			Response.Write "	</tr>"
		Else
			Response.Write "		<td width='45%'><div align = 'left'>"&Nation&"</div></td>"
			Response.Write "		<td width='15%'><div align = 'center'>"&Race&"</div></td>"
			Response.Write "		<td width='20%'><div align = 'center'>"&FormatNumber(Land, 0)&"</div></td>"
			Response.Write "		<td width='20%'><div align = 'center'>"&FormatNumber(Networth, 0)&"</div></td>"
			Response.Write "	</tr>"
		End If
		RS.MoveNext
	Loop
End If
RS.Close
Conn.Close
%>


</table>
</body></html>
