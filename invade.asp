<!--#Include File = "invade_code.asp"-->
<html>
<head>
<title> NightFire-Invade </title>
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
<!--#Include File = "ad.asp"-->
<SCRIPT SRC="fieldcheck.js"></SCRIPT>
<center>
<!--#include file = "statbar.asp"-->


<%
Response.Write SuccessString
Response.Write ErrorMessage
%>
<form action = "invade.asp" method = "post">
<table cellspacing=0 cellpadding=5 width=630>
<tr><td bgcolor="#0000CC" colspan=3><div align=center class="large">Invade</div></td></tr>
<tr><td bgcolor="#000066" colspan=3><div align=center><b>Unit Stats: <input type=text size=50 value="put mouse over unit to see stats" class="statbar"></b></div></td></tr>
<tr bgcolor="#000044">
<td align = "center" class = "Large" colspan=3> Availible Military to send</td>
</tr>
<tr bgcolor = "#00033">
<td align = "left" class = "small" size = "30%" bgcolor = "#00033"><a href = "#" onMouseover = "statbarland('1/1 | 1j')">Infantry</a></td><td align = "left" size = "40%" bgcolor = "#00066" class = "small"><%=AvailibleInfantry%></td><td size = "30%" align ="left" bgcolor = "#00033"><input name = "Infantry" size = "10" value = "0" class = "numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td></tr>
<td align = "left" class = "small" size = "30%" bgcolor = "#00033"><a href = "#" onMouseover = "statbarland('<%=Elite1Offense%>/<%=Elite1Defense%> | <%=Elite1Energy%>j | <%=Elite1Desc%>')"><%=Elite1Name%></a></td><td align = "left" size = "40%" bgcolor = "#00066" class = "small"><%=AvailibleElite1%></td><td size = "30%" align ="left" bgcolor = "#00033"><input name = "Elite1" size = "10" value = "0" class = "numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td></tr>
<td align = "left" class = "small" size = "30%" bgcolor = "#00033"><a href = "#" onMouseover = "statbarland('<%=Elite2Offense%>/<%=Elite2Defense%> | <%=Elite2Energy%>j | <%=Elite2Desc%>')"><%=Elite2Name%></a></td><td align = "left" size = "40%" bgcolor = "#00066" class = "small"><%=AvailibleElite2%></td><td size = "30%" align ="left" bgcolor = "#00033"><input name = "Elite2" size = "10" value = "0" class = "numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td></tr>
<td align = "left" class = "small" size = "30%" bgcolor = "#00033"><a href = "#" onMouseover = "statbarland('<%=Elite3Offense%>/<%=Elite3Defense%> | <%=Elite3Energy%>j | <%=Elite3Desc%>')" ><%=Elite3Name%></a></td><td align = "left" size = "40%" bgcolor = "#00066" class = "small"><%=AvailibleElite3%></td><td size = "30%" align ="left" bgcolor = "#00033"><input name = "Elite3" size = "10" value = "0" class = "numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td></tr>
<tr bgcolor="#000066">
<td colspan=3>
<p>Total Offensive Power: <input type="text" size=10 value="0" class="numbar" name = "Offense"><br>
Required Energy: <input type="text" size=10 value="0" class="numbar" name = "Energy"><br>
Available Transports: <span class = 'normal'><%=AvailibleTransports%></span> <br>
Needed Transports: <input type="text" size=5 value="0" class="numbar" name = "NeedTransports">
<p>Target System: <input size = "3" name = "System" value = "<%=SelectedSystem%>" class = "numbar"> <input type="submit" value="Change System" class="button">
<select name = "Nation" class = "numbar">
<option selected value = "0">Choose a Nation to Attack
<%
Do While Not objRS2invade.EOF
Response.Write "<option value = '"&objRS2invade("UserName")&"'>"&objRS2invade("Nation")
objRS2invade.MoveNext
Loop
%>
</select>
<input type= "hidden" name = "attack" value = "no"><input type = "submit" value = "Invade" class = "button" onClick = "document.forms[0].elements['attack'].value = 'yes';"></form>
</td></tr>
</table>

<%
objRSinvade.Close
objRS2invade.Close
RSUser.Close
RSUserMilitary.Close
RSUserAttackingMilitary.Close
Conn.Close

%>