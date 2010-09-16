<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="include.asp"-->
<% Response.Buffer = True %>
<%

Dim UserName
UserName = Request.Cookies("NightFire")("UserName")

Call OpenConnection()

Dim Capital, ElementX, Energy, Mana, Plutonium, Antimatter, Land

Dim RST
Set RST = Conn.Execute("SELECT Civillian, Land, UserName FROM GENERALSTATS WHERE UserName = '"&UserName&"'")
Capital = RST("Civillian")
Land = RST("Land")
RST.Close
Set RST = Conn.Execute("SELECT Consulates, ElementXMines, AtomicPlants, AntimatterPlants, SolarPanels, Temples, PlutoniumPlants, AntimatterLabs, UserName FROM DONEBUILDINGS WHERE UserName = '"&UserName&"'")
Capital = "+" & Int(Capital*(RST("Consulates")/Land) + Capital)
ElementX = "+" & RST("ElementXMines")*50
Energy = "N/A"
Mana = "+" & RST("Temples")*10
Plutonium = "+" & RST("PLutoniumPlants")*10
Antimatter = "+" & RST("AntimatterLabs")*3
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

	<body bgcolor="#000000" text="#ffffff" link="#ffffff" vlink="#ffffff" alink="ffcc33" MARGINWIDTH=0 MARGINHEIGHT=0 TOPMARGIN=0 BOTTOMMARGIN=0 LEFTMARGIN=0 RIGHTMARGIN=0>
		<center>
			<!--#Include File = "ad.asp"-->
			<!--#Include File = "statbar.asp"-->
			<table cellspacing=0 cellpadding=5 width="60%">
				<tr>
					<td bgcolor="#0000CC" colspan=20><div align=center class="large">Economy Profile</div></td>
				</tr>
				<tr bgcolor = "#000044">
					<td>Resource</td>
					<td>Net Change</td>
				</tr>
				<tr bgcolor = "#000066">
					<td>Capital[c]</td>
					<td class = "normal">+<%=FormatNumber(Capital, 0)%></td>
				</tr>
				<tr bgcolor = "#000044">
					<td>Element-X[t]</td>
					<td class = "normal">+<%=FormatNumber(ElementX, 0)%></td>
				</tr>
				<tr bgcolor = "#000066">
					<td>Energy[j]</td>
					<td class = "normal"><%=Energy%></td>
				</tr>
				<tr bgcolor = "#000044">
					<td>Mana[m]</td>
					<td class = "normal">+<%=FormatNumber(Mana, 0)%></td>
				</tr>
				<tr bgcolor = "#000066">
					<td>Plutonium[p]</td>
					<td class = "normal">+<%=FormatNumber(Plutonium, 0)%></td>
				</tr>
				<tr bgcolor = "#000044">
					<td>Antimatter[a]</td>
					<td class = "normal">+<%=FormatNumber(Antimatter, 0)%></td>
				</tr>
			</table>
		</center
	</body>
</html>
