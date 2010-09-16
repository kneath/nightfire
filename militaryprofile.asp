<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="include.inc"-->
<% Response.Buffer = True %>
<%
Dim UserNameM
UserNameM = Request.Cookies("NightFire")("UserName")

'Open Connection
Call OpenConnection()
%>
<!--#Include File = "militaryprofile_code.asp"-->
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
			<table cellspacing=0 cellpadding=5 width="100%">
				<tr>
					<td bgcolor="#0000CC" colspan=20><div align=center class="large">Military Profile</div></td>
				</tr>
				<tr bgcolor="#000044">
					<td><div><b>Incoming &nbsp;&nbsp;&nbsp;</b></div></td>
					<td align="center" class="small">1</td>
					<td align="center" class="small">2</td>
					<td align="center" class="small">3</td>
					<td align="center" class="small">4</td>
					<td align="center" class="small">5</td>
					<td align="center" class="small">6</td>
					<td align="center" class="small">7</td>
					<td align="center" class="small">8</td>
					<td align="center" class="small">9</td>
					<td align="center" class="small">10</td>
					<td align="center" class="small">11</td>
					<td align="center" class="small">12</td>
					<td align="center" class="small">13</td>
					<td align="center" class="small">14</td>
					<td align="center" class="small">15</td>
					<td align="center" class="small">16</td>
					<td align="center" class="small">17</td>
					<td align="center" class="small">18</td>
					<td align="center" class="small">Total</td>
				</tr>

				<tr bgcolor="#000066">
					<td><div>Infantry</div></td>
					<td align="center" class="small"><%=arrInfantry(0)%></td>
					<td align="center" class="small"><%=arrInfantry(1)%></td>
					<td align="center" class="small"><%=arrInfantry(2)%></td>
					<td align="center" class="small"><%=arrInfantry(3)%></td>
					<td align="center" class="small"><%=arrInfantry(4)%></td>
					<td align="center" class="small"><%=arrInfantry(5)%></td>
					<td align="center" class="small"><%=arrInfantry(6)%></td>
					<td align="center" class="small"><%=arrInfantry(7)%></td>
					<td align="center" class="small"><%=arrInfantry(8)%></td>
					<td align="center" class="small"><%=arrInfantry(9)%></td>
					<td align="center" class="small"><%=arrInfantry(10)%></td>
					<td align="center" class="small"><%=arrInfantry(11)%></td>
					<td align="center" class="small"><%=arrInfantry(12)%></td>
					<td align="center" class="small"><%=arrInfantry(13)%></td>
					<td align="center" class="small"><%=arrInfantry(14)%></td>
					<td align="center" class="small"><%=arrInfantry(15)%></td>
					<td align="center" class="small"><%=arrInfantry(16)%></td>
					<td align="center" class="small"><%=arrInfantry(17)%></td>
					<td align="center" class="small">(<%=TotInfantry%>) <%=DInfantry%></td>
				</tr>

				<tr bgcolor="#000044">
					<td><div><%=Elite1NM%></div></td>
					<td align="center" class="small"><%=arrElite1(0)%></td>
					<td align="center" class="small"><%=arrElite1(1)%></td>
					<td align="center" class="small"><%=arrElite1(2)%></td>
					<td align="center" class="small"><%=arrElite1(3)%></td>
					<td align="center" class="small"><%=arrElite1(4)%></td>
					<td align="center" class="small"><%=arrElite1(5)%></td>
					<td align="center" class="small"><%=arrElite1(6)%></td>
					<td align="center" class="small"><%=arrElite1(7)%></td>
					<td align="center" class="small"><%=arrElite1(8)%></td>
					<td align="center" class="small"><%=arrElite1(9)%></td>
					<td align="center" class="small"><%=arrElite1(10)%></td>
					<td align="center" class="small"><%=arrElite1(11)%></td>
					<td align="center" class="small"><%=arrElite1(12)%></td>
					<td align="center" class="small"><%=arrElite1(13)%></td>
					<td align="center" class="small"><%=arrElite1(14)%></td>
					<td align="center" class="small"><%=arrElite1(15)%></td>
					<td align="center" class="small"><%=arrElite1(16)%></td>
					<td align="center" class="small"><%=arrElite1(17)%></td>
					<td align="center" class="small">(<%=TotElite1%>) <%=DElite1%></td>
				</tr>

				<tr bgcolor="#000066">
					<td><div><%=Elite2NM%></div></td>
					<td align="center" class="small"><%=arrElite2(0)%></td>
					<td align="center" class="small"><%=arrElite2(1)%></td>
					<td align="center" class="small"><%=arrElite2(2)%></td>
					<td align="center" class="small"><%=arrElite2(3)%></td>
					<td align="center" class="small"><%=arrElite2(4)%></td>
					<td align="center" class="small"><%=arrElite2(5)%></td>
					<td align="center" class="small"><%=arrElite2(6)%></td>
					<td align="center" class="small"><%=arrElite2(7)%></td>
					<td align="center" class="small"><%=arrElite2(8)%></td>
					<td align="center" class="small"><%=arrElite2(9)%></td>
					<td align="center" class="small"><%=arrElite2(10)%></td>
					<td align="center" class="small"><%=arrElite2(11)%></td>
					<td align="center" class="small"><%=arrElite2(12)%></td>
					<td align="center" class="small"><%=arrElite2(13)%></td>
					<td align="center" class="small"><%=arrElite2(14)%></td>
					<td align="center" class="small"><%=arrElite2(15)%></td>
					<td align="center" class="small"><%=arrElite2(16)%></td>
					<td align="center" class="small"><%=arrElite2(17)%></td>
					<td align="center" class="small">(<%=TotElite2%>) <%=DElite2%></td>
				</tr>

				<tr bgcolor="#000044">
					<td><div><%=Elite3NM%></div></td>
					<td align="center" class="small"><%=arrElite3(0)%></td>
					<td align="center" class="small"><%=arrElite3(1)%></td>
					<td align="center" class="small"><%=arrElite3(2)%></td>
					<td align="center" class="small"><%=arrElite3(3)%></td>
					<td align="center" class="small"><%=arrElite3(4)%></td>
					<td align="center" class="small"><%=arrElite3(5)%></td>
					<td align="center" class="small"><%=arrElite3(6)%></td>
					<td align="center" class="small"><%=arrElite3(7)%></td>
					<td align="center" class="small"><%=arrElite3(8)%></td>
					<td align="center" class="small"><%=arrElite3(9)%></td>
					<td align="center" class="small"><%=arrElite3(10)%></td>
					<td align="center" class="small"><%=arrElite3(11)%></td>
					<td align="center" class="small"><%=arrElite3(12)%></td>
					<td align="center" class="small"><%=arrElite3(13)%></td>
					<td align="center" class="small"><%=arrElite3(14)%></td>
					<td align="center" class="small"><%=arrElite3(15)%></td>
					<td align="center" class="small"><%=arrElite3(16)%></td>
					<td align="center" class="small"><%=arrElite3(17)%></td>
					<td align="center" class="small">(<%=TotElite3%>) <%=DElite3%></td>
				</tr>

				<tr bgcolor="#000066">
					<td><div>Transport</div></td>
					<td align="center" class="small"><%=arrTransports(0)%></td>
					<td align="center" class="small"><%=arrTransports(1)%></td>
					<td align="center" class="small"><%=arrTransports(2)%></td>
					<td align="center" class="small"><%=arrTransports(3)%></td>
					<td align="center" class="small"><%=arrTransports(4)%></td>
					<td align="center" class="small"><%=arrTransports(5)%></td>
					<td align="center" class="small"><%=arrTransports(6)%></td>
					<td align="center" class="small"><%=arrTransports(7)%></td>
					<td align="center" class="small"><%=arrTransports(8)%></td>
					<td align="center" class="small"><%=arrTransports(9)%></td>
					<td align="center" class="small"><%=arrTransports(10)%></td>
					<td align="center" class="small"><%=arrTransports(11)%></td>
					<td align="center" class="small"><%=arrTransports(12)%></td>
					<td align="center" class="small"><%=arrTransports(13)%></td>
					<td align="center" class="small"><%=arrTransports(14)%></td>
					<td align="center" class="small"><%=arrTransports(15)%></td>
					<td align="center" class="small"><%=arrTransports(16)%></td>
					<td align="center" class="small"><%=arrTransports(17)%></td>
					<td align="center" class="small">(<%=TotTransports%>) <%=DTransports%></td>
				</tr>
				</span>
<!--#Include File = "militaryprofile_code2.asp"-->
			</table>
			<table cellspacing=0 cellpadding=5 width="100%">
				<tr bgcolor="#000000">
					<td><div align="center">&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
				</tr>
				<tr bgcolor="#000044">
					<td><div><b>Returning</b></div></td>
					<span class = "small">
					<td align="center" class="small">1</td>
					<td align="center" class="small">2</td>
					<td align="center" class="small">3</td>
					<td align="center" class="small">4</td>
					<td align="center" class="small">5</td>
					<td align="center" class="small">6</td>
					<td align="center" class="small">7</td>
					<td align="center" class="small">8</td>
					<td align="center" class="small">9</td>
					<td align="center" class="small">10</td>
					<td align="center" class="small">11</td>
					<td align="center" class="small">12</td>
					<td align="center" class="small">13</td>
					<td align="center" class="small">14</td>
					<td align="center" class="small">15</td>
					<td align="center" class="small">16</td>
					<td align="center" class="small">17</td>
					<td align="center" class="small">18</td>
					<td align="center" class="small">Total</td>
				</tr>

				<tr bgcolor="#000066">
					<td><div>Infantry</div></td>
					<td align="center" class="small"><%=arrInfantry(0)%></td>
					<td align="center" class="small"><%=arrInfantry(1)%></td>
					<td align="center" class="small"><%=arrInfantry(2)%></td>
					<td align="center" class="small"><%=arrInfantry(3)%></td>
					<td align="center" class="small"><%=arrInfantry(4)%></td>
					<td align="center" class="small"><%=arrInfantry(5)%></td>
					<td align="center" class="small"><%=arrInfantry(6)%></td>
					<td align="center" class="small"><%=arrInfantry(7)%></td>
					<td align="center" class="small"><%=arrInfantry(8)%></td>
					<td align="center" class="small"><%=arrInfantry(9)%></td>
					<td align="center" class="small"><%=arrInfantry(10)%></td>
					<td align="center" class="small"><%=arrInfantry(11)%></td>
					<td align="center" class="small"><%=arrInfantry(12)%></td>
					<td align="center" class="small"><%=arrInfantry(13)%></td>
					<td align="center" class="small"><%=arrInfantry(14)%></td>
					<td align="center" class="small"><%=arrInfantry(15)%></td>
					<td align="center" class="small"><%=arrInfantry(16)%></td>
					<td align="center" class="small"><%=arrInfantry(17)%></td>
					<td align="center" class="small"><%=TotInfantry%></td>
				</tr>

				<tr bgcolor="#000044">
					<td><div><%=Elite1NM%></div></td>
					<td align="center" class="small"><%=arrElite1(0)%></td>
					<td align="center" class="small"><%=arrElite1(1)%></td>
					<td align="center" class="small"><%=arrElite1(2)%></td>
					<td align="center" class="small"><%=arrElite1(3)%></td>
					<td align="center" class="small"><%=arrElite1(4)%></td>
					<td align="center" class="small"><%=arrElite1(5)%></td>
					<td align="center" class="small"><%=arrElite1(6)%></td>
					<td align="center" class="small"><%=arrElite1(7)%></td>
					<td align="center" class="small"><%=arrElite1(8)%></td>
					<td align="center" class="small"><%=arrElite1(9)%></td>
					<td align="center" class="small"><%=arrElite1(10)%></td>
					<td align="center" class="small"><%=arrElite1(11)%></td>
					<td align="center" class="small"><%=arrElite1(12)%></td>
					<td align="center" class="small"><%=arrElite1(13)%></td>
					<td align="center" class="small"><%=arrElite1(14)%></td>
					<td align="center" class="small"><%=arrElite1(15)%></td>
					<td align="center" class="small"><%=arrElite1(16)%></td>
					<td align="center" class="small"><%=arrElite1(17)%></td>
					<td align="center" class="small"><%=TotElite1%></td>
				</tr>

				<tr bgcolor="#000066">
					<td><div><%=Elite2NM%></div></td>
					<td align="center" class="small"><%=arrElite2(0)%></td>
					<td align="center" class="small"><%=arrElite2(1)%></td>
					<td align="center" class="small"><%=arrElite2(2)%></td>
					<td align="center" class="small"><%=arrElite2(3)%></td>
					<td align="center" class="small"><%=arrElite2(4)%></td>
					<td align="center" class="small"><%=arrElite2(5)%></td>
					<td align="center" class="small"><%=arrElite2(6)%></td>
					<td align="center" class="small"><%=arrElite2(7)%></td>
					<td align="center" class="small"><%=arrElite2(8)%></td>
					<td align="center" class="small"><%=arrElite2(9)%></td>
					<td align="center" class="small"><%=arrElite2(10)%></td>
					<td align="center" class="small"><%=arrElite2(11)%></td>
					<td align="center" class="small"><%=arrElite2(12)%></td>
					<td align="center" class="small"><%=arrElite2(13)%></td>
					<td align="center" class="small"><%=arrElite2(14)%></td>
					<td align="center" class="small"><%=arrElite2(15)%></td>
					<td align="center" class="small"><%=arrElite2(16)%></td>
					<td align="center" class="small"><%=arrElite2(17)%></td>
					<td align="center" class="small"><%=TotElite2%></td>
				</tr>

				<tr bgcolor="#000044">
					<td><div><%=Elite3NM%></div></td>
					<td align="center" class="small"><%=arrElite3(0)%></td>
					<td align="center" class="small"><%=arrElite3(1)%></td>
					<td align="center" class="small"><%=arrElite3(2)%></td>
					<td align="center" class="small"><%=arrElite3(3)%></td>
					<td align="center" class="small"><%=arrElite3(4)%></td>
					<td align="center" class="small"><%=arrElite3(5)%></td>
					<td align="center" class="small"><%=arrElite3(6)%></td>
					<td align="center" class="small"><%=arrElite3(7)%></td>
					<td align="center" class="small"><%=arrElite3(8)%></td>
					<td align="center" class="small"><%=arrElite3(9)%></td>
					<td align="center" class="small"><%=arrElite3(10)%></td>
					<td align="center" class="small"><%=arrElite3(11)%></td>
					<td align="center" class="small"><%=arrElite3(12)%></td>
					<td align="center" class="small"><%=arrElite3(13)%></td>
					<td align="center" class="small"><%=arrElite3(14)%></td>
					<td align="center" class="small"><%=arrElite3(15)%></td>
					<td align="center" class="small"><%=arrElite3(16)%></td>
					<td align="center" class="small"><%=arrElite3(17)%></td>
					<td align="center" class="small"><%=TotElite3%></td>
				</tr>

				<tr bgcolor="#000066">
					<td><div>Transport</div></td>
					<td align="center" class="small"><%=arrTransports(0)%></td>
					<td align="center" class="small"><%=arrTransports(1)%></td>
					<td align="center" class="small"><%=arrTransports(2)%></td>
					<td align="center" class="small"><%=arrTransports(3)%></td>
					<td align="center" class="small"><%=arrTransports(4)%></td>
					<td align="center" class="small"><%=arrTransports(5)%></td>
					<td align="center" class="small"><%=arrTransports(6)%></td>
					<td align="center" class="small"><%=arrTransports(7)%></td>
					<td align="center" class="small"><%=arrTransports(8)%></td>
					<td align="center" class="small"><%=arrTransports(9)%></td>
					<td align="center" class="small"><%=arrTransports(10)%></td>
					<td align="center" class="small"><%=arrTransports(11)%></td>
					<td align="center" class="small"><%=arrTransports(12)%></td>
					<td align="center" class="small"><%=arrTransports(13)%></td>
					<td align="center" class="small"><%=arrTransports(14)%></td>
					<td align="center" class="small"><%=arrTransports(15)%></td>
					<td align="center" class="small"><%=arrTransports(16)%></td>
					<td align="center" class="small"><%=arrTransports(17)%></td>
					<td align="center" class="small"><%=TotTransports%></td>
				</tr>
				</span>
			</table>
		</center>
	</body>
</html>
<%
Conn.Close
%>
