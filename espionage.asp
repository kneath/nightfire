<%Option Explicit%>
<%Response.Buffer = TRUE%>
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


<!--#Include File = "ad.asp"-->

<%
'Coded entirely by Duke Brak by hand
'This code may not be used in any other games than NightFire, a game by WarpSpire.
'www.warpspire.com/nightfire/
'Version 1.1 Alpha

'This is the espionage script

'My personalized minimum function
Dim x, y, z, r, final

Function MIN(x, y, z, r)
	If x<y AND x<z AND x<r Then
		MIN = Int(x)
	ElseIf y<x AND y<z AND y<r Then
		MIN = Int(y)
	ElseIf z<x AND z<y AND z<r Then
		MIN = Int(z)
	ElseIf r<x AND r<y AND r<z Then
		MIN = Int(r)
 ElseIf x=y OR x=z OR x=r Then
  MIN = Int(x)
 ElseIf y=z OR y=r Then
  MIN = Int(y)
 ElseIf z=r Then
  MIN = Int(z)
	End If
End Function

'Get Their UserName, The system they have and if they have arrived or not
Dim UserName, SelectedSystem, Arrived, Message
UserName = Request.Cookies("NightFire")("UserName")
Arrived = Request.Form("Arrived")

'Open up a connection to the database
Call OpenConnection()

'Declare the Recordsets and Open up the one to see what system they are in
Dim RST, RSD, RSGI, RSTemp, RSGS
Set RSGI = Server.CreateObject("ADODB.Recordset")
Set RST = Server.CreateObject("ADODB.Recordset")
Set RSD = Server.CreateObject("ADODB.Recordset")
Set RSTemp = Server.CreateObject("ADODB.Recordset")
RSGI.Open "SELECT * FROM Users WHERE UserName = '"&UserName&"'", Conn
Set RSGS = Server.CreateObject("ADODB.Recordset")

'Determine which system they have selected
Dim UserSystem
UserSystem = RSGI("System")
If Request.Form("System") <> Request.Cookies("NightFire")("SelectedSystem") AND Request.Form("System") <> ""  Then
	Response.Cookies("NightFire")("SelectedSystem") = Request.Form("System")
End If
If Request.Cookies("NightFire")("SelectedSystem") = ""  Then
	Response.Cookies("NightFire")("SelectedSystem") = UserSystem
End If
	'they have some other system in there that needs to stay.
SelectedSystem = Request.Cookies("NightFire")("SelectedSystem")
RSGI.Close

'Open the remainding recordsets
RST.Open "SELECT * FROM TRAINING WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
RSD.Open "SELECT * FROM DONEMILITARY WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
RSTemp.Open "SELECT * FROM GENERALINFO WHERE SystemNumber = "&SelectedSystem, Conn
RSGS.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic

'Determine the amount of resources they have
Dim ACapital, ACivillian
ACapital = RSGS("Capital") + 0
ACivillian = RSGS("Civillian") + 0

'See how many probes and agents are in training and done
Dim ProbesT, AgentsT, ProbesD, AgentsD
If RST.EOF = False Then
	RST.MoveFirst
	Do While Not RST.EOF
		ProbesT = ProbesT + RST("Probes") + 0
		AgentsT = AgentsT + RST("Agents") + 0
		RST.MoveNext
	Loop
Else
	ProbesT = 0
	AgentsT = 0
End If
ProbesD = RSD("Probes") + 0
AgentsD = RSD("Agents")

'Determine the maximum amount of probes they can build
Dim ProbesM, AgentsM, Num1, Num2, Num3, Num4
Num4 = 1000000000000000000000
Num3 = 1000000000000000000000
Num2 = 1000000000000000000000
Num1 = Int(ACapital/150)
ProbesM = MIN(Num1, Num2, Num3, Num4)
Num2 = ACivillian + 0
Num1 = Int(ACapital/250) + 0
AgentsM = MIN(Num1, Num2, Num3, Num4)

'See what they ordered
Dim RProbes, RAgents, RCapital, RCivillian
RProbes = Request.Form("Probes") + 0
RAgents = Request.Form("Agents") + 0
RCapital = (RProbes * 150) + (RAgents * 250)
RCivillian = RAgents + 0

'Check for dumbasses :P
If Arrived = "Yes" Then
	If RCapital > ACapital Then
		Message = "<p align = 'center' class = 'normal'>You don't have enough credits to build that many probes or agents</p>"
	ElseIf RCivillian > ACivillian Then
		Message = "<p align = 'center' class = 'normal'>You don't have enough civillians to train that many agents</p>"
	Else
		RST.AddNew
		RST("UserName") = UserName
		RST("Soldiers") = 0
		RST("Elite1") = 0
		RST("Elite2") = 0
		RST("Elite3") = 0
		RST("Transports") = 0
		RST("Probes") = RProbes
		RST("Agents") = RAgents
		RST("Mages") = 0
		RST("Time") = DateAdd("h", 8, Now())
		RST.Update
		RSGS("Capital") = ACapital - RCapital
		RSGS("Civillian") = ACivillian - RCivillian
		CapitalS = -1*RCapital
		RSGS.Update
		Message = "<p align = 'center' class = 'normal'>You have ordered your factories to construct "&FormatNumber(RProbes, 0)&" probes and "&FormatNumber(RCivillian, 0)&" civillians have began their training to become agents. This has cost you "&FormatNumber(RCapital, 0)&"c capial.</p>"
	End If
End If
RST.Close
RSGS.Close
RSD.Close
%>
<%

 'Okay let's start with the actual ops
Dim Op, TargetNation, TT

TargetNation = Request.Form("Nation")
Op = Request.Form("Operation")
If Request.Form("WhatDo") = "Operation" AND  Op <> "" AND TargetNation <> "" Then
	Dim TargetUserName, TargetProbes, TargetAgents
	
	RSTemp.Filter = "Nation = '"&TargetNation&"'"
	TargetUserName = RSTemp("UserName")
	Dim UserNameU, Protection
	Protection = RSTemp("Protection")
	UserNameU = TargetUserName

%>
<!--#Include File = "targetupdate.asp"-->

<!--#Include File = "statbar.asp"-->
<%

	TT = Conn.Execute("SELECT Probes, Agents, UserName FROM DONEMILITARY WHERE Username = '"&TargetUserName&"'")
	TargetProbes = TT("Probes") + 1
	TargetAgents = TT("Agents") + 1

Dim Go, ratio
ratio = (ProbesD + AgentsD*3)/(TargetProbes + TargetAgents*3)
Randomize()
If rnd() < .5 Then
	ratio = ratio + ratio*(.2*rnd())
	If ratio >= 1 Then
		Go = "Go"
	End If
Else
	ratio = ratio - ratio*(.2*rnd())
	If ratio >= 1 Then
		Go = "Go"
	End If	
End If

If DateDiff("h", Now(), Protection) > 0 Then
	Go = "NoGo"
End If

If Go = "Go" Then
Select Case Op
Case "Barrack Spy"
	Dim UserNameM
	UserNameM = TargetUserName
%>
<!--#Include File = "militaryprofile_code.asp"-->
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
			<br>
			<br>
<%
	Case "Civilian Terrorism"
	
	Dim TCivillians, MaxKill, RSM
	Set RSM = Server.CreateObject("ADODB.Recordset")

	RSGS.Open "SELECT UserName, Civillian, Population FROM generalstats WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
	TCivillians = RSGS("Civillian") 
	MaxKill = MIN(Int(TCivillians*.15), AgentsD*200, 1000000000000000, 10000000000000)
	
	RSGS("Civillian") = TCivillians - MaxKill
	RSGS("Population") = RSGS("Population") - MaxKill
	RSGS.Update
	
	Message = "<p class = 'normal'>Your agents successfully infiltrated "&TargetNation&"'s spy network.  They successfully detonated bombs in various cities killing <span class = 'alert'>"&FormatNumber(MaxKill, 0)&"</span> civillians!</p>"
	
	RSM.Open "MESSAGES", Conn, adOpenStatic, adLockPessimistic
	RSM.AddNew
	RSM("UserName") = TargetUserName
	RSM("Type") = "MainAttack"
	RSM("Seen") = "No"
	RSM("Time") = Time()
	RSM("Message") = "<p class = 'normal'>We have found that enemy agents have infiltrated our cities! They successfully detonated bombs in various cities killing <span class = 'alert'>"&FormatNumber(MaxKill, 0)&"</span> civillians!</p>"
	RSM.Update
	RSM.Close
	RSGS.Close
Case "Spy on Research"
	Dim RSR, RSRR
	Set RSR = Conn.Execute("SELECT * FROM RESEARCH WHERE UserName = '"&TargetUserName&"'")
	Set RSRR = Conn.Execute("SELECT * FROM RESEARCHING WHERE UserName = '"&TargetUserName&"'")
	Dim Elvl, Mlvl, Slvl, Alvl, Epts, Mpts, Spts, Apts, Eptsr, Mptsr, Sptsr, Aptsr
	Elvl = RSR("Energy")
	Mlvl = RSR("Military")
	Slvl = RSR("Stock")
	Alvl = RSR("Atomic")
	Epts = RSR("EnergyPts")
	Mpts = RSR("MilitaryPts")
	Spts = RSR("StockPts")
	Apts = RSR("AtomicPts")
	Eptsr = 0
	Mptsr = 0
	Sptsr = 0
	Aptsr = 0
	Do While RSRR.EOF = FALSE
		Eptsr = Eptsr + RSRR("Energy")
		Mptsr = Mptsr + RSRR("Military")
		Sptsr = Sptsr + RSRR("Stock")
		Aptsr = Aptsr + RSRR("Atomic")
		RSRR.MoveNext
	Loop
	RSR.Close
	RSRR.Close
%>
 <center>
<div class=normal>Your probes successfully infiltrate <%=TargetNation%>'s researching facilities and return with the following data.</div>
<table width='60%' align = 'center'>
<tr bgcolor = '#0000CC'>
<td align=center>Area</td>
<td align=center>Research Level</td>
<td align=center>Credits Researched</td>
<td align=center>Credits Researching</td>
</tr>
<tr bgcolor = '#000066'>
<td align=center>Energy</td>
<td align=center><%=Elvl%></td>
<td align=center><%=Epts%></td>
<td align=center><%=Eptsr%></td>
</tr>
<tr bgcolor = '#000033'>
<td align=center>Military</td>
<td align=center><%=Mlvl%></td>
<td align=center><%=Mpts%></td>
<td align=center><%=Mptsr%></td>
</tr>
<tr bgcolor = '#000066'>
<td align=center>Stock</td>
<td align=center><%=Slvl%></td>
<td align=center><%=Spts%></td>
<td align=center><%=Sptsr%></td>
</tr>
<tr bgcolor = '#000033'>
<td align=center>Atomic</td>
<td align=center><%=Alvl%></td>
<td align=center><%=Apts%></td>
<td align=center><%=Aptsr%></td>
</tr> 
</table>
</center>
<br><br>
<%
Case "Steal Technology"
	Set RSR = Server.CreateObject("ADODB.Recordset")
	RSR.Open "SELECT * FROM RESEARCH WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
	Dim Ecredits, Mcredits, Scredits, Acredits
	Ecredits = Int(RSR("EnergyPts")*.07)
	Mcredits = Int(RSR("MilitaryPts")*.07)
	Scredits = Int(RSR("StockPts")*.07)
	Acredits = Int(RSR("AtomicPts")*.07)
	RSR("EnergyPts") = RSR("EnergyPts") - Ecredits
	RSR("MilitaryPts") = RSR("MilitaryPts") - Mcredits
	RSR("StockPts") = RSR("StockPts") - Scredits
	RSR("AtomicPts") = RSR("AtomicPts") - Acredits
	RSR.Update
	RSR.Close
	RSR.Open "SELECT * FROM RESEARCH WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
	RSR("EnergyPts") = RSR("EnergyPts") + Ecredits
	RSR("MilitaryPts") = RSR("MilitaryPts") + Mcredits
	RSR("StockPts") = RSR("StockPts") + Scredits
	RSR("AtomicPts") = RSR("AtomicPts") + Acredits
	RSR.Update
	RSR.Close
	RSR.Open "MESSAGES", Conn, adOpenStatic, adLockPessimistic
	RSR.AddNew
	RSR("UserName") = TargetUserName
	RSR("Type") = "MainAttack"
	RSR("Seen") = "No"
	RSR("Time") = Now()
	RSR("Message") = "<span class = 'alert'>Alert!</span><span class = 'normal'>Enemy agents have infiltrated our researching facilities! We lost "&Ecredits+Mcredits+Scredits+Acredits&" credits worth of research!</span>"
	RSR.Update
	RSR.Close
	Message = "<center><span class = 'alert'>Success!</span><span class = 'normal'> Our Agents have infiltrated "&TargetNation&"'s researching facilities and have returned with "&Ecredits+Mcredits+Scredits+Acredits&" credits worth of research!</span>"

Case "Hack Civilian Security"
	Set RSR = Server.CreateObject("ADODB.Recordset")
	RSR.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
	RSR("Health") = RSR("Health") - 7
	RSR.Update
	Message = "<P class = 'normal'>Your probes successfully penatrate "&TargetNation&"'s Civilian Security and cause disruptions in the government. In turn, civilian health across the nation drops by 7%.</p>"

Case "Nuke Detection"
	Set RSR = Server.CreateObject("ADODB.Recordset")
	RSR.Open "SELECT * FROM NUKES WHERE UserName = '"&TargetUserName&"'", Conn
	Dim ABombs, HBombs, NBombs
	ABombs = RSR("AtomBombs")
	HBombs = RSR("HBombs")
	NBombs = RSR("NukeBombs")
	If ABombs = 0 Then
		ABombs = "no"
	End If
	If HBombs = 0 Then
		HBombs = "no"
		ABombs = "no"
	End If
	If NBombs = 0 Then
		NBombs = "no"
	End If
	
	Message = "<p class = 'normal'>Your probes return with the following information:<br>" & TargetNation & " has " & ABombs & " Atom Bombs, " & HBombs & " Hydrogen Bombs, and " & NBombs & " Nuclear Missiles.</p>"
End Select
Else
	Message = "<p class = 'normal'>Your covert operations department have failed the task</p>"
End If
	RSTemp.Filter = adFilterNone
	RSTemp.MoveFirst
End If
If DateDiff("h", Now(), Protection) > 0 Then
	Message = "<p class = 'alert'>That nation is under protection, no covert operations are allowed upon it</p>"
End If
		

%>


<%=Message%>

<table cellspacing=0 cellpadding=5 width="100%">
	<tr>
		<td bgcolor="#0000CC" colspan=5><div align=center class="large">Covert Operations</div></td>
	</tr>
	<tr>
		<td bgcolor="#000066" colspan=5><form action = "#" name = "statbar1"><center><input type=text size=50 value="Put mouse over to see cost" class = "statbar" name = "statbar11"><br><input type=text size=50 value="put mouse over unit to see stats" class = "statbar" name = "statbar12"></center></div></form></td>
	</tr>
	<tr>
		<td bgcolor="#000033" colspan=5><form action = "espionage.asp" method = "Post" name = "system"><input type = "hidden" name = "Arrived" value = "Yes"><input type = "hidden" name = "Arrived" value = "Yes"><center>Nation:<select name = "Nation" class = "numbar"><option value = "">Choose a Nation</option>


<%

Dim NationT
If RSTemp.EOF = False Then
	RSTemp.MoveFirst
	Do While Not RSTemp.EOF
		NationT = RSTemp("Nation")
		Response.Write "<option value = '"&NationT&"'>"&NationT&"</option>"
		RSTemp.MoveNext
	Loop
End If

%>
</select> &nbsp;&nbsp; System: <input name = "System" class = "numbar" size=2 value = "<%=SelectedSystem%>"> &nbsp;&nbsp; <input type = "submit" class = "button" value = "Change System"></center></td>
	</tr>
	<tr>
		<td bgcolor="#000033" colspan=5><input type = "hidden" name = "WhatDo" value = "Operation"><center>Select Operation: <select name = "Operation" class = "numbar"><option value = "">Choose an Operation</option><option value = "Barrack Spy" >Barrack Spy</option><option value = "Civilian Terrorism">Civilian Terrorism</option><option value = "Spy on Research">Spy on Research</option><option value = "Steal Technology">Steal Technology</option><option value = "Hack Civilian Security">Hack Civilian Security</option><option value = "Nuke Detection">Nuke Detection</option><option value = "Set off Nuke">Set off Nuke</option><option value = "Drain Energy">Drain Energy</option></select>&nbsp;&nbsp;<input type = "submit" class = "button" value = "Commence Operation"></center></form></td>
	</tr>
</table>
<br>
<br>
<br>

<form action = "espionage.asp" method = "post" name = "Train"><input type = "hidden" name = "Arrived" value = "Yes">
<table cellspacing=0 cellpadding=5 width="100%">
	<tr>
		<td bgcolor="#0000CC" colspan=5><div align=center class="large">Build</div></td>
	</tr>
	<tr bgcolor="#000044">
			<td>
				<div>Unit</div>
			</td>
			<td>
				<div>Processed</div>
			</td>
			<td>
				<div>Ordering</div>
			</td>

			<td><div>Max</div></td>

			<td><div>Order</div></td>

		</tr>
		<tr bgcolor="#000044">
			<td><div><a href = "#" onMouseOver="document.forms['statbar1'].elements['statbar11'].value='150c | 6hours'; document.forms['statbar1'].elements['statbar12'].value='A covert operations drone';">Probe</a></div></td>
			<td>
				<div><%=FormatNumber(ProbesD, 0)%></div>
			</td>

			<td>
				<div><%=FormatNumber(ProbesT, 0)%></div>
			</td>

			<td>
				<div><%=FormatNumber(ProbesM, 0)%></div>
			</td>

			<td>
				<div><input type="text" size=10 value="0" name="Probes" class = "numbar"></div>
			</td>

		</tr>
		<tr bgcolor="#000044">
			<td><div><a href = "#" onMouseOver="document.forms['statbar1'].elements['statbar11'].value='250c | 10hours | 1 civillian'; document.forms['statbar1'].elements['statbar12'].value='A covert operations agent';">Agent</a></div></td>
			<td>
				<div><%=FormatNumber(AgentsD, 0)%></div>
			</td>

			<td>
				<div><%=FormatNumber(AgentsT, 0)%></div>
			</td>

			<td>
				<div><%=FormatNumber(AgentsM, 0)%></div>
			</td>

			<td>
				<div><input type="text" size=10 value="0" name="Agents" class = "numbar"></div>
			</td>

		</tr>
		<tr bgcolor="#000066">
			<td colspan=5><div><center><input type = "submit" class = "button" value = "Train"></form></center></div></td>
		</tr>
</table>
</body>
</html>