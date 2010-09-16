<%Option Explicit%>
<%Response.Buffer = FALSE%>
<!--#Include File = "include.asp"-->
<%
'Coded entirely by Duke Brak by hand
'This code may not be used in any other games than NightFire, a game by WarpSpire.
'www.warpspire.com/nightfire/
'Version 1.1 Alpha

Dim UserName
UserName = Request.Cookies("Nightfire")("UserName")

Dim RSTemp, RST
Set RST = Server.CreateObject("ADODB.Recordset")
RST.Open "SELECT UserName, SystemNumber FROM GENERALINFO WHERE UserName = '"&UserName&"'", Conn

'Determine which system they have selected
Dim UserSystem, SelectedSystem
UserSystem = RST("SystemNumber")
If Request.Form("System") <> Request.Cookies("NightFire")("SelectedSystem") AND Request.Form("System") <> "" Then
	Response.Cookies("NightFire")("SelectedSystem") = Request.Form("System")
End If
If Request.Cookies("NightFire")("SelectedSystem") = ""  Then
	Response.Cookies("NightFire")("SelectedSystem") = UserSystem
End If
	'they have some other system in there that needs to stay.
SelectedSystem = Request.Cookies("NightFire")("SelectedSystem")
RST.Close
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
<!--#include file="statbar.asp"-->
<%

Set RSTemp = Conn.Execute("SELECT a.*, b.* FROM GENERALINFO a, GENERALSTATS b WHERE a.UserName = b.UserName AND a.SystemNumber = "&SelectedSystem&" ORDER BY Land DESC")

Dim Temples, Land
RST.Open "SELECT a.*, b.* FROM DONEBUILDINGS a, GENERALSTATS b WHERE a.UserName = b.UserName AND a.UserName = '"&UserName&"'", Conn
Temples = RST("Temples")
Land = RST("Land")
Temples = FormatNumber(Temples/Land, 4)
RST.Close

If Request.Form("Action") = "Pray" AND Request.Form("Prayer") <> "" Then
	Dim TargetUserName, TargetNation, Protection
	TargetNation = Request.Form("Nation")
	RST.Open "SELECT UserName, Nation, Protection FROM GENERALINFO WHERE Nation = '"&TargetNation&"'"
	TargetUserName = RST("UserName")
	Protection = RST("Protection")
	Dim RSTD
	Set RSTD = Server.CreateObject("ADODB.Recordset")
	RSTD.Open "SELECT a.*, b.* FROM DONEBUILDINGS a, GENERALSTATS b WHERE a.UserName = b.UserName AND a.UserName = '"&TargetUserName&"'", Conn
	Dim TTemples, TLand
	TTemples = RSTD("Temples")
	TLand = RSTD("Land")
	TTemples = TTemples/Land

	Dim Go, ratio
	If Temples = 0 Then
	   Temples = .0000000000000001
	ElseIf TTemples = 0 Then
	   TTemples = .0000000000000001
	End If
	ratio = Temples/TTemples
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
	Dim Message

	If Go = "NoGo" Then
		Message = "<p class = 'normal'>Your team of psychics failed the task.</p>"
	Else
		Dim Prayer
		Prayer = Request.Form("Prayer")	
		Select Case Prayer
			Case "MindReadMilitary"
				Dim RSTDM, RSTTM, RSTAM, RSTC
				Set RSTDM = Server.CreateObject("ADODB.Recordset")
				Set RSTTM = Server.CreateObject("ADODB.Recordset")
				Set RSTAM = Server.CreateObject("ADODB.Recordset")
				Set RSTC = Server.CreateObject("ADODB.Recordset")
				RSTDM.Open "SELECT a.*, b.UserName, b.Race FROM DONEMILITARY a, GENERALINFO b WHERE a.UserName = b.UserName AND a.UserName = '"&TargetUserName&"'", Conn
				RSTTM.Open "SELECT * FROM TRAINING WHERE UserName = '"&TargetUserName&"'", Conn
				RSTAM.Open "SELECT * FROM ATTACKING WHERE UserName = '"&TargetUserName&"'", Conn
				RSTC.Open "SELECT * FROM COSTS WHERE Race = '"&RSTDM("Race")&"'", Conn, adOpenStatic, adLockPessimistic
				
				Dim Infantry, Elite1, Elite2, Elite3, Transports, AInfantry, AElite1, AElite2, AElite3, ATransports, TInfantry, TElite1, TElite2, TElite3, TTransports
				
				Infantry = RSTDM("Soldiers")
				Elite1 = RSTDM("Elite1")
				Elite2 = RSTDM("Elite2")
				Elite3 = RSTDM("Elite3")
				Transports = RSTDM("Transports")
				
				If RSTTM.EOF = FALSE Then
				   Do While Not RSTTM.EOF
				   	  TInfantry = TInfantry + RSTTM("Soldiers")
					  TElite1 = TElite1 + RSTTM("Elite1")
				   	  TElite2 = TElite2 + RSTTM("Elite2")
				  	  TElite3 = TElite3 + RSTTM("Elite3")
				 	  TTransports = TTransports + RSTTM("Transports")
				  	  RSTTM.MoveNext
				   Loop
				End If
				
				If RSTAM.EOF = FALSE Then
				Do While Not RSTAM.EOF
				   AInfantry = AInfantry + RSTAM("Soldiers")
				   AElite1 = AElite1 + RSTAM("Elite1")
				   AElite2 = AElite2 + RSTAM("Elite2")
				   AElite3 = AElite3 + RSTAM("Elite3")
				   ATransports = ATransports + RSTAM("Transports")
				   RSTAM.MoveNext
				Loop
				End If
				
				Dim Elite1N, Elite2N, Elite3N
				RSTC.Filter = "Unit = 'Elite1'"
				Elite1N = RSTC("Name")
				RSTC.Filter = "Unit = 'Elite2'"
				Elite2N = RSTC("Name")
				RSTC.Filter = "Unit = 'Elite3'"
				Elite3N = RSTC("Name")
%>
<p class = "normal" align = "center">Your psychics were able to penetrate the enemy's minds. They offer you a summary of their forces.</p>
<table width = "50%" align = "center">
	   <tr bgcolor = "#0000CC">
	   	   <td>Unit</td>
		   <td>Units Home</td>
		   <td>Units Away</td>
		   <td>Units Traning</td>
	   </tr>
	   <tr>
	   	   <td bgcolor = "#000066">Infantry</td>
		   <td bgcolor = "#000033"><%=FormatNumber(Infantry, 0)%></td>
		   <td bgcolor = "#000066"><%=FormatNumber(AInfantry, 0)%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(TInfantry, 0)%></td>
	   </tr>
	   <tr>
	   	   <td bgcolor = "#000066"><%=Elite1N%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(Elite1, 0)%></td>
		   <td bgcolor = "#000066"><%=FormatNumber(AElite1, 0)%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(TElite1, 0)%></td>
	   </tr>
	   	   <tr>
	   	   <td bgcolor = "#000066"><%=Elite2N%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(Elite2, 0)%></td>
		   <td bgcolor = "#000066"><%=FormatNumber(AElite2, 0)%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(TElite2, 0)%></td>
	   </tr>
	   	   <tr>
	   	   <td bgcolor = "#000066"><%=Elite3N%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(Elite3, 0)%></td>
		   <td bgcolor = "#000066"><%=FormatNumber(AElite3, 0)%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(TElite3, 0)%></td>
	   </tr>
	   <tr>
	   	   <td bgcolor = "#000066">Transports</td>
		   <td bgcolor = "#000033"><%=FormatNumber(Transports, 0)%></td>
		   <td bgcolor = "#000066"><%=FormatNumber(ATransports, 0)%></td>
		   <td bgcolor = "#000033"><%=FormatNumber(TTransports, 0)%></td>
	   </tr>
</table>
	   
	   
	   
<%
  	   		  RSTDM.Close
			  RSTTM.Close
			  RSTAM.Close
			  RSTC.Close
			 
		Case "MindReadResearch"
			 Dim RSTR, RSTRR
			 Set RSTR = Server.CreateObject("ADODB.Recordset")
			 Set RSTRR = Server.CreateObject("ADODB.Recordset")
			 RSTR.Open "SELECT * FROM RESEARCH WHERE UserName = '"&TargetUserName&"'", Conn
			 RSTRR.Open "SELECT * FROM RESEARCHING WHERE UserName = '"&TargetUserName&"'", Conn
			 
			 Dim EnergyPtsR, MilitaryPtsR, StockPtsR, AtomicPtsR
			 Do While Not RSTRR.EOF = TRUE
			 	EnergyPtsR = EnergyPtsR + FormatNumber(RSTRR("Energy"), 0)
				MilitaryPtsR = MilitaryPtsR + FormatNumber(RSTRR("Military"), 0)
				StockPtsR = StockPtsR + FormatNumber(RSTRR("Stock"), 0)
				AtomicPtsR = AtomicPtsR + FormatNumber(RSTRR("Atomic"), 0)
				RSTRR.MoveNext
			 Loop
			 EnergyPtsR = FormatNumber(EnergyPtsR, 0)
			 MilitaryPtsR = FormatNumber(MilitaryPtsR, 0)
			 StockPtsR = FormatNumber(StockPtsR, 0)
			 AtomicPtsR = FormatNumber(AtomicPtsR, 0)
%>
<p class = 'normal' align = 'center'>Your psychics successfully retrieved the information you requested.</p>
<table width = "50%" align = "center">
	   <tr bgcolor = "#0000CC">
	   	   <td>Area of Research</td>
		   <td>Level</td>
		   <td>Credits</td>
		   <td>Credits in route</td>
	   <tr>
	   	   <td bgcolor = "#000066">Energy</td>
		   <td bgcolor = "#000033"><%= RSTR("Energy") %></td>
		   <td bgcolor = "#000066"><%= RSTR("EnergyPts") %></td>
		   <td bgcolor = "#000033"><%= EnergyPtsR %>
	   </tr>
	   	   <tr>
	   	   <td bgcolor = "#000066">Stock Market</td>
		   <td bgcolor = "#000033"><%= RSTR("Stock") %></td>
		   <td bgcolor = "#000066"><%= RSTR("StockPts") %></td>
		   <td bgcolor = "#000033"><%= StockPtsR %>
	   </tr>
	   	   <tr>
	   	   <td bgcolor = "#000066">Atomic Weapons</td>
		   <td bgcolor = "#000033"><%= RSTR("Atomic") %></td>
		   <td bgcolor = "#000066"><%= RSTR("AtomicPts") %></td>
		   <td bgcolor = "#000033"><%= AtomicPtsR %>
	   </tr>
</table>
	  
		   
<%
  		   RSTR.Close
		   RSTRR.Close
		Case "ForcedSuicide"
			 Dim RSTB, RSTG, RSTM
			 Set RSTB = Server.CreateObject("ADODB.Recordset")
			 Set RSTG = Server.CreateObject("ADODB.Recordset")
			 Set RSTM = Server.CreateObject("ADODB.Recordset")
			 RSTB.Open "SELECT * FROM DONEBUILDINGS WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
			 RSTG.Open "SELECT UserName, Population, Civillian FROM GENERALSTATS WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
			 RSTM.Open "MESSAGES", Conn, adOpenStatic, adLockPessimistic
			 
			 Dim Damage(1)
			 Damage(0) = Int(RSTG("Civillian") * .018)
			 RSTG("Civillian") = RSTG("Civillian") - Damage(0)
			 RSTG("Population") = RSTG("Population") - Damage(0)
			 RSTG.Update
			 
			 Dim i
			 For i = 2 to 14
			 	 Damage(1) = Damage(1) + Int(RSTB(i)*.013)
			 	 RSTB(i) = RSTB(i) - Int(RSTB(i)*.013)
			 Next
			 RSTB.Update
			 
			 RSTM.AddNew
			 RSTM("UserName") = TargetUserName
			 RSTM("Message") = "Enemy psychics have infiltrated the minds of our civilians! We lost <span class = 'alert'>"&Damage(0)&"</span> civilians due to suicide and <span class = 'alert'>"&Damage(1)&"</span> due to rioting."
			 RSTM("Type") = "MainAttack"
			 RSTM("Seen") = "No"
			 RSTM("Time") = Now()
			 RSTM.Update
			 
%>
<p align = "center" class = "normal"> Your psychics penetrate the minds of the enemy. They cause mass suicides and rioting all over the nation. The enemy looses <span class = "alert"><%=FormatNumber(Damage(0), 0)%></span> civillians and <span class = "alert"><%=Damage(1)%></span> buildings.</p>

<%
   		    RSTB.Close
			RSTG.Close
			RSTM.Close
		Case "MilitaryStrength"
			Dim RSTS
			Set RSTS = Server.CreateObject("ADODB.Recordset")
			Set RSTM = Server.CreateObject("ADODB.Recordset")
			RSTS.Open "SELECT * FROM Prayers", Conn, adOpenStatic, adLockPessimistic
			RSTM.Open "SELECT UserName, Morale FROM DONEMILITARY WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
			
			RSTS.AddNew
			RSTS("UserName") = UserName
			RSTS("Spell") = "Ares Guidence"
			RSTS("FinalTime") = DateAdd("h", 12, Now())
			RSTS.Update
			
			RSTM("Morale") = 100
			RSTM.Update
%>
<p class = 'normal' align = 'center'>Your psychics are glad to be able to aid your military.  They will be able to retain this state for 12 hours.</p>
<%
   		   RSTS.Close
		   RSTM.Close
		Case "InstillFear"	
			 Set RSTD = Server.CreateObject("ADODB.Recordset")
			 Set RSTM = Server.CreateObject("ADODB.Recordset")
			 RSTM.Open "MESSAGES", Conn, adOpenStatic, adLockPessimistic
			 RSTD.Open "SELECT UserName, Morale FROM DONEMILITARY WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
			 RSTD("Morale") = RSTD("Morale") - 3
			 RSTD.Update
			 
			 RSTM.AddNew
			 RSTM("UserName") = TargetUserName
			 RSTM("Message") = "Enemy psychics have striken fear into our military.  Military morale has been reduced by 3%."
			 RSTM("Type") = "MainAttack"
			 RSTM("Seen") = "No"
			 RSTM("Time") = Now()
			 RSTM.Update
			 
			 RSTM.Close
			 RSTD.Close
%>
<p class = 'normal' align = 'center'>Your psychics fill the minds of the enemy with nightmares and reduce overall military morale by 3%.</p>
<%
   		Case "ConvertPsi"
			 Set RSTG = Server.CreateObject("ADODB.Recordset")
			 RSTG.Open "SELECT Energy, UserName FROM GENERALSTATS WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
			 
			 RSTG("Energy") = RSTG("Energy") + 10000
			 RSTG.Update
%>
<p class = "normal">Your psychics are able to convert the power of psi into 10,000j of energy for use in your nation</p>
<%
			 RSTG.Close
			 
		Case "MassSuicide"
			 Set RSTG = Server.CreateObject("ADODB.Recordset")
			 Set RSTD = Server.CreateObject("ADODB.Recordset")
			 Set RSTM = Server.CreateObject("ADODB.Recordset")
			 RSTM.Open "MESSAGES", Conn, adOpenStatic, adLockPessimistic
			 RSTG.Open "SELECT UserName, Civillian, Population FROM GENERALSTATS WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
			 RSTD.Open "SELECT UserName, Soldiers FROM DONEMILITARY WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
			 
			 Dim Civilians
			 Civilians = RSTG("Civillian")
			 Infantry = RSTD("Soldiers")
%>
<p class = 'normal'>Your psychics cause mass suicides amoung civillians and lower level military infantry.  The enemy looses <%=Int(Civilians*.03)%> civilians and <%=Int(Infantry*.04)%> infantry.</p>
<%
			 Dim InfantryLoss, CivilianLoss
			 CivilianLoss = Int(Civilians*.03)
			 InfantryLoss = Int(Infantry*.04)
			 Infantry = Infantry - InfantryLoss
			 Civilians = Civilians - CivilianLoss
			 RSTG("Population") = RSTG("Population") - CivilianLoss
			 RSTG("Civillian") = Civilians
			 RSTD("Soldiers") = Infantry
			 RSTG.Update
			 RSTD.Update
			 
			 RSTM("UserName") = TargetUserName
			 RSTM("Message") = "Enemy psychics have infiltrated the minds of our civilians and infantry We lost <span class = 'alert'>"&CivilianLoss&"</span> and <span class = 'alert'>"&InfantryLoss&"</span> infantry due to mass suicides."
			 RSTM("Type") = "MainAttack"
			 RSTM("Seen") = "No"
			 RSTM("Time") = Now()
			 RSTM.Update
			 
			 RSTG.Close
			 RSTD.Close
			 RSTM.Close
		Case "EnlightenCivilians"
			 
		End Select
				
	End If
End If

%>
<br><br>
<form action = "magic.asp" method = "post">
<table width="95%" align="center">
	<tr>
		<td bgcolor="#0000CC" align = "center" class="large">Psionics Headquarters</td>
	</tr>
	<tr>
		<td bgcolor = "#0000CC" align = "left" class="small">Your psychics are ready to carry out your orders.</td>
	</tr>
	<tr>
		<td bgcolor="#000033" align = "center">Nation :  <select name = "Nation" class = "numbar"><option value="">-Choose a Nation-</option>
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
		</select> System : <input size=4 name = "System" value = "<%=SelectedSystem%>" class="numbar"> &nbsp;<input type = "submit" value = "Change System" class = "button"></td>
	</tr>
	<tr>
		<td bgcolor = "#000033" align = "center"><input type = "hidden" name = "Action" value = "Pray">Select Operation: <select name = "Prayer" class = "numbar">
<option value = "">-Choose an Operation-</option>
<option value = "MindReadMilitary">Mind Reading - Military</option>
<option value = "MindReadResearch">Mind Reading - Research</option>
<option value = "ForcedSuicide">Force Riots</option>
<option value = "MilitaryStrength">Military Strength</option>
<option value = "InstillFear">Instill Fear</option>
<option value = "ConvertPsi">Convert Psi</option>
<option value = "MassSuicide">Mass Suicide</option>
<option value = "EnlightenCivilians">Enlighten Civilians</option>
<option value = "InspireRevenge">Inspire Revenge</option>
		</select> &nbsp;&nbsp;<input type = "submit" class = "button" value = "Send out Orders"></td>
	</tr>
	<tr>
	<td bgcolor="#000033" align="center"><a href="/nightfire/guide/covertmagic.html" class="menu_main" target="new">Psionic Guide</a>*

</table>
</body>
</html>