<%Option Explicit%>
<!--#Include File = "include.asp"-->
<%

Dim UserName
UserName = Request.Cookies("NightFire")("UserName")

Dim RSGI
Set RSGI = Server.CreateObject("ADODB.Recordset")
RSGI.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserName&"'", Conn

'Determine which system they have selected
Dim UserSystem, SelectedSystem
UserSystem = RSGI("SystemNumber")
If Request.Form("System") <> Request.Cookies("NightFire")("SelectedSystem") AND Request.Form("System") <> ""  Then
	Response.Cookies("NightFire")("SelectedSystem") = Request.Form("System")
End If
If Request.Cookies("NightFire")("SelectedSystem") = ""  Then
	Response.Cookies("NightFire")("SelectedSystem") = UserSystem
End If
	'they have some other system in there that needs to stay.
SelectedSystem = Request.Cookies("NightFire")("SelectedSystem")
RSGI.Close
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

<form action = "nuke.asp" method = "post" name = "statbar1">
<%
Dim Nation, RSTemp, TUserName
Set RSTemp = Server.CreateObject("ADODB.Recordset")
Nation = Request.Form("Nation")
Dim RS
Set RS = Server.CreateObject("ADODB.Recordset")


If Request.Form("Arrived") = "Yes" and Nation <> "" Then
	'Damages, *B is buildings, *C is civillians
	Dim Error, Message, BombName, ACDamage, ABDamage, HCDamage, HBDamage, NCDamage, NBDamage, CDamage, BDamage, Buildings, Lands
	Buildings = 0
	Lands = 0
	ACDamage = 10000
	ABDamage = 50
	HCDamage = 100000
	HBDamage = 250
	NCDamage = 1000000
	NBDamage = 1000
	RS.Open "SELECT * FROM NUKES WHERE UserName = '"&UserName&"'", Conn, adOpenStatic, adLockPessimistic
	Select Case Request.Form("Bomb")
	Case "AtomBomb"
		BombName = "Atom Bomb"
		CDamage = ACDamage
		BDamage = ABDamage
		If RS("AtomBombs") < 1 Then
			Error = TRUE
		Else
			RS("AtomBombs") = RS("AtomBombs") - 1
		End If
	Case "HBomb"
		BombName = "Hydrogen Bomb"
		CDamage = HCDamage
		BDamage = HBDamage
		If RS("HBombs") < 1 Then
			Error = TRUE
		Else
			RS("HBombs") = RS("HBombs") - 1
		End If
	Case "NukeBomb"
		BombName = "Nuclear Missile"
		CDamage = NCDamage
		BDamage = NBDamage
		If RS("NukeBombs") < 1 Then
			Error = TRUE
		Else
			RS("NukeBombs") = RS("NukeBombs") - 1
		End If
	End Select
	RS.Update
	RS.Close
	If Error = TRUE Then
		Message = "<p class = 'alert'>You do not have enough bombs</p>"
	Else
	'Git der username
	Dim Protection
	RSTemp.Open "SELECT * FROM GENERALINFO WHERE Nation = '"&Nation&"'", Conn
	Protection = RSTemp("Protection")
	TUserName = RSTemp("UserName")
	RSTemp.Close
	If DateDiff("h", Now(), Protection) > 0 Then
		Message = "<p class = 'alert'>This nation is under protection, no nuclear weapons may be used</p>"
	Else
	RS.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&TUserName&"'", Conn, adOpenStatic, adLockPessimistic
	If CDamage > RS("Civillian") Then
		CDamage = RS("Civillian") - 1000
		RS("Civillian") = 1000
	Else
		RS("Civillian") = RS("Civillian") - CDamage
	End If
	Select Case Request.Form("Bomb")
	Case "AtomBomb"
		RS("Energy") = RS("Energy") - 50000
	Case "HBomb"
		RS("Energy") = RS("Energy") - 100000
	Case "NukeBomb"
		RS("Energy") = RS("Energy") - 150000
	End Select
	RS.Update
	RS.Close
	RS.Open "SELECT * FROM BUILDING WHERE UserName = '"&TUserName&"'", Conn, adOpenStatic, adLockPessimistic
	If RS.EOF = FALSE Then
	Do While Not RS.EOF
		Buildings = Buildings + RS("Barracks") + RS("Factories") + RS("SolarPanels") + RS("Consulates") + RS("ReasearchCenters") + RS("ElementXMines") + RS("DefenseStations") + RS("AtomicPlants") + RS("AntimatterPlants") + RS("PlutoniumPlants") + RS("AntimatterLabs") + RS("NuclearSilos") + RS("Temples") + RS("Cities")
		Lands = Lands + RS("PlainLand")
		RS.Delete
		RS.Update
		Rs.MoveNext
	Loop
	End If
	RS.Close

	Message = "<p class = 'normal'>The "&BombName&" was detonated inside the borders of "&Nation&". <span class = 'alert'>"&CDamage&"</span> civilians were killed instantly, <span class = 'alert'>"&Buildings&"</span> buildings in construction were destoryed, and <span class = 'alert'>"&Lands&"</span>km of land in exploration was destroyed.</p>"
	RS.Open "MESSAGES", Conn, 3, 3
	RS.AddNew
	RS("UserName") = TUserName
	RS("Message")  = "<p class = 'normal'>A fellow nation has launched their "&BombName&" in our borders!. <span class = 'alert'>"&CDamage&"</span> civilians were killed instantly, <span class = 'alert'>"&Buildings&"</span> buildings in construction were destoryed, and <span class = 'alert'>"&Lands&"</span>km of land in exploration was destroyed.</p>"
	RS("Type") = "MainAttack"
	RS("Seen") = "No"
	RS("Time") = Now()
	RS.Update
	RS.Close
End If
End If
End If
%>

<%=Message%>


<table width = "100%">
<tr>
		<td bgcolor="#0000CC" colspan=4><div align=center class="large">Nuclear Operations</div></td>
	</tr>
	<tr>
		<td bgcolor="#000066" colspan=4><center><input type=text size=50 value="Put mouse over to see cost" class = "statbar" name = "statbar11"><br><input type=text size=50 value="put mouse over unit to see stats" class = "statbar" name = "statbar12"></center></div></td>
	</tr>
	<tr bgcolor = "#000033">
		<td width = "30%" align = "center">Missile</td>
		<td width = "20%" align = "center">Quantity</td>
		<td width = "15%" align = "center">Launch</td>
		<td width = "25%" align = "center">Energy Required</td>
	</tr>
<%
RS.Open "SELECT * FROM NUKES WHERE UserName = '"&UserName&"'", Conn

Dim Atom, HBomb, Nuke, Atoms, HBombs, Nukes

Dim RST
Set RST = Conn.Execute("SELECT * FROM RESEARCH WHERE UserName = '"&UserName&"'")

Select Case RST("Atomic")
Case 0
Atom = FALSE
HBomb = FALSE
Nuke = FALSE
Case 1
Atom = TRUE
Case 2
Atom = TRUE
HBomb = TRUE
Case 3
Atom = TRUE
HBomb = TRUE
Nuke = TRUE
End Select

If Atom = TRUE Then
Atoms = RS("AtomBombs")
%>
	<tr bgcolor = "#000066">
		<td><a href = "#" onMouseOver="document.forms['statbar1'].elements['statbar11'].value='Kills <%=formatNumber(ACDamage, 0)%> Civilians'; document.forms['statbar1'].elements['statbar12'].value='Destroys All Buildings In Construction';">Atom Bomb</a></td>
		<td align = "center"><%=Atoms%></td>
		<td align = "center"><input type = "radio" class = "button" name = "Bomb" value = "AtomBomb" checked ></td>
		<td>50,000j</td>
	</tr>
<%
End If

If HBomb = TRUE Then
HBombs = RS("HBombs")
%>
	<tr bgcolor = "#000044">
		<td><a href = "#" onMouseOver="document.forms['statbar1'].elements['statbar11'].value='Kills <%=FormatNumber(HCDamage, 0)%> Civilians'; document.forms['statbar1'].elements['statbar12'].value='Destroys All Buildings In Construction';">Hydrogen Bomb</a></td>
		<td align = "center"><%=HBombs%></td>
		<td align = "center"><input type = "radio" class = "button" name = "Bomb" value = "HBomb"></td>
		<td>100,000j</td>
	</tr>
<%
End If

If Nuke = TRUE Then
Nukes = RS("NukeBombs")
%>
	<tr bgcolor = "#000066">
		<td><a href = "#" onMouseOver="document.forms['statbar1'].elements['statbar11'].value='Kills <%=FormatNumber(NCDamage, 0)%> Civilians'; document.forms['statbar1'].elements['statbar12'].value='Destroys All Buildings In Construction';">Nuclear Warhead</a></td>
		<td align = "center"><%=Nukes%></td>
		<td align = "center"><input type = "radio" class = "button" name = "Bomb" value = "NukeBomb"></td>
		<td>150,000j</td>
	</tr>
<%
End If

RS.Close

RSTemp.Open "SELECT * FROM GENERALINFO WHERE SystemNumber = "&SelectedSystem, Conn
%>
<input type = "hidden" name = "Arrived" value = "Yes">
	<tr bgcolor = "#000033">
		<td colspan=4>System: <input type = "text" size = "3" name = "System" value = "<%=SelectedSystem%>" class = "numbar"> <input type = "submit" class = "button" value = "Change">
Nation:<select name = "Nation" class = "numbar"><option value = "">Choose a Nation</option>
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
</select>
<input type = "submit" class = "button" value = "Launch"></form>
		</td>
	</tr>
</table>
