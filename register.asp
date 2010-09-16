<% Option Explicit %>
<% Response.Buffer = "TRUE" %>
<!--#Include File = "adovbs.inc"-->
<html>
<head>
<title>Register a New User</title>
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

<%

If Request("new") = "" Then
%>
<table cellspacing=0 cellpadding=5 width = "75%" align = "center">
<tr><td bgcolor="#0000CC" colspan=15><p align = "center" class = "large">Register a New User!</p></td></tr>
<br>
<br>
<tr><td bgcolor="#000033" colspan=15>
<p align = "left" class = "normal">&nbsp;&nbsp;&nbsp;&nbsp; Welcome to NightFire a completely free interactive web-based strategey game. Currently, NightFire is still in it's testing stages therefore, there are bugs and not all pages exsist  quite yet.<br>Thank you for playing NightFire </p></td></tr>
<br>
<br>
</table>
<form action = "register.asp" method = "POST">

<table border="0" cellpadding="3" cellspacing="0" width="75%" align = "center">
<tr>
<td>Desired User Name:</td>
<td><input name = "UserName" size = "10" maxlength = "20" class = "numbar"></td>
</tr>
<tr>
<td>Desired Password:</td>
<td><input type = "password" name = "Password1" size = "8" maxlength = "20" class = "numbar"></td>
</tr>
<tr>
<td> Confirm Password</td>
<td><input type = "password" name = "Password2" size = "8" maxlenght = "20" class = "numbar"></td>
</tr>
<tr>
<td>Nation:</td>
<td><input name = "Nation" size = "20" maxlength = "30" class = "numbar"></td>
</tr>
<tr>
<td>Dictator:</td>
<td><input name = "Dictator" size = "15" maxlenth = "20" class = "numbar"></td>
</tr>
<tr>
<td>Race:</td>
<td>
<select name = "Race" class = "numbar">
<option value = "Human">Human</option>
<option value = "Larxon">Larxon</option>
<option value = "Rekkan">Rekkan</option>
<option value = "Ostarian">Ostarian</option>
<option value = "Volron">Volron</option>
</select>
</td>
</tr>
<tr>
<td>System Type:</td>
<td>
<select name = "SystemType" class = "numbar">
<option value = "Generic Star">Generic Star</option>
<option value = "Black Hole">Black Hole</option>
<option value = "Brown Dwarf">Brown Dwarf</option>
<option value = "Red Giant">Red Giant</option>
</select>
</td>
</tr>
<tr>
<br>
<td><input type = "hidden" value = "1" name = "new"></td>
<td><input type = "submit" value = "Register" class = "button">
</form>
</td></tr></table>
<% 
Else
Dim UserName, Password, ConfirmPass, Nation, Dictator, ErrorYet, SuccessMessage

UserName = Request("UserName")
UserName = Replace(UserName, "'", "")
Password = Request("Password1")
Password = Replace(Password, "'", "")
ConfirmPass = Request("Password2")
Nation = Request("Nation")
Nation = Server.HTMLEncode(Nation)
Nation = Replace(Nation, "'", "")
Dictator = Request("Dictator")
Dictator = Server.HTMLEncode(Dictator)
Dictator = Replace(Dictator, "'", "")

If isNull(UserName) Or isNull(Password) or isNull(ConfirmPass) or isNull(Nation) or isNull(Dictator) Then
	ErrorYet = 1
	SuccessMessage = "<p align = 'center' class = 'large'>You have not filled out all of the fields. Click the Back Button on your browser</p>"
End If

If ErrorYet <> 1 Then
	If Password <> ConfirmPass Then
		ErrorYet = 1
		SuccessMessage = "<p align = 'center' class = 'large'>Your passwords do not match. Click the Back Button on your browser</p>"
	End if
End If

If ErrorYet <> 1 Then

'Lets have fun making connections!

Dim Conn

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.ConnectionString="Provider=SQLOLEDB;User ID=sa;Password=Kt6261;Initial Catalog=warpspire"

Conn.Open

Dim RSGS, RSGI, RSL, RSDB, RSDM, RSU, RSR, RSN

Set RSGS = Server.CreateObject("ADODB.Recordset")
Set RSGI = Server.CreateObject("ADODB.Recordset")
Set RSL = Server.CreateObject("ADODB.Recordset")
Set RSDB = Server.CreateObject("ADODB.Recordset")
Set RSDM = Server.CreateObject("ADODB.Recordset")
Set RSU = Server.CreateObject("ADODB.Recordset")
Set RSR = Server.CreateObject("ADODB.Recordset")
Set RSN = Server.CreateObject("ADODB.Recordset")

RSGS.Open "SELECT * FROM GeneralStats", Conn, 3, 3
RSGI.Open "SELECT * FROM GeneralInfo", Conn, adOpenForwardOnly, adLockPessimistic
RSL.Open "SELECT * FROM Land", Conn, adOpenForwardOnly, adLockPessimistic
RSDB.Open "SELECT * FROM DoneBuildings", Conn, adOpenForwardOnly, adLockPessimistic
RSDM.Open "SELECT * FROM DoneMilitary", Conn, adOpenForwardOnly, adLockPessimistic
RSU.Open "SELECT * FROM Users", Conn, adOpenStatic, adLockPessimistic
RSR.Open "SELECT * FROM Research", Conn, adOpenStatic, adLockPessimistic
RSN.Open "SELECT * FROM Nukes", Conn, adOpenStatic, adLockPessimistic

Dim RSCheck, RSCheck2
Set RSCheck = Conn.Execute("SELECT UserName FROM Users WHERE UserName = '"&UserName&"'")
Set RSCheck2 = Conn.Execute("SELECT Nation FROM GeneralInfo WHERE Nation = '"&Nation&"'")

If RSU.RecordCount>300 Then 
	SuccessMessage = "<p align = 'center' class = 'large'>I'm sorry, but there is already 300 players</p>"
	Set RSCheck = Nothing
	Set RSCheck = Nothing
	RSGS.Close
	RSGI.Close
	RSL.Close
	RSDB.Close
	RSDM.Close
	RSU.Close
	Conn.Close
ElseIf RSCheck.EOF = FALSE Then
	SuccessMessage = "<p align = 'center' class = 'large'>I'm sorry, someone else has that username please click the Back Button on your browser</p>"
	Set RSCheck = Nothing
	Set RSCheck = Nothing
	RSGS.Close
	RSGI.Close
	RSL.Close
	RSDB.Close
	RSDM.Close
	RSU.Close
	Conn.Close
ElseIf RSCheck2.EOF = FALSE Then
	SuccessMessage = "<p align = 'center' class = 'large'>I'm sorry, someone else already has that Nation Name please click the Back Button on your browser and choose a different name.</p>"
	Set RSCheck = Nothing
	Set RSCheck = Nothing
	RSGS.Close
	RSGI.Close
	RSL.Close
	RSDB.Close
	RSDM.Close
	RSU.Close
	Conn.Close
Else
'That wasn't fun... now lets have fun declaring starting stuff! whoohoo!

Dim Population, CivillianPop, MilitaryPop, TotalLand, Networth, Capital, ElementX, Energy, Mana, Plutonium, Antimatter, PlainLand, ElementXLand, PlutoniumLand, AntimatterLand, SmoothLand, RuggedLand, FreePlain, Barracks, Factories, SolarPanels, Consulates, ReaserchCenters, ElementXMines, DefenseStations, AtomicPlants, AntimatterPlants, PlutoniumPlants, AntimatterLabs, NukeSilos, Temples, Cities, FactorySlots, Soldiers, Elite1, Elite2, Elite3, Transports, Mages, Probes, Agents, Morale, BarracksSlots, UsedBarracksSlots, UsedFactorySlots, System, Race, SystemType, Health, LastUpdate
Population = 31000
CivillianPop = 30000
MilitaryPop = Population - CivillianPop
TotalLand = 1000
Networth = 10000
Capital = 100000
ElementX = 10000
Energy = 50000
Mana = 10000
Plutonium = 10000
Antimatter = 10000
PlainLand = 1000
ElementXLand = 0
PlutoniumLand = 0
AntimatterLand = 0
SmoothLand = 0
RuggedLand = 0
FreePlain = 480
Barracks = 15
Factories = 5
SolarPanels = 20
Consulates = 0
ReaserchCenters = 0
ElementXMines = 80
DefenseStations = 0
AtomicPlants = 0
AntimatterPlants = 0
PlutoniumPlants = 0
AntimatterLabs = 0
NukeSilos = 0
Temples = 0
Cities = 400
BarracksSlots = 0
FactorySlots = 0
Soldiers = 500
Elite1 = 500
Elite2 = 0
Elite3 = 0
Transports = 0
Mages = 0
Probes = 0
Agents = 0
Morale = 100
BarracksSlots = 1200
FactorySlots = 400
UsedBarracksSlots = 1000
UsedFactorySlots = 0
SystemType = Request.Form("SystemType")
Race = Request.Form("Race")
Health = 100

Dim RSTT
Set RSTT = Server.CreateObject("ADODB.Recordset")
RSTT.Open "SELECT * FROM SYSTEMS WHERE SystemType = '"&SystemType&"' AND Nations < 10", Conn, adOpenStatic, adLockPessimistic
System = RSTT("SystemNumber")
RSTT("Nations") = RSTT("Nations") + 1
RSTT.Update
RSTT.Close

If DateDiff("h","6/8/2001 12:00 PM", Now()) > 0 Then
	LastUpdate = Now()
Else
	LastUpdate = "6/8/2001 12:00:00 PM"
End If

'That wasn't fun either.... oh well 

'GeneralStats Insertion

RSGS.AddNew
RSGS("UserName") = UserName
RSGS("Population") = Population
RSGS("Civillian") = CivillianPop
RSGS("Military") = MilitaryPop
RSGS("Land") = TotalLand
RSGS("Networth") = Networth 
RSGS("Capital") = Capital
RSGS("ElementX") = ElementX
RSGS("Energy") = Energy
RSGS("Mana") = Mana
RSGS("Plutonium") = Plutonium
RSGS("Antimatter") = AntiMatter
RSGS.Update
RSGS.Close

'GeneralInfo Insertion

RSGI.AddNew
RSGI("UserName") = UserName
RSGI("Nation") = Nation
RSGI("Minister") = Dictator
RSGI("Race") = Race
RSGI("SystemType") = SystemType
RSGI("SystemNumber") = System
RSGI("Health") = Health
RSGI("LastUpdate") = LastUpdate
RSGI("Protection") = DateAdd("h", 72, Now())
RSGI("King") = "No"
RSGI("KingChoice") = "-"
RSGI.Update
RSGI.Close

'Land insertion

RSL.AddNew
RSL("UserName") = UserName
RSL("Plain") = PlainLand
RSL("ElementX") = ElementXLand
RSL("Plutonium") = PlutoniumLand
RSL("Antimatter") = AntimatterLand
RSL("Smooth") = SmoothLand
RSL("Rugged") = RuggedLand
RSL("FreePlain") = FreePlain
RSL("LandBontime") = Now()
RSL.Update
RSL.Close

'DoneBuildings insertion

RSDB.AddNew
RSDB("UserName") = UserName
RSDB("Barracks") = Barracks
RSDB("Factories") = Factories
RSDB("SolarPanels") = SolarPanels
RSDB("Consulates") = Consulates
RSDB("ReasearchCenters") = ReaserchCenters
RSDB("ElementXMines") = ElementXMines
RSDB("DefenseStations") = DefenseStations
RSDB("AtomicPlants") = AtomicPlants
RSDB("AntimatterPlants") = AntimatterPlants
RSDB("PlutoniumPlants") = PlutoniumPlants
RSDB("AntimatterLabs") = AntimatterLabs
RSDB("NuclearSilos") = NukeSilos
RSDB("Temples") = Temples
RSDB("Cities") = Cities
RSDB.Update
RSDB.Close

'DoneMilitary insertion

RSDM.AddNew
RSDM("UserName") = UserName
RSDM("Soldiers") = Soldiers
RSDM("Elite1") = Elite1
RSDM("Elite2") = Elite2
RSDM("Elite3") = Elite3
RSDM("Transports") = Transports
RSDM("Mages") = Mages
RSDM("Probes") = Probes
RSDM("Agents") = Agents
RSDM("Morale") = Morale
RSDM("BarracksSlots") = BarracksSlots
RSDM("FactorySlots") = FactorySlots
RSDM("UsedBarracksSlots") = UsedBarracksSlots
RSDM("UsedFactorySlots") = UsedFactorySlots
RSDM.Update
RSDM.Close

'Users insertion

RSU.AddNew
RSU("UserName") = UserName
RSU("Password") = Password
RSU("System") = System
RSU.Update
RSU.Close

'Research stuff
RSR.AddNew
RSR("UserName") = UserName
RSR("Energy") = 0
RSR("Military") = 0
RSR("Stock") = 0
RSR("Atomic") = 0
RSR("EnergyPts") = 0
RSR("MilitaryPts") = 0
RSR("StockPts") = 0
RSR("AtomicPts") = 0
RSR("AttackPts") = 0
RSR("StealthPts") = 0
RSR.Update
RSR.Close

'Nukes
RSN.AddNew
RSN("UserName") = UserName
RSN("AtomBombs") = 0
RSN("HBombs") = 0
RSN("NukeBombs") = 0
RSN.Update
RSN.Close

Conn.Close

'All done with insertions

SuccessMessage = "<p align = 'center' class = 'normal'> Congragulations! You have successfully created  an account in NightFire</p><br><table cellspacing=0 cellpadding=5 width = '50%' align = 'center'><tr><td colspan=10 class = 'large' align = 'center' bgcolor = '#ee5900' >Your Stats</td></tr><tr><td bgcolor = '#000033'>Nation:</td><td class = 'normal' bgcolor = '#000066'>"&Nation&"</td></tr><tr><td bgcolor = '#000033'>Ruler:</td><td class = 'normal' bgcolor = '#000066'> "&Dictator&" </td></tr><tr><td bgcolor = '#000033'>System: <td class = 'normal' bgcolor = '#000066'> "&System&"</td></tr><tr><td bgcolor = '#000033'>User Name: <td class = 'normal' bgcolor = '#000066'>"&UserName&" </td></tr></table> <p>Thank you for helping NightFire become a reality! Please go to the boards and report any errors, or just to talk with the game developers</p>"

End If
End If
End If
%>
<%=SuccessMessage%>
<a href = "index.asp">Back to NightFire</a>
</body>
</html>
