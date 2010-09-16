<%@ Language=VBScript %>
<% Option Explicit %>
<!--#Include File = "include.asp"-->
<%
Response.Buffer =  TRUE

Function MIN(x, y, z, r)
	If x<y AND x<z AND x<r Then
		MIN = x
	ElseIf y<x AND y<z AND y<r Then
		MIN = y
	ElseIf z<x AND z<y AND z<r Then
		MIN = z
	ElseIf r<x AND r<y AND r<z Then
		MIN = r
 ElseIf x=y OR x=z OR x=r Then
  MIN = x
 ElseIf y=z OR y=r Then
  MIN = y
 ElseIf z=r Then
  MIN = z
	End If
End Function

Dim UserName
UserName = Request.Cookies("NightFire")("UserName")

Dim RST
Set RST = Conn.Execute("SELECT UserName, ReasearchCenters FROM DONEBUILDINGS WHERE UserName ='"&UserName&"'")
Dim ResCen, TLand, ACapital, AElementX
ResCen = RST("ReasearchCenters")
Set RST = Conn.Execute("SELECT UserName, Land, Capital, ElementX FROM GENERALSTATS WHERE UserName = '"&UserName&"'")
TLand = RST("Land")
ACapital = RST("Capital")
AElementX = RST("ElementX")
Set RST = Nothing

Dim CreditCost
CreditCost = Int(1.465*TLand - 1.465*TLand*(ResCen*2/TLand))

'CREATE TABLE RESEARCH (
'UserName TEXT(20),
'Energy INTEGER,
'Military INTEGER,
'Stock INTEGER,
'Atomic INTEGER,
'Attack INTEGER,
'Stealth INTEGER,
'EnergyPts INTEGER,
'MilitaryPts INTEGER,
'StockPts INTEGER,
'AtomicPts INTEGER,
'AttackPts INTEGER,
'StealthPts INTEGER
')
'CREATE TABLE RESEARCHING (
'UserName TEXT(20),
'Energy INTEGER,
'Military INTEGER,
'Stock INTEGER,
'Atomic INTEGER,
'Attack INTEGER,
'Stealth INTEGER,
'TimeR TIME
')


Dim RS, EnergyL, MilitaryL, StockL, AtomicL, EnergyP, MilitaryP, StockP, AtomicP, EnergyPR, MilitaryPR, StockPR, AtomicPR
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open "SELECT * FROM RESEARCH WHERE UserName = '"&UserName&"'", Conn
EnergyL = RS("Energy")
MilitaryL = RS("Military")
StockL = RS("Stock")
AtomicL = RS("Atomic")
EnergyP = RS("EnergyPts")
MilitaryP = RS("MilitaryPts")
StockP = RS("StockPts")
AtomicP = RS("AtomicPts")
EnergyPR = 0
MilitaryPR = 0
StockPR = 0
AtomicPR = 0
RS.Close

RS.Open "SELECT * FROM RESEARCHING WHERE UserName = '"&UserName&"'", Conn
If RS.EOF = FALSE Then
	Do While Not RS.EOF
		EnergyPR = EnergyPR + RS("Energy")
		MilitaryPR = MilitaryPR + RS("Military")
		StockPR = StockPR + RS("Stock")
		AtomicPR = AtomicPR + RS("Atomic")
		RS.MoveNext
	Loop
End If
RS.Close

Dim Research, CostE, CostM, CostS, CostA, CostAT, CostST, Message
	
Dim NLvlE, NLvlM, NLvlS, NLvlA, LvlE, LvlM, LvlS, LvlA
Select Case EnergyL
	Case 0
		CostE = 500
		LvlE = "Currently our nation is able to collect energy through Solar Panels and Atomic Plants.<br>Both Energy Plants are running at 100% efficiency"
		NLvlE = "Our scientists believe if they are given 500 credits, they will be able to boost energy plant output by 10%"
	Case 1
		CostE = 1000
		LvlE = "Currently our nation is able to collect energy through Solar Panels and Atomic Plants.<br>Energy Production is running at 110%"
		NLvlE = "Our scientists believe if they are given 1000, they will be able to boost energy output by an additional 10%"
	Case 2
		CostE = 1500
		lvlE = "Our power production is done through Solar Panels and Atomic plants at the moment. Our energy production is running at 120% through researching"
		NLvlE = "Our scientists are very close to discovering the secret of antimatter energy production.  If an additional 1500 credits were put into researching, we will have antimatter production within our grasps!"
	Case 3
		lvlE = "Our nation is able to collect energy through Solar Panels, Atomic plants, and Antimatter plants.<br>Through vigorous researching, our power plants are running at 120%"
		NLvlE = "Our scientists have no further researching that can be attained."
End Select

Select Case MilitaryL
	Case 0
		CostM = 550
		LvlM = "Currently our military strategists use strategies that are equal to the defenses of standard armies."
		NLvlM = "Military intellegence believes that if they are allowed 550 credits, they will be able to increase offensive military strategies by 10%"
	Case 1
		CostM = 1700
		LvlM = "Our military strategists are using strategies that are able to boost our offensive forces strength by 10% right now."
		NLvlM = "Military intellegence officers are certain that if they are given 1700 credits, they will be able to increase our defensive powers by an additional 10%"
	Case 2
		LvlM = "Our military strategists are using strategies that boost our offense and defense by 10%"
		NLvlM = "Our military strategists are unable to further the development of strategies."
End Select

Select Case StockL
	Case 0
		CostS = 600
		LvlS = "Our GDP is steadily rising and the stock market is generaly increasing, although, no sharp upward trend is expected."
		NLvlS = "Our economists believe with the proper funding of 600 credits, they will be able to increase the rise of GDP. The stock market will flourish and civillians will make approxomately 10% more than current situations."
	Case 1
		LvlS = "Our civilians are making 10% more than similar nation's civilians. The GDP is rising faster than expected and many new companies are flourishing in the stock market."
		NLvlS = "Our economists are not able to aid the economy at the moment."
End Select

Select Case AtomicL
	Case 0
		CostA = 800
		LvlA = "Our physicists are experimenting with the power of splitting the atom. No breakthroughs have occured yet however."
		NLvlA = "Our physicists believe with 800 credits they will be able to split the atom. However, there is a chance that you will loose research centers in the process"
	Case 1
		CostA = 2000
		LvlA = "Advancing the uses of atomic energy for ways of good is no longer an option.  The atomic bomb has been born and your nation has forever changed, you now have access to atomic bombs."
		NLvlA = "Our physicists believe they can extend the blast radius of the Atomic bomb by developing Hydrogen bombs. This will take 2000 credits. This will most definately result in the destruction of some cities."
	Case 2
		CostA = 5000
		LvlA = "With the advent of the Hydrogen bomb, our military officers are greedily awaiting as many of these weapons as they can get. "
		NLvlA = "Nuclear fission is on the brink of discovery. However, our physicists will require 5000 credits. It is predicted that we will loose 100 research centers in the process."
	Case 3
		LvlA = "Nuculer fission bombs are now in position. There is no limit to the destructiveness of our nation."
		NLvlA = "No more research can be done without considerable damage of the nation."
End Select

Research = Request.Form("Research")
If Research <> "" Then
	Set RST = Server.CreateObject("ADODB.Recordset")
	
	RST.Open "SELECT Land, Capital, UserName, ElementX, Energy FROM GENERALSTATS WHERE UserName = '"&Username&"'", Conn, adOpenStatic, adLockPessimistic
	
	Dim Cost, CostC, CostX, REnergy, RMilitary, RStock, RAtomic, RTotal
	REnergy = Request.Form("Energy")
	RMilitary = Request.Form("Military")
	RStock = Request.Form("Stock")
	RAtomic = Request.Form("Atomic")
	RTotal = REnergy + 0 + RMilitary + 0 + RStock + 0 + RAtomic + 0
	CostC = CreditCost*RTotal
	CostX = Int(CreditCost/10)*RTotal

	If CostC > RST("Capital") OR CostX > RST("ElementX") Then
		Message = "<p class = 'normal'>You don't have enough resources to invest that many credits</p>"
	Else
		Dim RST2
		Set RST2 = Server.CreateObject("ADODB.Recordset")
		Message = "<p class = 'normal'>Thank you for contributing to the furtherment of science and technology in your nation. "&RTotal&" Credits have been invested in your Researching Centers, they should arrive in "&Int(5 - 5*ResCen*2/TLand*5)&" hours </p>"
		RST2.Open "RESEARCHING", Conn, adOpenStatic, adLockPessimistic
		RST2.AddNew
		RST2("UserName") = UserName
		RST2("TimeR") = DateAdd("h", 5 - Int(10*ResCen/TLand), Now())
		RST2("Energy") = REnergy
		RST2("Military") = RMilitary
		RST2("Stock") = RStock
		RST2("Atomic") = RAtomic
		RST2("Stealth") = 0
		RST2("Attack") = 0
		RST2.Update
		RST("Capital") = RST("Capital") - CostC
		RST("ElementX") = RST("ElementX") - CostX
		RST.Update
		RST2.Close
	End If
End If
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

<body bgcolor="#000000" text="#ffffff" linkr="#ffffff" vlink="#ffffff" alink="ffcc33" >
<SCRIPT SRC="fieldcheck.js"></SCRIPT>
<!--#Include File = "ad.asp"-->
<!--#Include File = "statbar.asp"-->
<br>
<br>
<center>
<p class = "large">Researching Facilities<br>
<span class = "small">You currently have <%=ResCen%> research centers operational, 1 credit will cost you <%=CreditCost%>c,  <%=Int(CreditCost/10)%>t, and will take <%= 5 - Int(10*ResCen/TLand)%> hours to commence researching. You can afford a maximum of <%=MIN(Int(ACapital/CreditCost), Int(AElementX/(CreditCost/10)), 999999999, 999999999)%> credits.</span></p>
<%=Message%>
<form action = "research.asp" method = "post">
<table width="100%">
<tr bgcolor = "#0000CC">
<td align = "center">Research Project</td>
<td align = "center">Current Level</td>
<td align = "center">Next Level</td>
<td align = "center">Credits</td>
<td align = "center">Research</td>
</tr>
<tr bgcolor = "#000066">
<td>Energy Production - <%=LvlE%></td>
<td><%=EnergyL%></td>
<td><%=NLvlE%></td>
<td><%=EnergyP%> <span class = "small">(<%=EnergyPR%>)</span></td>
<td><input type="text" size=6 value="0" name="Energy" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td>
</tr>
<tr bgcolor = "#000033">
<td>Military Strategy - <%=LvlM%></td>
<td><%=MilitaryL%></td>
<td><%=NLvlM%></td>
<td><%=MilitaryP%>  <span class = "small">(<%=MilitaryPR%>)</span></td>
<td><input type="text" size=6 value="0" name="Military" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td>
</tr>
<tr bgcolor = "#000066">
<td>Stock Market Development - <%=LvlS%></td>
<td><%=StockL%></td>
<td><%=NLvlS%></td>
<td><%=StockP%>  <span class = "small">(<%=StockPR%>)</span></td>
<td><input type="text" size=6 value="0" name="Stock" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td>
</tr>
<tr bgcolor = "#000033">
<td>Atomic Researching - <%=LvlA%></td>
<td><%=AtomicL%></td>
<td><%=NLvlA%></td>
<td><%=AtomicP%>  <span class = "small">(<%=AtomicPR%>)</span></td>
<td><input type="text" size=6 value="0" name="Atomic" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></td>
</tr>
</table>
<input type=hidden value = "BooYeah" name = "Research">
<input type=SUBMIT value="Invest Credits" class = "button" onMouseOver="ButtonLight(this);" onMouseOut="ButtonDark(this);">
</form>
</center>
