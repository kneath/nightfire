<%@ Language=VBScript %>
<% Option Explicit %>

<!--#include file="include.asp"-->

<%
Response.Buffer = True

Dim UserNameLP
UserNameLP = Request.Cookies("NightFire")("UserName")
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

<!--#Include File = "statbar.asp"-->
<%
Dim RSLP
Set RSLP = Server.CreateObject("ADODB.Recordset")
RSLP.Open "SELECT * FROM BUILDING WHERE UserName = '"&UserNameLP&"'", Conn, adOpenForwardOnly

Dim iIndex, BuildLand, arrBarracks(12), arrFactories(12), arrSolarPanels(12), arrConsulates(12), arrResearchCenters(15), arrElementXMines(12), arrDefenseStations(12), arrAtomicPlants(12), arrAntimatterPlants(12), arrPlutoniumMines(12), arrAntimatterLabs(12), arrNuclearSilos(23), arrTemples(12), arrCities(12)

BuildLand = 0
If RSLP.EOF = FALSE Then
	Do While RSLP.EOF = FALSE
		iIndex = DateDiff("h", Now, RSLP("Time")) - 1
		arrBarracks(iIndex) = arrBarracks(iIndex) + RSLP("Barracks")
		arrFactories(iIndex) = arrFactories(iIndex) + RSLP("Factories")
		arrSolarPanels(iIndex) = arrSolarPanels(iIndex) + RSLP("SolarPanels")
		arrConsulates(iIndex) = arrConsulates(iIndex) + RSLP("Consulates")	
		arrResearchCenters(iIndex) = arrResearchCenters(iIndex) + RSLP("ReasearchCenters")
		arrElementXMines(iIndex) = arrElementXMines(iIndex) + RSLP("ElementXMines")
		arrDefenseStations(iIndex) = arrDefenseStations(iIndex) + RSLP("DefenseStations")
		arrAtomicPlants(iIndex) = arrAtomicPlants(iIndex) + RSLP("AtomicPlants")
		arrAntimatterPlants(iIndex) = arrAntimatterPlants(iIndex) + RSLP("AntimatterPlants")
		arrPlutoniumMines(iIndex) = arrPlutoniumMines(iIndex) + RSLP("PlutoniumPlants")
		arrAntimatterLabs(iIndex) = arrAntimatterLabs(iIndex) + RSLP("AntimatterLabs")
		arrNuclearSilos(iIndex) = arrNuclearSilos(iIndex) + RSLP("NuclearSilos")
		arrTemples(iIndex) = arrTemples(iIndex) + RSLP("Temples")
		arrCities(iIndex) = arrCities(iIndex) + RSLP("Cities")
		BuildLand = BuildLand + RSLP("PlainLand")
		RSLP.MoveNext
	Loop
RSLP.MoveFirst
End If

Dim TBarracks, TFactories, TSolarPanels, TConsulates, TResearchCenters, TElementXMines, TDefenseStations, TAtomicPlants, TAntimatterPlants, TPlutoniumMines, TAntimatterLabs, TNuclearSilos, TTemples, TCities

TBarracks = 0
TFactories = 0
TSolarPanels = 0
TConsulates = 0
TResearchCenters = 0
TElementXMines = 0
TDefenseStations = 0
TAtomicPlants = 0
TAntimatterPlants = 0
TPlutoniumMines = 0
TAntimatterLabs = 0
TNuclearSilos = 0
TTemples = 0
TCities = 0

Dim iIndex2, iIndex3
iIndex = 0
iIndex2 = 0
iIndex3 = 0

For iIndex = 0 to 11
	If arrBarracks(iIndex) = "" OR arrBarracks(iIndex) = "0" Then
		arrBarracks(iIndex) = "-"
	Else
		TBarracks = TBarracks + arrBarracks(iIndex)
	End If
	If arrFactories(iIndex) = "" OR arrFactories(iIndex) = "0" Then
		arrFactories(iIndex) = "-"
	Else
		TFactories = TFactories + arrFactories(iIndex)

	End If
	If arrSolarPanels(iIndex) = "" OR arrSolarPanels(iIndex) = "0" Then
		arrSolarPanels(iIndex) = "-"
	Else
		TSolarPanels = TSolarPanels + arrSolarPanels(iIndex)
	End If
	If arrConsulates(iIndex) = "" OR arrConsulates(iIndex) = "0" Then
		arrConsulates(iIndex) = "-"
	Else
		TConsulates = TConsulates + arrConsulates(iIndex)
	End If
	If arrElementXMines(iIndex) = "" OR arrElementXMines(iIndex) = "0" Then
		arrElementXMines(iIndex) = "-"
	Else
		TElementXMines = TElementXMines + arrElementXMines(iIndex)
	End If
	If arrDefenseStations(iIndex) = "" OR arrDefenseStations(iIndex) = "0" Then
		arrDefenseStations(iIndex) = "-"
	Else
		TDefenseStations = TDefenseStations + arrDefenseStations(iIndex)
	End If
	If arrAtomicPlants(iIndex) = "" OR arrAtomicPlants(iIndex) = "0" Then
		arrAtomicPlants(iIndex) = "-"
	Else
		TAtomicPlants = TAtomicPlants + arrAtomicPlants(iIndex)
	End If
	If arrAntimatterPlants(iIndex) = "" OR arrAntimatterPlants(iIndex) = "0" Then
		arrAntimatterPlants(iIndex) = "-"
	Else
		TAntimatterPlants = TAntimatterPlants + arrAntimatterPlants(iIndex)
	End If
	If arrPlutoniumMines(iIndex) = "" OR arrPlutoniumMines(iIndex) = "0" Then
		arrPlutoniumMines(iIndex) = "-"
	Else
		TPlutoniumMines = TPlutoniumMines + arrPlutoniumMines(iIndex)
	End If
	If arrAntimatterLabs(iIndex) = "" OR arrAntimatterLabs(iIndex) = "0" Then
		arrAntimatterLabs(iIndex) = "-"
	Else
		TAntimatterLabs = TAntimatterLabs + arrAntimatterLabs(iIndex)
	End If
	If arrTemples(iIndex) = "" OR arrTemples(iIndex) = "0" Then
		arrTemples(iIndex) = "-"
	Else
		TTemples = TTemples + arrTemples(iIndex)
	End If
	If arrCities(iIndex) = "" OR arrCities(iIndex) = "0" Then
		arrCities(iIndex) = "-"
	Else
		TCities = TCities + arrCities(iIndex)
	End If
Next

For iIndex2 = 0 to 15
	If arrResearchCenters(iIndex2) = "" OR arrResearchCenters(iIndex2) = "0" Then
		arrResearchCenters(iIndex2) = "-"
	End If
Next

For iIndex3 = 0 to 23
	If arrNuclearSilos(iIndex3) = "" OR arrNuclearSilos(iIndex3) = "0" Then
		arrNuclearSilos(iIndex3) = "-"
	End If
Next

RSLP.Close
RSLP.Open "SELECT * FROM DONEBUILDINGS WHERE UserName = '"&UserNameLP&"'", Conn, adOpenForwardOnly
%>

</tr></table><br>


<table cellspacing=0 cellpadding=5>

<tr><td bgcolor="#ee5900" colspan=26><div align=center class="large">Land - Plain - Profile</div></td></tr>

<tr><td bgcolor="#000066" colspan=26><div align=center>

</div></td></tr>



<tr bgcolor="#000044">

<td><div><b>Building<b></div></td>

<td><div class = "small">1</div></td>

<td><div class = "small">2</div></td>

<td><div class = "small">3</div></td>

<td><div class = "small">4</div></td>

<td><div class = "small">5</div></td>

<td><div class = "small">6</div></td>

<td><div class = "small">7</div></td>

<td><div class = "small">8</div></td>

<td><div class = "small">9</div></td>

<td><div class = "small">10</div></td>

<td><div class = "small">11</div></td>

<td><div class = "small">12</div></td>

<td><div class = "small">13</div></td>

<td><div class = "small">14</div></td>

<td><div class = "small">15</div></td>

<td><div class = "small">16</div></td>

<td><div class = "small">17</div></td>

<td><div class = "small">18</div></td>

<td><div class = "small">19</div></td>

<td><div class = "small">20</div></td>

<td><div class = "small">21</div></td>

<td><div class = "small">22</div></td>

<td><div class = "small">23</div></td>

<td><div class = "small">24</div></td>

<td><div class = "small"><b>Total</b></div></td>

</tr>



<tr bgcolor="#000066">

<td><div><b>Barracks<b></div></td>

<td><div class = "small"><%=arrBarracks(0)%></div></td>

<td><div class = "small"><%=arrBarracks(1)%></div></td>

<td><div class = "small"><%=arrBarracks(2)%></div></td>

<td><div class = "small"><%=arrBarracks(3)%></div></td>

<td><div class = "small"><%=arrBarracks(4)%></div></td>

<td><div class = "small"><%=arrBarracks(5)%></div></td>

<td><div class = "small"><%=arrBarracks(6)%></div></td>

<td><div class = "small"><%=arrBarracks(7)%></div></td>

<td><div class = "small"><%=arrBarracks(8)%></div></td>

<td><div class = "small"><%=arrBarracks(9)%></div></td>

<td><div class = "small"><%=arrBarracks(10)%></div></td>

<td><div class = "small"><%=arrBarracks(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-<div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TBarracks, 0)%>) <%=FormatNumber(RSLP("Barracks"), 0)%></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><b>Factories<b></div></td>

<td><div class = "small"><%=arrFactories(0)%></div></td>

<td><div class = "small"><%=arrFactories(1)%></div></td>

<td><div class = "small"><%=arrFactories(2)%></div></td>

<td><div class = "small"><%=arrFactories(3)%></div></td>

<td><div class = "small"><%=arrFactories(4)%></div></td>

<td><div class = "small"><%=arrFactories(5)%></div></td>

<td><div class = "small"><%=arrFactories(6)%></div></td>

<td><div class = "small"><%=arrFactories(7)%></div></td>

<td><div class = "small"><%=arrFactories(8)%></div></td>

<td><div class = "small"><%=arrFactories(9)%></div></td>

<td><div class = "small"><%=arrFactories(10)%></div></td>

<td><div class = "small"><%=arrFactories(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TFactories, 0)%>)<%=FormatNumber(RSLP("Factories"), 0)%></div></td>

</tr>



<tr bgcolor="#000066">

<td><div><b>Solar Panels<b></div></td>

<td><div class = "small"><%=arrSolarPanels(0)%></div></td>

<td><div class = "small"><%=arrSolarPanels(1)%></div></td>

<td><div class = "small"><%=arrSolarPanels(2)%></div></td>

<td><div class = "small"><%=arrSolarPanels(3)%></div></td>

<td><div class = "small"><%=arrSolarPanels(4)%></div></td>

<td><div class = "small"><%=arrSolarPanels(5)%></div></td>

<td><div class = "small"><%=arrSolarPanels(6)%></div></td>

<td><div class = "small"><%=arrSolarPanels(7)%></div></td>

<td><div class = "small"><%=arrSolarPanels(8)%></div></td>

<td><div class = "small"><%=arrSolarPanels(9)%></div></td>

<td><div class = "small"><%=arrSolarPanels(10)%></div></td>

<td><div class = "small"><%=arrSolarPanels(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TSolarPanels, 0)%>)<%=FormatNumber(RSLP("SolarPanels"), 0)%></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><b>Consulates<b></div></td>

<td><div class = "small"><%=arrConsulates(0)%></div></td>

<td><div class = "small"><%=arrConsulates(1)%></div></td>

<td><div class = "small"><%=arrConsulates(2)%></div></td>

<td><div class = "small"><%=arrConsulates(3)%></div></td>

<td><div class = "small"><%=arrConsulates(4)%></div></td>

<td><div class = "small"><%=arrConsulates(5)%></div></td>

<td><div class = "small"><%=arrConsulates(6)%></div></td>

<td><div class = "small"><%=arrConsulates(7)%></div></td>

<td><div class = "small"><%=arrConsulates(8)%></div></td>

<td><div class = "small"><%=arrConsulates(9)%></div></td>

<td><div class = "small"><%=arrConsulates(10)%></div></td>

<td><div class = "small"><%=arrConsulates(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TConsulates, 0)%>)<%=FormatNumber(RSLP("Consulates"), 0)%></div></td>

</tr>



<tr bgcolor="#000066">

<td><div><b>Research Centers<b></div></td>

<td><div class = "small"><%=arrResearchCenters(0)%></div></td>

<td><div class = "small"><%=arrResearchCenters(1)%></div></td>

<td><div class = "small"><%=arrResearchCenters(2)%></div></td>

<td><div class = "small"><%=arrResearchCenters(3)%></div></td>

<td><div class = "small"><%=arrResearchCenters(4)%></div></td>

<td><div class = "small"><%=arrResearchCenters(5)%></div></td>

<td><div class = "small"><%=arrResearchCenters(6)%></div></td>

<td><div class = "small"><%=arrResearchCenters(7)%></div></td>

<td><div class = "small"><%=arrResearchCenters(8)%></div></td>

<td><div class = "small"><%=arrResearchCenters(9)%></div></td>

<td><div class = "small"><%=arrResearchCenters(10)%></div></td>

<td><div class = "small"><%=arrResearchCenters(11)%></div></td>

<td><div class = "small"><%=arrResearchCenters(12)%></div></td>

<td><div class = "small"><%=arrResearchCenters(13)%></div></td>

<td><div class = "small"><%=arrResearchCenters(14)%></div></td>

<td><div class = "small"><%=arrResearchCenters(15)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TResearchCenters, 0)%>)<%=FormatNumber(RSLP("ReasearchCenters"), 0)%></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><b>Element-X Mines<b></div></td>

<td><div class = "small"><%=arrElementXMines(0)%></div></td>

<td><div class = "small"><%=arrElementXMines(1)%></div></td>

<td><div class = "small"><%=arrElementXMines(2)%></div></td>

<td><div class = "small"><%=arrElementXMines(3)%></div></td>

<td><div class = "small"><%=arrElementXMines(4)%></div></td>

<td><div class = "small"><%=arrElementXMines(5)%></div></td>

<td><div class = "small"><%=arrElementXMines(6)%></div></td>

<td><div class = "small"><%=arrElementXMines(7)%></div></td>

<td><div class = "small"><%=arrElementXMines(8)%></div></td>

<td><div class = "small"><%=arrElementXMines(9)%></div></td>

<td><div class = "small"><%=arrElementXMines(10)%></div></td>

<td><div class = "small"><%=arrElementXMines(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TElementXMines, 0)%>)<%=FormatNumber(RSLP("ElementXMines"), 0)%></div></td>

</tr>



<tr bgcolor="#000066">

<td><div><b>Cities<b></div></td>

<td><div class = "small"><%=arrCities(0)%></div></td>

<td><div class = "small"><%=arrCities(1)%></div></td>

<td><div class = "small"><%=arrCities(2)%></div></td>

<td><div class = "small"><%=arrCities(3)%></div></td>

<td><div class = "small"><%=arrCities(4)%></div></td>

<td><div class = "small"><%=arrCities(5)%></div></td>

<td><div class = "small"><%=arrCities(6)%></div></td>

<td><div class = "small"><%=arrCities(7)%></div></td>

<td><div class = "small"><%=arrCities(8)%></div></td>

<td><div class = "small"><%=arrCities(9)%></div></td>

<td><div class = "small"><%=arrCities(10)%></div></td>

<td><div class = "small"><%=arrCities(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TCities, 0)%>)<%=FormatNumber(RSLP("Cities"), 0)%></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><b>Defense Stations<b></div></td>

<td><div class = "small"><%=arrDefenseStations(0)%></div></td>

<td><div class = "small"><%=arrDefenseStations(1)%></div></td>

<td><div class = "small"><%=arrDefenseStations(2)%></div></td>

<td><div class = "small"><%=arrDefenseStations(3)%></div></td>

<td><div class = "small"><%=arrDefenseStations(4)%></div></td>

<td><div class = "small"><%=arrDefenseStations(5)%></div></td>

<td><div class = "small"><%=arrDefenseStations(6)%></div></td>

<td><div class = "small"><%=arrDefenseStations(7)%></div></td>

<td><div class = "small"><%=arrDefenseStations(8)%></div></td>

<td><div class = "small"><%=arrDefenseStations(9)%></div></td>

<td><div class = "small"><%=arrDefenseStations(10)%></div></td>

<td><div class = "small"><%=arrDefenseStations(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TDefenseStations, 0)%>)<%=FormatNumber(RSLP("DefenseStations"), 0)%></div></td>

</tr>



<tr bgcolor="#000066">

<td><div><b>Atomic Plants<b></div></td>

<td><div class = "small"><%=arrAtomicPlants(0)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(1)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(2)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(3)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(4)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(5)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(6)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(7)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(8)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(9)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(10)%></div></td>

<td><div class = "small"><%=arrAtomicPlants(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TAtomicPlants, 0)%>)<%=FormatNumber(RSLP("AtomicPlants"), 0)%></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><b>Antimatter Plants<b></div></td>

<td><div class = "small"><%=arrAntimatterPlants(0)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(1)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(2)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(3)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(4)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(5)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(6)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(7)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(8)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(9)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(10)%></div></td>

<td><div class = "small"><%=arrAntimatterPlants(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TAntimatterLabs, 0)%>)<%=FormatNumber(RSLP("AntimatterPlants"), 0)%></div></td>

</tr>



<tr bgcolor="#000066">

<td><div><b>Plutonium Mines<b></div></td>

<td><div class = "small"><%=arrPlutoniumMines(0)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(1)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(2)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(3)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(4)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(5)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(6)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(7)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(8)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(9)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(10)%></div></td>

<td><div class = "small"><%=arrPlutoniumMines(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TPlutoniumMines, 0)%>)<%=FormatNumber(RSLP("PlutoniumPlants"), 0)%></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><b>Antimatter Labs<b></div></td>

<td><div class = "small"><%=arrAntimatterLabs(0)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(1)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(2)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(3)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(4)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(5)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(6)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(7)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(8)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(9)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(10)%></div></td>

<td><div class = "small"><%=arrAntimatterLabs(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TAntimatterLabs, 0)%>)<%=FormatNumber(RSLP("AntimatterLabs"), 0)%></div></td>

</tr>



<tr bgcolor="#000066">

<td><div class = "small"><b>Nuclear Silos<b></div></td>

<td><div class = "small"><%=arrNuclearSilos(0)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(1)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(2)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(3)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(4)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(5)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(6)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(7)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(8)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(9)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(10)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(11)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(12)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(13)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(14)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(15)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(16)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(17)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(18)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(19)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(20)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(21)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(22)%></div></td>

<td><div class = "small"><%=arrNuclearSilos(23)%></div></td>

<td><div class = "small">(<%=FormatNumber(TNuclearSilos, 0)%>)<%=FormatNumber(RSLP("NuclearSilos"), 0)%></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><b>Temples<b></div></td>

<td><div class = "small"><%=arrTemples(0)%></div></td>

<td><div class = "small"><%=arrTemples(1)%></div></td>

<td><div class = "small"><%=arrTemples(2)%></div></td>

<td><div class = "small"><%=arrTemples(3)%></div></td>

<td><div class = "small"><%=arrTemples(4)%></div></td>

<td><div class = "small"><%=arrTemples(5)%></div></td>

<td><div class = "small"><%=arrTemples(6)%></div></td>

<td><div class = "small"><%=arrTemples(7)%></div></td>

<td><div class = "small"><%=arrTemples(8)%></div></td>

<td><div class = "small"><%=arrTemples(9)%></div></td>

<td><div class = "small"><%=arrTemples(10)%></div></td>

<td><div class = "small"><%=arrTemples(11)%></div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">-</div></td>

<td><div class = "small">(<%=FormatNumber(TTemples, 0)%>)<%=FormatNumber(RSLP("Temples"), 0)%></div></td>

</tr>

<%
RSLP.Close
RSLP.Open "SELECT * FROM LAND WHERE UserName = '"&UserNameLP&"'", Conn, adOpenForwardOnly

Dim BuiltLand, RSLPT

Set RSLPT = Server.CreateObject("ADODB.Recordset")
RSLPT.Open "SELECT Land, UserName FROM GENERALSTATS WHERE UserName = '"&UserNameLP&"'", Conn

BuiltLand = RSLPT("Land") - (RSLP("FreePlain") + 0 + RSLP("Smooth") + RSLP("Rugged"))

%>

<tr bgcolor="#000066">

<td colspan=26><div align=center>

Total Land: <%=FormatNumber(RSLPT("Land"), 0)%>&nbsp;&nbsp;Built Land: <%=FormatNumber(BuiltLand, 0)%>&nbsp;&nbsp;Incoming Land: <%=FormatNumber(BuildLand, 0)%>

</div></td></tr>
</table>
</body>
</html>
<%
RSLPT.Close
RSLP.Close
%>
