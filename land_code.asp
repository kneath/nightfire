<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="include.asp"-->
<% Response.Buffer = True
Response.Expires = -1

Dim x, y, z, r, final

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

Dim MessageL





Dim UserNameL

UserNameL = Request.Cookies("NightFire")("UserName")

Dim ArrivedL

ArrivedL = Request.Form("Arrived")

Dim ConnL

Set ConnL = Server.CreateObject("ADODB.Connection")

ConnL.ConnectionString="Provider=SQLOLEDB;User ID=sa;Password=Kt6261;Initial Catalog=warpspire"

ConnL.Open 

Dim RSGSL, RSDBL, RSBL, RSLL, RSDML

Set RSGSL = Server.CreateObject("ADODB.Recordset")
Set RSDBL = Server.CreateObject("ADODB.Recordset")
Set RSBL = Server.CreateObject("ADODB.Recordset")
Set RSLL = Server.CreateObject("ADODB.Recordset")
Set RSDML = Server.CreateObject("ADODB.Recordset")

If Request("WhatDo") = "Raze" Then
	RSDBL.Open "SELECT * FROM DONEBUILDINGS WHERE UserName = '"&UserNameL&"'", ConnL, adOpenDynamic, adLockPessimistic
Else
	RSDBL.Open "SELECT * FROM DONEBUILDINGS WHERE UserName = '"&UserNameL&"'", ConnL, adOpenForwardOnly, adLockPessimistic
End If
If ArrivedL = "" Then
	RSBL.Open "SELECT * FROM BUILDING WHERE UserName = '"&UserNameL&"'", ConnL, adOpenForwardOnly, adLockPessimistic
	RSLL.Open "SELECT * FROM LAND WHERE UserName = '"&UserNameL&"'", ConnL, adOpenForwardOnly, adLockPessimistic
Else
	RSBL.Open "SELECT * FROM BUILDING WHERE UserName = '"&UserNameL&"'", ConnL, adOpenDynamic, adLockPessimistic
	RSLL.Open "SELECT * FROM LAND WHERE UserName = '"&UserNameL&"'", ConnL, adOpenDynamic, adLockPessimistic
End If

If RSLL.EOF = TRUE Then
Response.Redirect "login.html"
Else
Response.Cookies("NightFire").Expires = DateAdd("h", 2, Now())
End If

If ArrivedL = "" Then
	RSGSL.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&UserNameL&"'", ConnL, adOpenForwardOnly, adLockPessimistic
Else
	RSGSL.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&UserNameL&"'", ConnL, adOpenDynamic, adLockOptimistic
End If
Dim RSTT
Set RSTT = Conn.Execute("SELECT UserName, Energy FROM RESEARCH Where UserName = '"&UserNameL&"'")



Dim TLandL, ACapitalL, AElementXL, AAntimatterL, APlutoniumL, AInfantryL, PopulationL

ACapitalL = RSGSL("Capital")
AElementXL = RSGSL("ElementX")
AAntimatterL = RSGSL("Antimatter")
APlutoniumL = RSGSL("Plutonium")
TLandL = RSGSL("Land")
PopulationL = RSGSL("Population")

Dim TotLandL, BuildingLandL, ElementXLandL, SmoothLandL, RuggedLandL, PlutoniumLandL, FreePlainL

ElementXLandL = RSLL("ElementX")
SmoothLandL = RSLL("Smooth")
RuggedLandL = RSLL("Rugged")
PlutoniumLandL = RSLL("Plutonium")
FreePlainL = RSLL("FreePlain")

TotLandL = TLandL
TLandL = TLandL - ElementXLandL - SmoothLandL - RuggedLandL - PlutoniumLandL


If ArrivedL = "" Then
	RSDML.Open "SELECT UserName, Soldiers FROM DONEMILITARY WHERE UserName = '"&UserNameL&"'", ConnL, adOpenForwardOnly, adLockPessimistic

Else
	RSDML.Open "SELECT UserName, Soldiers, UsedBarracksSlots FROM DONEMILITARY WHERE UserName = '"&UserNameL&"'", ConnL, adOpenDynamic, adLockPessimistic

End If

AInfantryL = RSDML("Soldiers")


Dim RAcresL, MessageEL, CapitalCostEL, SoldierCostEL

RAcresL = Request("ExploreAcres")
CapitalCostEL = RAcresL * 5000
SoldierCostEL = RAcresL * 10

If RAcresL > 0 AND ArrivedL = "yes" Then
  If CapitalCostEL > ACapitalL Then
    MessageEL = "<p align = 'center'>You do not have enough capital to explore that many acres</p>"
  ElseIf SoldierCostEL > AInfantryL Then
    MessageEL = "<p align = 'center'>You do not have enough infantry to explore that many acres</p>"
  Else

  Dim LandType, Plainkm, ElementXkm, Plutoniumkm, Smoothkm, Ruggedkm
  LandType = Request.Form("LandType")
  Select Case LandType
  Case "Plain"
    Plainkm = RAcresL
    ElementXkm = 0
    Plutoniumkm = 0
    Smoothkm = 0
    Ruggedkm = 0
  Case "ElementX"
    Plainkm = 0
    ElementXkm = RAcresL
    Plutoniumkm = 0
    Smoothkm = 0
    Ruggedkm = 0
   Case "Plutonium"
    Plainkm = 0
    ElementXkm = 0
    Plutoniumkm = RAcresL
    Smoothkm = 0
    Ruggedkm = 0
   Case "Smooth"
    Plainkm = 0
    ElementXkm = 0
    Plutoniumkm = 0
    Smoothkm = RAcresL
    Ruggedkm = 0
   Case "Rugged"
    Plainkm = 0
    ElementXkm = 0
    Plutoniumkm = 0
    Smoothkm = 0
    Ruggedkm = RAcresL
  End Select
   

RSBL.AddNew
RSBL("UserName") = UserNameL
RSBL("Barracks") = 0
RSBL("Factories") = 0
RSBL("SolarPanels") = 0
RSBL("Consulates") = 0
RSBL("ReasearchCenters") = 0
RSBL("ElementXMines") = 0
RSBL("Cities") = RCitiesL
RSBL("DefenseStations") = 0
RSBL("AtomicPlants") = 0
RSBL("PlutoniumPlants") = 0
RSBL("AntimatterLabs") = 0
RSBL("NuclearSilos") = 0
RSBL("Temples")  = 0
RSBL("AntimatterPlants") = 0
RSBL("Cities") = 0
RSBL("PlainLand") = Plainkm
RSBL("ElementXLand") = ElementXkm
RSBL("PlutoniumLand") = Plutoniumkm
RSBL("AntimatterLand") = 0
RSBL("SmoothLand") = Smoothkm
RSBL("RuggedLand") = Ruggedkm
RSBL("Time") = DateAdd("h", 12, Now)
RSBL.Update
RSGSL.MoveFirst
RSGSL("Capital") = ACapitalL - CapitalCostEL
RSGSL.Update
RSDML.MoveFirst
RSDML("Soldiers") = AInfantryL - SoldierCostEL
RSDML.Update
MessageEL = "<p align = 'center' class = 'normal'>Exploration has begun for "&FormatNumber(RAcresL, 0)&" square kilometers of "&LandType&" land at a cost of "&FormatNumber(CapitalCostEL, 0)&"c and "&FormatNumber(SoldierCostEL, 0)&" Infantry</p>"
  End If
End If

ACapitalL = RSGSL("Capital")

If Request("WhatDo") = "Build" Then
 Dim RBarracksL, RFactoriesL, RSolarPanelsL, RConsulatesL, RReasearchCentersL, RElementXMinesL, RCitiesL, RDefenseStationsL, RAtomicPlantsL, RPlutoniumPlantsL, RAntimatterLabsL, RNuclearSilosL, RTemplesL, RAntimatterPlantsL
 RBarracksL = Request("Barracks")
 RFactoriesL = Request("Factories")
 RSolarPanelsL = Request("SolarPanels")
 RConsulatesL = Request("Consulates")
 RReasearchCentersL = Request("ResearchCenters")
 RElementXMinesL = Request("ElementXMines")
 RCitiesL = Request("Cities")
 RDefenseStationsL = Request("DefenseStations")
 RAtomicPlantsL = Request("AtomicPlants")
 RPlutoniumPlantsL = Request("PlutoniumMines")
 RNuclearSilosL = Request("NuclearSilos")
 RTemplesL = Request("Temples")
If RSTT("Energy") = 3 Then
 RAntimatterPlantsL = Request("AntimatterPlants")
 RAntimatterLabsL = Request("AntimatterLabs")
Else
 RAntimatterLabsL = 0 
 RAntimatterPlantsL = 0
End If

Dim CapitalCostL, ElementXCostL, AcreCostL

CapitalCostL = (RBarracksL * 1500) + 0 + (RFactoriesL * 2000) + 0 + (RSolarPanelsL * 300) + 0 + (RConsulatesL * 1000) + 0 + (RReasearchCentersL * 200000) + 0 + (RElementXMinesL * 1000) + 0 + (RCitiesL* 0) + 0 + (RDefenseStationsL * 1300) + 0 + (RAtomicPlantsL * 1200) + (RAntimatterPlantsL * 3000) + 0 + (RPlutoniumPlantsL * 1300) + 0 + (RAntimatterLabsL * 2000) + 0 + (RNuclearSilosL * 1200000) + 0 + (RTemplesL * 1000)
AcreCostL = RBarracksL + 0 + RFactoriesL + 0 + RSolarPanelsL + 0 + RConsulatesL + 0 + RReasearchCentersL + 0 + RElementXMinesL + 0 + RCitiesL + 0 + RDefenseStationsL + 0 + RAtomicPlantsL + 0 + RPlutoniumPlantsL + 0 + RAntimatterLabsL + 0 + RNuclearSilosL + 0 + RTemplesL + 0 + RAntimatterPlantsL
ElementXCostL = (RBarracksL * 50) + 0 + (RFactoriesL * 200) + 0 + (RSolarPanelsL * 20) + 0 + (RConsulatesL * 50) + 0 + (RReasearchCentersL * 100) + 0 + (RElementXMinesL * 100) + 0 + (RCitiesL* 0) + 0 + (RDefenseStationsL * 140) + 0 + (RAtomicPlantsL * 120) + (RAntimatterPlantsL * 3000) + 0 + (RPlutoniumPlantsL * 100) + 0 + (RAntimatterLabsL * 50) + 0 + (RNuclearSilosL * 1000) + 0 + (RTemplesL * 10)

  If CapitalCostL > ACapitalL Then
    MessageL = "<p align = 'center' class = 'normal'>You can not afford that many buildings</p>"
  ElseIf AcreCostL > FreePlainL Then
    MessageL = "<p align = 'center' class = 'normal'>You do not have enough free plain to construct that many buildings</p>"
  ElseIf ElementXCostL > AElementXL Then
    MessageL = "<p algin = 'center' class = 'normal'>You do not have enough Element X to start construction</p>"
  Else
RSBL.AddNew
RSBL("UserName") = UserNameL
RSBL("Barracks") = RBarracksL
RSBL("Factories") = RFactoriesL
RSBL("SolarPanels") = RSolarPanelsL
RSBL("Consulates") = RConsulatesL
RSBL("ReasearchCenters") = RReasearchCentersL
RSBL("ElementXMines") = RElementXMinesL
RSBL("Cities") = RCitiesL
RSBL("DefenseStations") = RDefenseStationsL
RSBL("AtomicPlants") = RAtomicPlantsL
RSBL("PlutoniumPlants") = RPlutoniumPlantsL
RSBL("AntimatterLabs") = RAntimatterLabsL
RSBL("NuclearSilos") = RNuclearSilosL
RSBL("Temples")  = RTemplesL
RSBL("AntimatterPlants") = RAntimatterPlantsL
RSBL("PlainLand") = 0
RSBL("ElementXLand") = 0
RSBL("PlutoniumLand") = 0
RSBL("AntimatterLand") = 0
RSBL("SmoothLand") = 0
RSBL("RuggedLand") = 0
RSBL("Time") = DateAdd("h", 12, Now)
RSBL.Update
RSGSL.MoveFirst
RSGSL("Capital") = ACapitalL - 0 -CapitalCostL
RSGSL("ElementX") = AElementXL - 0 - ElementXCostL
ACapitalL = ACapitalL - CapitalCostL
AElementXL = AElementXL - ElementXCostL
RSGSL.Update
RSLL("FreePlain") = FreePlainL - AcreCostL
RSLL.Update
If AcreCostL <> 0 Then
	MessageL = "<p align = 'center' class = 'normal'>Construction was started at a cost of  "&FormatNumber(CapitalCostL, 0)&"c and "&FormatNumber(ElementXCostL, 0)&"t and will take up "&FormatNumber(AcreCostL, 0)&" square kilometers."
End If
   End If
End If

	


Dim AllotCitiesL, DBarracksL, DFactoriesL, DSolarPanelsL, DConsulatesL, DReasearchCentersL, DElementXMinesL, DCitiesL, DDefenseStationsL, DAtomicPlantsL, DPlutoniumPlantsL, DAntimatterLabsL, DNuclearSilosL, DTemplesL, DAntimatterPlantsL, DPlutoniumMinesL

AllotCitiesL = PopulationL/100
DBarracksL = RSDBL("Barracks") + 0
DFactoriesL = RSDBL("Factories") + 0
DSolarPanelsL = RSDBL("SolarPanels") + 0
DConsulatesL = RSDBL("Consulates") + 0
DReasearchCentersL = RSDBL("ReasearchCenters") + 0
DElementXMinesL = RSDBL("ElementXMines") + 0
DCitiesL = RSDBL("Cities") + 0
DDefenseStationsL = RSDBL("DefenseStations") + 0
DAtomicPlantsL = RSDBL("AtomicPlants") + 0
DPlutoniumPlantsL = RSDBL("PlutoniumPlants") + 0
DAntimatterLabsL = RSDBL("AntimatterLabs") + 0
DNuclearSilosL = RSDBL("NuclearSilos") + 0
DTemplesL = RSDBL("Temples") + 0
DAntimatterPlantsL = RSDBL("AntimatterPlants") + 0

If Request("WhatDo") = "Raze" Then

	
 	RBarracksL = Request("Barracks") + 0
 	RFactoriesL = Request("Factories") + 0
 	RSolarPanelsL = Request("SolarPanels") + 0
 	RConsulatesL = Request("Consulates") + 0
 	RReasearchCentersL = Request("ResearchCenters") + 0
 	RElementXMinesL = Request("ElementXMines") + 0
 	RCitiesL = Request("Cities") + 0
 	RDefenseStationsL = Request("DefenseStations") + 0
 	RAtomicPlantsL = Request("AtomicPlants") + 0
 	RPlutoniumPlantsL = Request("PlutoniumMines") + 0
 	RAntimatterLabsL = Request("AntimatterLabs") + 0
 	RNuclearSilosL = Request("NuclearSilos") + 0
 	RTemplesL = Request("Temples") + 0
 	RAntimatterPlantsL = Request("AntimatterPlants") + 0

	If RBarracksL>DBarracksL OR RFactoriesL>DFactoriesL OR RSolarPanelsL>DSolarPanelsL OR RConsulatesL>DConsulatesL OR RReasearchCentersL>DReasearchCentersL OR RElementXMinesL>DElementXMinesL OR RCitiesL>DCitiesL OR RDefenseStationsL>DDefenseStationsL OR RAtomicPlantsL>DAtomicPlantsL OR RPlutoniumPlantsL>DPlutoniumPlantsL OR RAntimatterLabsL>DAntimatterLabsL OR RNuclearSilosL>DNuclearSilosL OR RTemplesL>DTemplesL OR RAntimatterPlantsL>DAntimatterPlantsL Then
		MessageL = "<p align = 'center' class = 'normal'>You do not have that many buildings to raze</p>"&" "&RBarracksL&">"&DBarracksL &" "& RFactoriesL&">"&DFactoriesL &" "& RSolarPanelsL&">"&DSolarPanelsL &" "& RConsulatesL&">"&DConsulatesL &" "& RReasearchCentersL&">"&DReasearchCentersL &" "&RElementXMinesL&">"&DElementXMinesL &" "& RCitiesL&">"&DCitiesL &" "& RDefenseStationsL&">"&DDefenseStationsL &" "&RAtomicPlantsL&">"&DAtomicPlantsL &" "& RPlutoniumPlantsL&">"&DPlutoniumPlantsL &" "& RAntimatterLabsL&">"&DAntimatterLabsL &" "& RNuclearSilosL&">"&DNuclearSilosL &" "& RTemplesL&">"&DTemplesL &" "& RAntimatterPlantsL&">"&DAntimatterPlantsL
	Else
		RSDBL("Barracks") = DBarracksL - RBarracksL
		RSDBL("Factories") = DFactoriesL - RFactoriesL
		RSDBL("SolarPanels") = DSolarPanelsL - RSolarPanelsL
		RSDBL("Consulates") = DConsulatesL - RConsulatesL
		RSDBL("ReasearchCenters") = DReasearchCentersL - RReasearchCentersL
		RSDBL("ElementXMines") = DElementXMinesL - RElementXMinesL
		RSDBL("Cities") = DCitiesL - RCitiesL
		RSDBL("DefenseStations") = DDefenseStationsL - RDefenseStationsL
		RSDBL("AtomicPlants") = DAtomicPlantsL - RAtomicPlantsL
		RSDBL("PlutoniumPlants") = DPlutoniumPlantsL - RPlutoniumPlantsL
		RSDBL("AntimatterPlants") = DAntimatterPlantsL - RAntimatterPlantsL
		RSDBL("AntimatterLabs") = DAntimatterLabsL - RAntimatterLabsL
		RSDBL("NuclearSilos") = DNuclearSilosL - RNuclearSilosL
		RSDBL("Temples") = DTemplesL - RTemplesL
		RSDBL.Update
 		RSLL("FreePlain") = RSLL("FreePlain") + (RBarracksL + 0 +  RFactoriesL + 0 + RSolarPanelsL + 0 +  RConsulatesL + 0 + RReasearchCentersL + 0 + RElementXMinesL + 0 + RCitiesL + 0 + RDefenseStationsL + 0 + RAtomicPlantsL + 0 + RPlutoniumPlantsL + 0 + RAntimatterLabsL + 0 + RNuclearSilosL + 0 + RTemplesL + 0 + RAntimatterPlantsL + 0)
		RSLL.Update
	MessageL = "<p align = 'center' class = 'normal'>"&RBarracksL + 0 +  RFactoriesL + 0 + RSolarPanelsL + 0 +  RConsulatesL + 0 + RReasearchCentersL + 0 + RElementXMinesL + 0 + RCitiesL + 0 + RDefenseStationsL + 0 + RAtomicPlantsL + 0 + RPlutoniumPlantsL + 0 + RAntimatterLabsL + 0 + RNuclearSilosL + 0 + RTemplesL + 0 + RAntimatterPlantsL + 0&" square km of buildings was destroyed.</p>"
	End If
	DBarracksL = RSDBL("Barracks")
	DFactoriesL = RSDBL("Factories")
	DSolarPanelsL = RSDBL("SolarPanels")
	DConsulatesL = RSDBL("Consulates")
	DReasearchCentersL = RSDBL("ReasearchCenters")
	DElementXMinesL = RSDBL("ElementXMines")
	DCitiesL = RSDBL("Cities")
	DDefenseStationsL = RSDBL("DefenseStations")
	DAtomicPlantsL = RSDBL("AtomicPlants")
	DPlutoniumPlantsL = RSDBL("PlutoniumPlants")
	DAntimatterLabsL = RSDBL("AntimatterLabs")
	DNuclearSilosL = RSDBL("NuclearSilos")
	DTemplesL = RSDBL("Temples")
	DAntimatterPlantsL = RSDBL("AntimatterPlants")

End If


Dim BBarracksL, BFactoriesL, BSolarPanelsL, BConsulatesL, BReasearchCentersL, BElementXMinesL, BCitiesL, BDefenseStationsL, BAtomicPlantsL, BPlutoniumPlantsL, BAntimatterLabsL, BNuclearSilosL, BTemplesL, BAntimatterPlantsL, BPlutoniumMinesL

If RSBL.EOF = FALSE Then
RSBL.MoveFirst
Do While RSBL.EOF = FALSE
	BBarracksL = BBarracksL + RSBL("Barracks")
	BFactoriesL = BFactoriesL + RSBL("Factories")
	BSolarPanelsL = BSolarPanelsL + RSBL("SolarPanels")
	BConsulatesL = BConsulatesL + RSBL("Consulates")
	BReasearchCentersL = BReasearchCentersL + RSBL("ReasearchCenters") 
	BElementXMinesL = BElementXMinesL + RSBL("ElementXMines")
	BCitiesL = BCitiesL + RSBL("Cities")
	BDefenseStationsL = BDefenseStationsL + RSBL("DefenseStations")
	BAtomicPlantsL = BAtomicPlantsL + RSBL("AtomicPlants")
	BPlutoniumPlantsL = BPlutoniumPlantsL + RSBL("PlutoniumPlants")
	BAntimatterLabsL = BAntimatterLabsL + RSBL("AntimatterLabs")
	BNuclearSilosL = BNuclearSilosL +RSBL("NuclearSilos")
	BTemplesL = BTemplesL + RSBL("Temples")
	BAntimatterPlantsL = BAntimatterPlantsL +RSBL("AntimatterPlants")
	BuildingLandL =  BuildingLandL +RSBL("PlainLand")
	RSBL.MoveNext
Loop
End If

AInfantryL = RSDML("Soldiers")
FreePlainL = RSLL("FreePlain")

Dim MBarracksL, MFactoriesL, MSolarPanelsL, MConsulatesL, MReasearchCentersL, MElementXMinesL, MCitiesL, MDefenseStationsL, MAtomicPlantsL, MPlutoniumPlantsL, MAntimatterLabsL, MNuclearSilosL, MTemplesL, MAntimatterPlantsL, MPlutoniumMinesL, TempCapitalL, TempElementXL, Num1, Num2, Num3, Num4

Num4 = 100000000000

Num1 = ACapitalL\1500
Num2 = AElementXL\150
Num3 = FreePlainL
MBarracksL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\2000
Num2 = AElementXL\200
Num3 = FreePlainL
MFactoriesL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\300
Num2 = AElementXL\20
Num3 = FreePlainL
MSolarPanelsL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\1000
Num2 = AElementXL\50
Num3 = FreePlainL
MConsulatesL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\20000
Num2 = AElementXL\100
Num3 = FreePlainL
MReasearchCentersL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\1000
Num2 = AElementXL\100
Num3 = FreePlainL
MElementXMinesL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\1300
Num2 = AElementXL\140
Num3 = FreePlainL
MDefenseStationsL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\1200
Num2 = AElementXL\120
Num3 = FreePlainL
MAtomicPlantsL = MIN(Num1, Num2, Num3, Num4)

If RSTT("Energy") = 3 Then
Num1 = 3000
Num2 = 100
Num3 = FreePlainL
MAntimatterPlantsL = MIN(Num1, Num2, Num3, Num4)

Num1 = 2000
Num2 = 50
Num3 = FreePlain
MAntimatterLabsL = MIN(Num1, Num2, Num3, Num4)
Else
MAntimatterPlantsL = 0
MAntimatterLabsL = 0
End If

Num1 = ACapitalL\1300
Num2 = AElementXL\100
Num3 = FreePlainL
MPlutoniumPlantsL = MIN(Num1, Num2, Num3, Num4)

Num1 = ACapitalL\1000
Num2 = AElementXL\10
Num3 = FreePlainL
MTemplesL = MIN(Num1, Num2, Num3, Num4)

MNuclearSilosL = 0

Dim MLandL
Num1 = ACapitalL\5000
Num2 = AInfantryL\10
Num3 = 100000000000000000
MLandL = MIN(Num1, Num2, Num3, Num4)


RSTT.Close
RSDBL.Close
RSBL.Close
RSGSl.Close
RSDML.Close
ConnL.Close

%>
