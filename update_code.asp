<%
'Make Connection

Call OpenConnection()

'Make Recordsets

Dim RSGIU, RSGSU, RSTU, RSDMU, RSAU, RSDBU, RSBU, RSLU, RSMG, RSRU

Set RSGIU = Server.CreateObject("ADODB.Recordset")
Set RSGSU = Server.CreateObject("ADODB.Recordset")
Set RSTU = Server.CreateObject("ADODB.Recordset")
Set RSDMU = Server.CreateObject("ADODB.Recordset")
Set RSAU = Server.CreateObject("ADODB.Recordset")
Set RSDBU = Server.CreateObject("ADODB.Recordset")
Set RSBU = Server.CreateObject("ADODB.Recordset")
Set RSLU = Server.CreateObject("ADODB.Recordset")
Set RSMG = Server.CreateObject("ADODB.Recordset")

RSGIU.Open "SELECT * FROM GeneralInfo WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic


'Figure out how long it has been

Dim LastU, DifferenceU, SystemNumber, Nation

Nation = RSGIU(1)
SystemNumber = RSGIU(5)
LastU = RSGIU(7)
DifferenceU = DateDiff("h", LastU, Now())\1
RSGIU(7) = Now()
Dim HealthU
MoraleU = RSGIU("Health")
If MoraleU < 65 Then
   MoraleU = MoraleU + 5
Else
	MoraleU = MoraleU + 3
End If
RSGIU.Update
RSGIU.Close

'Bring in the buildings!

RSDBU.Open "SELECT * FROM DoneBuildings WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic
RSBU.Open "SELECT * FROM Building WHERE UserName = '"&UserNameU&"'", Conn, adOpenDynamic, adLockPessimistic
RSLU.Open "SELECT * FROM Land WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic

'Bonus stuff
If RSLU("LandBonTime") = "" Then
	RSLU("LandBonTime") = Now()
	RSLU.Update
End If

Dim HasBuilding, BarracksU, FactoriesU, SolarPanelsU, ConsulatesU, ReasearchCentersU, ElementXMinesU, DefenseStationsU, AtomicPlantsU, AntimatterPlantsU, PlutoniumPlantsU, AntimatterLabsU, NuclearSilosU, TemplesU, CitiesU, PlainLandU, ElementXLandU, PlutoniumLandU, SmoothLandU, RuggedLandU, FreeLandU

If RSBU.EOF = FALSE Then
	Do While Not RSBU.EOF
		If DateDiff("h", RSBU(21), Now()) >= 0 Then
			BarracksU = BarracksU + 0 + RSBU(1) + 0
			FactoriesU = FactoriesU + 0 + RSBU(2) + 0
			SolarPanelsU = SolarPanelsU + 0 + RSBU(3) + 0
			ConsulatesU = ConsulatesU + 0 + RSBU(4) + 0
			ReasearchCentersU = ReasearchCentersU + 0 + RSBU(5) + 0
			ElementXMinesU = ElementXMinesU + 0 + RSBU(6) + 0
			DefenseStationsU = DefenseStationsU + 0 + RSBU(7) + 0
			AtomicPlantsU = AtomicPlantsU + 0 + RSBU(8) + 0
			AntimatterPlantsU = AntimatterPlantsU + 0 + RSBU(9) + 0
			PlutoniumPlantsU = PlutoniumPlantsU + 0 + RSBU(10) + 0
			AntimatterLabsU = AntimatterLabsU + 0 + RSBU(11) + 0
			NuclearSilosU = NuclearSilosU + 0 + RSBU(12) + 0
			TemplesU = TemplesU + 0 + RSBU(13) + 0
			CitiesU = CitiesU + 0 + RSBU(14) + 0
			PlainLandU = PlainLandU + 0 + RSBU(15) + 0
			ElementXLandU = ElementXLandU + 0 + RSBU(16) + 0
			PlutoniumLandU = PlutoniumLandU + 0 + RSBU(17) + 0
			SmoothLandU = SmoothLandU + 0 + RSBU(19) + 0
			RuggedLandU = RuggedLandU + 0 + RSBU(20) + 0
			RSBU.Delete
		Else
			HasBuilding = TRUE
		End If
		RSBU.MoveNext
	Loop
End If

'Figure out how many buildings for future uses
BarracksU = BarracksU + RSDBU(1) +0
FactoriesU = FactoriesU + RSDBU(2) +0
SolarPanelsU = SolarPanelsU + RSDBU(3) +0
ConsulatesU = ConsulatesU + RSDBU(4) +0
ReasearchCentersU = ReasearchCentersU + RSDBU(5) +0
ElementXMinesU = ElementXMinesU + RSDBU(6) +0
DefenseStationsU = DefenseStationsU + RSDBU(7) +0
AtomicPlantsU = AtomicPlantsU + RSDBU(8) +0
AntimatterPlantsU = AntimatterPlantsU + RSDBU(9) +0
PlutoniumPlantsU = PlutoniumPlantsU + RSDBU(10) +0
AntimatterLabsU = AntimatterLabsU + RSDBU(11) +0
NuclearSilosU = NuclearSilosU + RSDBU(12) +0
TemplesU = TemplesU + RSDBU(13) +0
CitiesU = CitiesU + RSDBU(14) +0
PlainLandU = PlainLandU + RSLU(1) +0
ElementXLandU = ElementXLandU + RSLU(2) +0
PlutoniumLandU = PlutoniumLandU + RSLU(3) +0
SmoothLandU = SmoothLandU + RSLU(5) +0
RuggedLandU = RuggedLandU + RSLU(6) +0

'This is used to fix the amount of free plain kms when buildings are being built

Dim TBarracksU, TFactoriesU, TSolarPanelsU, TConsulatesU, TReasearchCentersU, TElementXMinesU, TDefenseStationsU, TAtomicPlantsU, TAntimatterPlantsU, TPlutoniumPlantsU, TAntimatterLabsU, TNuclearSilosU, TTemplesU, TCitiesU, FreePlainU

TBarracksU = BarracksU + 0 
TFactoriesU = FactoriesU + 0 
TSolarPanelsU = SolarPanelsU + 0 
TConsulatesU = ConsulatesU + 0 
TReasearchCentersU = ReasearchCentersU + 0 
TElementXMinesU = ElementXMinesU + 0 
TDefenseStationsU = DefenseStationsU + 0 
TAtomicPlantsU = AtomicPlantsU + 0 
TAntimatterPlantsU = AntimatterPlantsU + 0 
TPlutoniumPlantsU = PlutoniumPlantsU + 0 
TAntimatterLabsU = AntimatterLabsU + 0 
TNuclearSilosU = NuclearSilosU + 0 
TTemplesU =TemplesU + 0 
TCitiesU = CitiesU + 0 
If HasBuilding = TRUE Then
	RSBU.MoveFirst
End If

If RSBU.EOF = FALSE Then
	RSBU.MoveFirst
	Do While Not RSBU.EOF
		TBarracksU = TBarracksU + 0 + RSBU(1) + 0
		TFactoriesU = TFactoriesU + 0 + RSBU(2) + 0
		TSolarPanelsU = TSolarPanelsU + 0 + RSBU(3) + 0
		TConsulatesU = TConsulatesU + 0 + RSBU(4) + 0
		TReasearchCentersU = TReasearchCentersU + 0 + RSBU(5) + 0
		TElementXMinesU = TElementXMinesU + 0 + RSBU(6) + 0
		TDefenseStationsU = TDefenseStationsU + 0 + RSBU(7) + 0
		TAtomicPlantsU = TAtomicPlantsU + 0 + RSBU(8) + 0
		TAntimatterPlantsU = TAntimatterPlantsU + 0 + RSBU(9) + 0
		TPlutoniumPlantsU = TPlutoniumPlantsU + 0 + RSBU(10) + 0
		TAntimatterLabsU = TAntimatterLabsU + 0 + RSBU(11) + 0
		TNuclearSilosU = TNuclearSilosU + 0 + RSBU(12) + 0
		TTemplesU = TTemplesU + 0 + RSBU(13) + 0
		TCitiesU = TCitiesU + 0 + RSBU(14) + 0
		RSBU.MoveNext
	Loop
End If


FreeLandU = PlainLandU - (TBarracksU + TFactoriesU + TSolarPanelsU + TConsulatesU + TReasearchCentersU + TElementXMinesU + TDefenseStationsU + TAtomicPlantsU + TAntimatterPlantsU + TNuclearSilosU + TTemplesU + TCitiesU + TPlutoniumPlantsU) +0


RSBU.Close

'Actually updating the buildings/land now

RSDBU(1) = BarracksU +0
RSDBU(2) = FactoriesU + 0
RSDBU(3) = SolarPanelsU + 0
RSDBU(4) = ConsulatesU + 0
RSDBU(5) = ReasearchCentersU + 0
RSDBU(6) = ElementXMinesU + 0
RSDBU(7) = DefenseStationsU + 0
RSDBU(8) = AtomicPlantsU + 0
RSDBU(9) = AntimatterPlantsU + 0
RSDBU(10) = PlutoniumPlantsU +0
RSDBU(11) = AntimatterLabsU + 0
RSDBU(12) = NuclearSilosU + 0
RSDBU(13) = TemplesU + 0
RSDBU(14) = CitiesU +0
RSDBU.Update
RSDBU.Close

RSLU(1) = PlainLandU + 0
RSLU(2) = ElementXLandU + 0
RSLU(3) = PlutoniumLandU + 0
RSLU(5) = SmoothLandU + 0
RSLU(6) = RuggedLandU + 0
RSLU(7) = FreeLandU + 0
RSLU.Update
RSLU.Close

'Open up military recordset

RSTU.Open "SELECT * FROM Training WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic
RSAU.Open "SELECT * FROM Attacking WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic
RSDMU.Open "SELECT * FROM DoneMilitary WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic

'update morale

Dim MoraleU
MoraleU = RSDMU("Morale")

MoraleU = MoraleU + 3*DifferenceU

If MoraleU > 100 Then
	MoraleU = 100
End If
'Bring in training units

Dim TBSlots, TFSlots, ProbesU, AgentsU, MilitaryTrain, MilitaryAttack, InfantryU, Elite1U, Elite2U, Elite3U, TransportsU, FactorySlotsU, UsedFactorySlotsU, BarracksSlotsU, UsedBarracksSlotsU
 
If RSTU.EOF = FALSE Then
	MilitaryTrain = "yes"
	Do While Not RSTU.EOF
			If DateDiff("h", RSTU(9), Now()) >= 0 Then
			MilitaryTrain = "yes"
			InfantryU = InfantryU + RSTU(1)
			Elite1U = Elite1U + RSTU(2)
			Elite2U = Elite2U + RSTU(3)
			Elite3U = Elite3U + RSTU(4)
			TransportsU = TransportsU + RSTU(5)
			ProbesU = ProbesU + RSTU(7)
			AgentsU = AgentsU + RSTU(8)
			RSTU.Delete
			Else
			TBSlots = TBSlots + RSTU(1) + RSTU(2)
			TFSlots = TFSlots + RSTU(3) + RSTU(4)*2 + RSTU(5)
			End If
		RSTU.MoveNext
	Loop
End If
RSTU.Close

'Bring in units that are atttacking
Dim Elite1AU, Elite2AU, Elite3AU, InfantryAU, TransportsAU, MagesAU, ProbesAU, AgentsAU

If RSAU.EOF = FALSE Then
	MilitaryAttack = "yes"
	Do While Not RSAU.EOF
		If DateDiff("h", RSAU(9), Now()) >= 0 Then
			InfantryU = InfantryU + RSAU(1)
			Elite1U = Elite1U + RSAU(2)
			Elite2U = Elite2U + RSAU(3)
			Elite3U = Elite3U + RSAU(4)
			TransportsU = TransportsU + RSAU(5)
			ProbesU = ProbesU + RSAU(7)
			AgentsU = AgentsU + RSAU(8)
			RSAU.Delete
		Else
			InfantryAU = InfantryAU + RSAU(1)
			Elite1AU = Elite1AU + RSAU(2)
			Elite2AU = Elite2AU + RSAU(3)
			Elite3AU = Elite3AU + RSAU(4)
			TransportsAU = TransportsAU + RSAU(5)
			ProbesAU = ProbesAU + RSAU(7)
			AgentsAU = AgentsAU + RSAU(8)
		End If
		RSAU.MoveNext
	Loop
End If
RSAU.Close

'Store amount of units in variables for further use 

InfantryU = InfantryU + RSDMU(1)
Elite1U = Elite1U + RSDMU(2)
Elite2U = Elite2U + RSDMU(3)
Elite3U = Elite3U + RSDMU(4)
ProbesU = ProbesU + RSDMU(7)
AgentsU = AgentsU + RSDMU(8)
TransportsU = TransportsU + RSDMU(5)

'Calculate factory/barracks slots availible

FactorySlotsU = FactoriesU * 80
BarracksSlotsU = BarracksU * 80

'Calculate used barrack/factory slots

UsedBarracksSlotsU = TBslots + 0
UsedFactorySlotsU = TFSlots + 0

'Update the database with all done military in it

RSDMU(1) = InfantryU
RSDMU(2) = Elite1U
RSDMU(3) = Elite2U
RSDMU(4) = Elite3U
RSDMU(5) = TransportsU
RSDMU("Morale") = MoraleU
RSDMU(7) = ProbesU
RSDMU(8) = AgentsU
RSDMU(11) = FactorySlotsU
RSDMU(13) = UsedFactorySlotsU
RSDMU(10) = BarracksSlotsU
RSDMU(12) = UsedBarracksSlotsU
RSDMU.Update
RSDMU.Close


Dim RSRRU, RlvlE, RlvlM, RlvlS, RlvlA, RptsE, RptsM, RptsS, RptsA, EnergyBonus, IncomeBonus
Set RSRRU = Server.CreateObject("ADODB.Recordset")
Set RSRU = Server.CreateObject("ADODB.Recordset")
RSRU.Open "SELECT * FROM RESEARCH WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic

RlvlE = RSRU("Energy")
RlvlM = RSRU("Military")
RlvlS = RSRU("Stock")
RlvlA = RSRU("Atomic")
RptsE = RSRU("EnergyPts")
RptsM = RSRU("MilitaryPts")
RptsS = RSRU("StockPts")
RptsA = RSRU("AtomicPts")

RSRRU.Open "SELECT * FROM RESEARCHING WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic

If RSRRU.EOF = FALSE Then
	Do While Not RSRRU.EOF
		If DateDiff("h", RSRRU("TimeR"), Now()) >= 0 Then

			RptsE = RptsE + RSRRU("Energy")
			RptsM = RptsM + RSRRU("Military")
			RptsS = RptsS + RSRRU("Stock")
			RptsA = RptsA + RSRRU("Atomic")
			RSRRU.Delete
			RSRRU.Update
		End If
		RSRRU.MoveNext
	Loop
	
	Select Case RlvlE
	Case 0
		If RptsE >= 500 Then
			RSRU("Energy") = RSRU("Energy") + 1
			RlvlE = RlvlE + 1
			RptsE = RptsE - 500
		End If
	Case 1
		If RptsE >= 1000 Then
			RSRU("Energy") = RSRU("Energy") + 1
			RlvlE = RlvlE + 1
			RptsE = RptsE - 500
		End If
	Case 2
		If RptsE >= 1500 Then
			RSRU("Energy") = RSRU("Energy") + 1
			RlvlE = RlvlE + 1
			RptsE = RptsE - 1500
		End If
	End Select
	Select Case RlvlM
	Case 0
		If RptsM >= 550 Then
			RSRU("Military") = RSRU("Military") + 1
				RlvlM = RlvlM + 1
			RptsM = RptsM - 550
		End If
	Case 1
		If RptsM >= 1700 Then
			RSRU("Military") = RSRU("Military") + 1
			RlvlM = RlvlM + 1
			RptsM = RptsM - 1700
		End If
	End Select
	Select Case RlvlS
	Case 0
		If RptsS >= 600 Then
			RSRU("Stock") = RSRU("Stock") + 1
			RlvlS = RlvlS + 1
			RptsS = RptsS - 600
		End If
	End Select
	Select Case RlvlA
	Case 0
		If RptsA >= 800 Then
			RSRU("Atomic") = RSRU("Atomic") + 1
			RlvlA = RlvlA + 1
			RptsA = RptsA - 800
		End If
	Case 1
		If RptsA >= 2000 Then
			RSRU("Atomic") = RSRU("Atomic") + 1
			RlvlA = RlvlA + 1
			RptsA = RptsA - 2000
		End If
	Case 2
		If RptsA >= 5000 Then
			RSRU("Atomic") = RSRU("Atomic") + 1
			RlvlA = RlvlA + 1
			RptsA = RptsA - 5000
		End If
	End Select		
	RSRU("EnergyPts") = RptsE
	RSRU("MilitaryPts") = RptsM
	RSRU("StockPts") = RptsS
	RSRU("AtomicPts") = RptsA
	RSRU.Update
End If

RSRRU.Close
RSRU.Close

Select Case RlvlE
	Case 1
		EnergyBonus = .1
	Case 2
		EnergyBonus = .2
	Case 3
		EnergyBonus = .2
End Select

Select Case RlvlS
	Case 1
		IncomeBonus = .1
End Select

'Start Adding New Civillians and the Such

Dim TotalLandU

TotalLandU = PlainLandU + ElementXLandU + PlutoniumLandU + SmoothLandU + RuggedLandU

Dim MaxPopU, CurrentPopU, NewPopU

RSGSU.Open "SELECT * FROM GeneralStats WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic

MaxPopU = CitiesU * 100
If MaxPopU < 0 Then
	MaxPopU = 0
End If

CurrentPopU = RSGSU(1)

Dim HoursLeft
HoursLeft = DifferenceU

Do While HoursLeft >= 0 
	If CurrentPopU < MaxPopU Then
		NewPopU = Int(((CurrentPopU * 1.03))/100)*100
		CurrentPopU = Int(((CurrentPopU * 1.03))/100)*100
	ElseIf CurrentPopU > MaxPopU Then
		NewPopU = CurrentPopU - Int(((CurrentPopU * .03))/100)*100
		CurrentPopU = CurrentPopU - Int(((CurrentPopU * .03))/100)*100
	ElseIf CurrentPopU = MaxPopU Then
		NewPopU = MaxPopU
	Else
		NewPopU = MaxPopU
	End If
	HoursLeft = HoursLeft - 1
Loop


NewPopU = MaxPopU
NewPopU = NewPopU -  InfantryU - Elite1U - Elite2U - Elite3U - TransportsU
Dim LandU, CapitalU, IncomeU, ModU, EnergyU, EnergyChangeU, EnergyModU, ManaU, ManaChangeU, AntimatterU, AntimatterChangeU, AntimatterModU, PlutoniumU, PlutoniumChangeU, PlutoniumModU, ElementXU, ElementXChangeU, ElementXModU


LandU = RSGSU(4)
CapitalU = RSGSU(6)
ModU = (ConsulatesU/TotalLandU )+ 1 + IncomeBonus
IncomeU = ((NewPopU) * ModU )* DifferenceU
CapitalU = CapitalU + IncomeU
CapitalU = Int(CapitalU)

AntimatterU = RSGSU(11)
AntimatterChangeU = (AntimatterLabsU) * DifferenceU*3
AntimatterModU = 0
AntimatterU = Int(AntimatterU + AntimatterChangeU + AntimatterModU)

PlutoniumU = RSGSU(10)
PlutoniumChangeU = (PlutoniumPlantsU * 10) * DifferenceU
PlutoniumU = Int(PlutoniumU + PlutoniumChangeU)

EnergyU = RSGSU(8) + 0
EnergyChangeU = (1 - (FactoriesU * 13) - (BarracksU * 10 ) - ((ConsulatesU + ReasearchCentersU + ElementXMinesU + DefenseStationsU + AntimatterLabsU + NuclearSilosU + TemplesU) * 2)) * DifferenceU
EnergyModU = (SolarPanelsU*60)*DifferenceU

If AtomicPlantsU*DifferenceU >= PlutoniumU Then
   EnergyModU = EnergyModU + PlutoniumU*180
   PlutoniumU = 0
Else
	EnergyModU = EnergyModU + (AtomicPlantsU)*180*DifferenceU
	PlutoniumU = PlutoniumU - AtomicPlantsU*DifferenceU
End If

If AntimatterPlantsU*DifferenceU <= AntimatterU Then
	EnergyModU = EnergyModU + (AntimatterPlantsU)*540*DifferenceU
	AntimatterU = AntimatterU - AntimatterPlantsU*DifferenceU
Else
	EnergyModU = EnergyModU + AntimatterU*180
	AntimatterU = 0
End If

EnergyU = Int((EnergyU + EnergyChangeU + EnergyModU)*(1 + EnergyBonus))

RSMG.Open "SELECT * FROM Messages WHERE Type = 'SystemNews" & SystemNumber & "' OR UserName = '" & UserNameU & "' AND Seen = 'Yes'", Conn, adOpenStatic, adLockPessimistic

If RSMG.EOF = FALSE Then
	Dim MessagesU
	MessagesU = "yes"
End if

Dim RSMGU, RSNU
Set RSNU = Server.CreateObject("ADODB.Recordset")
Set RSMGU = Server.CreateObject("ADODB.Recordset")
RSMGU.Open "MESSAGES", Conn, adOpenStatic, adLockPessimistic
RSNU.Open "Select * from NUKES WHERE UserName = '"&UserNameU&"'", Conn, adOpenStatic, adLockPessimistic

Dim AtomBombsU, HBombsU, NukeBombsU, i, x1
Randomize()

AtomBombsU = RSNU("AtomBombs")
HBombsU = RSNU("HBombs")
NukeBombsU = RSNU("NukeBombs")

For i = 0 to DifferenceU
Select Case RlvlA

Case 1
	If Rnd() < ((ReasearchCentersU/PlainLandU)/.2) Then
		AtomBombsU = AtomBombsU + 1
		RSMGU.AddNew
		RSMGU(0) = UserNameU
		RSMGU(4) = Now()
		RSMGU(1) = "<p>Our research centers have produced an <span class = 'alert'>Atomic Bomb!!</span></p>"
		RSMGU(3) = "No"
		RSMGU(2) = "MainAttack"
		RSMGU.Update
	End If
Case 2
	If Rnd() < ((ReasearchCentersU/PlainLandU)/.2) Then
		x1 = Rnd()
		If x1 < .5 Then
			AtomBombsU = AtomBombsU + 1
			RSMGU.AddNew
			RSMGU(0) = UserNameU
			RSMGU(4) = Now()
			RSMGU(1) = "<p>Our research centers have produced an <span class = 'alert'>Atomic Bomb!!</span></p>"
			RSMGU(3) = "No"
			RSMGU(2) = "MainAttack"
			RSMGU.Update
		Else
			HBombsU = HBombsU + 1
			RSMGU.AddNew
			RSMGU(0) = UserNameU
			RSMGU(4) = Now()
			RSMGU(1) = "<p>Our research centers have produced a <span class = 'alert'>Hydrogen Bomb!!</span></p>"
			RSMGU(3) = "No"
			RSMGU(2) = "MainAttack"
			RSMGU.Update
		End If
	End If
Case 3
	If Rnd() < ((ReasearchCentersU/PlainLandU)/.2) Then
		x1 = Rnd()
		If x1 < .3 Then
			AtomBombsU = AtomBombsU + 1
			RSMGU.AddNew
			RSMGU(0) = UserNameU
			RSMGU(4) = Now()
			RSMGU(1) = "<p>Our research centers have produced an <span class = 'alert'>Atomic Bomb!!</span></p>"
			RSMGU(3) = "No"
			RSMGU(2) = "MainAttack"
			RSMGU.Update
		Elseif x1 < .6 AND x > .3 Then
			HBombsU = HBombsU + 1
			RSMGU.AddNew
			RSMGU(0) = UserNameU
			RSMGU(4) = Now()
			RSMGU(1) = "<p>Our research centers have produced a <span class = 'alert'>Hydrogen Bomb!!</span></p>"
			RSMGU(3) = "No"
			RSMGU(2) = "MainAttack"
			RSMGU.Update
		Else
			NukeBombsU = NukeBombsU + 1
			RSMGU.AddNew
			RSMGU(0) = UserNameU
			RSMGU(4) = Now()
			RSMGU(1) = "<p>Our research centers have produced a <span class = 'alert'>Nuclear Bomb!!</span></p>"
			RSMGU(3) = "No"
			RSMGU(2) = "MainAttack"
			RSMGU.Update
		End If
	End If
End Select
Next

RSNU("AtomBombs") = AtomBombsU
RSNU("HBombs") = HBombsU
RSNU("NukeBombs") = NukeBombsU
RSNU.Update
RSNU.Close

If EnergyU < 0 Then
	EnergyU = 0
	RSMGU.AddNew
	RSMGU(0) = UserNameU
	RSMGU(4) = Now()
	RSMGU(1) = "<p>Your nation is in serious need of <span class = 'alert'>Energy!</span></p>"
	RSMGU(3) = "No"
	RSMGU(2) = "MainAttack"
	RSMGU.Update
	RSMGU.AddNew
	RSMGU(0) = UserNameU
	RSMGU(4) = Now()
	RSMGU(1) = "<p>"&Nation&" is in serious need of <span class = 'alert'>Energy!</span></p>"
	RSMGU(3) = "No"
	RSMGU(2) = "SystemNews"&SystemNumber
	RSMGU.Update
End If

ManaU = RSGSU(9)
ManaChangeU = (TemplesU * 10) * DifferenceU
ManaU = Int(ManaU + ManaChangeU)

ElementXU = RSGSU(7)
ElementXChangeU = ElementXMinesU * 50 * DifferenceU
ElementXU = Int(ElementXU + ElementXChangeU)

Dim NetworthU
NetworthU = (LandU * 8) + ((InfantryU + InfantryAU) * 2) + ((Elite1U + Elite1AU) * 5) + ((Elite2U + Elite2AU) * 15) + ((Elite3U + Elite3AU) * 25) + ((TransportsU + TransportsAU) * 5) + (BarracksU + FactoriesU + SolarPanelsU + ConsulatesU + ReasearchCentersU + ElementXMinesU + DefenseStationsU + AtomicPlantsU + AntimatterPlantsU + PlutoniumPlantsU + AntimatterLabsU + NuclearSilosU + TemplesU + CitiesU) * 15
NetworthU = Int(NetworthU)

'Update the database with general numbers in it

RSGSU(2) = NewPopU
RSGSU(3) = InfantryU + Elite1U + Elite2U + Elite3U + TransportsU + 0
RSGSU(1) = NewPopU +  InfantryU + Elite1U + Elite2U + Elite3U + TransportsU
RSGSU(4) = TotalLandU + 0
RSGSU(6) = CapitalU + 0 
RSGSU(7) = ElementXU + 0
RSGSU(8) = EnergyU + 0
RSGSU(9) = ManaU + 0
RSGSU(10) = PlutoniumU + 0
RSGSU(11) = AntimatterU + 0
RSGSU(5) = NetworthU
RSGSU.Update
RSGSU.Close

'Delete old messages 
If MessagesU = "yes" Then
   	RSMG.MoveFirst
	Do While Not RSMG.EOF
		If DateDiff("d", RSMG(4), Now()) > 1 Then
			RSMG.Delete
			RSMG.Update
		End If
		RSMG.MoveNext
	Loop
End If
RSMG.Close
%>
