<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="include.asp"-->
<% Response.Buffer = True %>
<%
Dim Timer1, Timer2, TotalTime
Timer1 = Timer()

'Coded entirely by Duke Brak by hand
'This code may not be used in any other games than NightFire, a game by WarpSpire.
'www.warpspire.com/nightfire/
'Version 1.1 Alpha

Dim UserNameM
UserNameM = Request.Cookies("NightFire")("UserName")

'My personalized minimum function
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

Call OpenConnection()

Dim RSDMM, RSTM, RSGSM

Set RSDMM = Server.CreateObject("ADODB.Recordset")
Set RSTM = Server.CreateObject("ADODB.Recordset")
Set RSGSM = Server.CreateObject("ADODB.Recordset")

If Request("arrived") = "yes" Then
RSDMM.Open "SELECT * FROM DoneMilitary WHERE UserName = '"&UserNameM&"'", Conn, adOpenStatic, adLockOptimistic
Else
RSDMM.Open "SELECT * FROM DoneMilitary WHERE UserName = '"&UserNameM&"'", Conn, adOpenForwardOnly, adLockOptimistic
End If
RSTM.Open "SELECT * FROM Training WHERE UserName = '"&UserNameM&"'", Conn, adOpenStatic, adLockOptimistic


Dim TInfantryM, TElite1M, TElite2M, TElite3M, TTransportsM

	If RSTM.EOF = FALSE Then
 Do While RSTM.EOF = FALSE
		TInfantryM = TInfantryM + RSTM("Soldiers")
		TElite1M = TElite1M +RSTM("Elite1")
		TElite2M = TElite2M + RSTM("Elite2")
		TElite3M = TElite3M + RSTM("Elite3")
		TTransportsM = TTransportsM + RSTM("Transports")
		RSTM.MoveNext
 Loop
	Else
		TInfantryM = 0
		TElite1M = 0
		TElite2M = 0
		TElite3M = 0
		TTransportsM = 0
	End If
	

	RSGSM.Open "SELECT UserName, Civillian, Capital, ElementX FROM GeneralStats WHERE UserName = '"&UserNameM&"'", Conn, adOpenStatic, adLockPessimistic

	Dim AInfantryM, AElite1M, AElite2M, AElite3M, ATransportsM, TotBarSlots, AvBarSlots, TotFacSlots, AvFacSlots
	
	AInfantryM = RSDMM("Soldiers")
	AElite1M = RSDMM("Elite1")
	AElite2M = RSDMM("Elite2")
	AElite3M = RSDMM("Elite3")
	ATransportsM = RSDMM("Transports")

	

	Dim InfantryM, Elite1M, Elite2M, Elite3M, TransportsM
	
	InfantryM = Request.Form("Soldiers")
	Elite1M = Request.Form("Elite1")
	Elite2M = Request.Form("Elite2")
	Elite3M = Request.Form("Elite3")
	TransportsM = Request.Form("Transports")

Dim ErrorMessage

	If isNumeric(InfantryM) = FALSE or isNumeric(Elite1M) = FALSE or isNumeric(Elite2M) = FALSE or isNumeric(Elite3M) = FALSE or isNumeric(TransportsM) = FALSE Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry but the values you entered were not numbers, please try again</p>"
	ElseIf InfantryM < 0 or Elite1M < 0 or Elite2M < 0 or Elite3M < 0 or TransportsM < 0 Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry but the values you entered were not positive</p>"
	End If
	
	If Elite1M = "" Then
		Elite1M = 0
	ElseIf Elite2M = "" Then
		Elite2M = 0
	ElseIf Elite3M = "" Then
		Elite3M = 0
	ElseIf TransportsM = "" Then
		TransportsM = 0
	ElseIf InfantryM = "" Then
		InfantryM = 0
	End If
	

	
	Dim PCCivillianM, PCInfantryM, PCCapitalM, PCElementXM, PCBarracksM, PCFactoryM, PCUsedBSlots, PCUsedFSlots
	
	PCCivillianM = RSGSM("Civillian") + 0
	PCInfantryM = RSDMM("Soldiers")
	PCCapitalM = RSGSM("Capital")
	PCElementXM = RSGSM("ElementX")
	PCBarracksM = RSDMM("BarracksSlots")
	PCUsedBSlots = RSDMM("UsedBarracksSlots")
	PCBarracksM = PCBarracksM - PCUsedBSlots
	PCFactoryM = RSDMM("FactorySlots")
	PCUsedFSlots = RSDMM("UsedFactorySlots")
	PCFactoryM = PCFactoryM - PCUsedFSlots

 'If PCCivillianM =0 Then
 ' PCCivillianM = 1
 'ElseIf PCInfantryM = 0 Then
 ' PCInfantryM = 1
 'ElseIf PCCapitalM = 0 Then
 ' PCCapitalM = 1
 'ElseIf PCElementXM = 0 Then
 ' PCElementXM = 1
 'ElseIf PCBarracksM = 0 Then
 ' PCBarracksM = 1
 'ElseIf PCFactoryM = 0 Then
 ' PCFactoryM = 1
 'End If
 
 Dim RSRace

Set RSRace = Conn.Execute("SELECT Race, UserName FROM GENERALINFO WHERE UserName = '"&UserNameM&"'")

Dim RaceRace

RaceRace = RSRace("Race")
	
Dim RSStats, InfantryHours, Elite1Hours, Elite2Hours, Elite3Hours, TransportHours, InfantryCost, Elite1Cost, Elite2Cost, Elite3Cost, TransportCost, InfantryCostX, Elite1CostX, Elite2CostX, Elite3CostX, TransportCostX,InfantryStat, Elite1Stat, Elite2Stat, Elite3Stat, Elite1Name, Elite2Name, Elite3Name, InfantryDesc, Elite1Desc, Elite2Desc, Elite3Desc

Set RSStats = Server.CreateObject("ADODB.Recordset")
RSStats.Open "SELECT * FROM COSTS WHERE Type = 'Military' AND Race = '" & RaceRace & "'", Conn, adOpenStatic, adLockPessimistic

RSStats.Filter = "Unit = 'Infantry'"
InfantryStat = RSStats("CapitalCost") & "c | 1 Infantry | " & RSStats("Hours") & "Hours"
InfantryCost = RSStats("CapitalCost")
InfantryCostX = RSStats("ElementXCost")
InfantryHours = RSStats("Hours")
InfantryDesc = "1/1"
RSStats.Filter = adFilterNone
RSStats.Filter = "Unit = 'Elite1'"
Elite1Name = RSStats("Name")
Elite1Stat = RSStats("CapitalCost") & "c | " & RSStats("ElementXCost") & "t | " & RSStats("InfantryCost") & " infantry | " & RSStats("Hours") & " Hours"
Elite1Desc = RSStats("Offense") & "/" & RSStats("Defense") & " " & RSStats("Description")
Elite1Cost = RSStats("CapitalCost")
Elite1CostX = RSStats("ElementXCost")
Elite1Hours = RSStats("Hours")
RSStats.Filter = adFilterNone
RSStats.Filter = "Unit = 'Elite2'"
Elite2Stat = RSStats("CapitalCost") & "c | " & RSStats("ElementXCost") & "t | " & RSStats("InfantryCost") & " infantry | " & RSStats("Hours") & " Hours"
Elite2Desc = RSStats("Offense") & "/" & RSStats("Defense") & " " & RSStats("Description")
Elite2Name = RSStats("Name")
Elite2Cost = RSStats("CapitalCost")
Elite2CostX = RSStats("ElementXCost")
Elite2Hours = RSStats("Hours")
RSStats.Filter = adFilterNone
RSStats.Filter = "Unit = 'Elite3'"
Elite3Stat = RSStats("CapitalCost") & "c | " & RSStats("ElementXCost") & "t | " & RSStats("InfantryCost") & " infantry | " & RSStats("Hours") & " Hours"
Elite3Desc = RSStats("Offense") & "/" & RSStats("Defense") & " " & RSStats("Description")
Elite3name = RSStats("Name")
Elite3Cost = RSStats("CapitalCost")
Elite3CostX = RSStats("ElementXCost")
Elite3Hours = RSStats("Hours")
RSStats.Close

If Request("arrived") = "yes" Then
	Dim RCivillianM, RInfantryM, RCapitalM, RElementXM, RBarracksM, RFactoryM
	
	RCivillianM = InfantryM + 0
	RInfantryM = Elite1M + 0 + Elite2M + 0 + Elite3M + 0 + Elite3M + 0 + TransportsM + 0
	RCapitalM = (InfantryM * 50) + (Elite1M * Elite1Cost) + (Elite2M * Elite2Cost) + (Elite3M * Elite3Cost) + (TransportsM * 200)
	RElementXM = (Elite1M * Elite1CostX) + (Elite2M * Elite2CostX) + (Elite3M * Elite3CostX)
	RFactoryM = Elite2M + (Elite3M * 2) + TransportsM
	RBarracksM = InfantryM + 0 + Elite1M
	End If
	
	
If Request.Form("arrived") = "yes" AND ErrorMessage = "" Then
	


	If RCivillianM > PCCivillianM Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry, but you currently do not have enough Civilians to train that many infantry</p>"
	ElseIf RInfantryM>PCInfantryM Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry, but you currently do not have enough Infantry to train those troops.</p>"
	ElseIf RCapitalM>PCCapitalM Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry, but you currently do not have enough Capital to train those troops</p>"
	ElseIf RElementXM>PCElementXM Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry, but you don't have enough Element X to train those troops</p>"
	ElseIf RFactoryM>PCFactoryM Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry, but you don't have enough Factory slots to train those troops</p>"
	ElseIf RBarracksM>PCBarracksM Then
		ErrorMessage = "<p class = 'small' align = 'center'>I'm sorry but you don't have enough Barracks Slots to train those troops</p>"
	End If
	
	If ErrorMessage = "" Then
		RBarracksM = 0
		RBarracksM = InfantryM + 0 + Elite1M - 0 - Elite2M - 0 - (Elite3M*2) - 0 - TransportsM + 0
		TInfantryM = TInfantryM + InfantryM
		TElite1M = TElite1M + Elite1M
		TElite2M = TElite2M + Elite2M
		TElite3M = TElite3M + Elite3M
		TTransportsM = TTransportsM + TransportsM
		
		RSTM.AddNew
		RSTM("UserName") = UserNameM
		RSTM("Soldiers") = InfantryM
		RSTM("Elite1") = 0
		RSTM("Elite2") = 0
		RSTM("Elite3") = 0
		RSTM("Transports") = TransportsM
		RSTM("Probes") = 0
		RSTM("Agents") = 0
		RSTM("Mages") = 0
		RSTM("Time") = DateAdd("h", (InfantryHours-1), Now())
		RSTM.Update
		RSTM.AddNew
		RSTM("UserName") = UserNameM
		RSTM("Soldiers") = 0
		RSTM("Elite1") = Elite1M
		RSTM("Elite2") = 0
		RSTM("Elite3") = 0
		RSTM("Transports") = 0
		RSTM("Probes") = 0
		RSTM("Agents") = 0
		RSTM("Mages") = 0
		RSTM("Time") = DateAdd("h", (Elite1Hours-1), Now())
		RSTM.Update
		RSTM.AddNew
		RSTM("UserName") = UserNameM
		RSTM("Soldiers") = 0
		RSTM("Elite1") = 0
		RSTM("Elite2") = Elite2M
		RSTM("Elite3") = 0
		RSTM("Transports") = 0
		RSTM("Probes") = 0
		RSTM("Agents") = 0
		RSTM("Mages") = 0
		RSTM("Time") = DateAdd("h", (Elite2Hours-1), Now())
		RSTM.Update
		RSTM.AddNew
		RSTM("UserName") = UserNameM
		RSTM("Soldiers") = 0
		RSTM("Elite1") = 0
		RSTM("Elite2") = 0
		RSTM("Elite3") = Elite3M
		RSTM("Transports") = 0
		RSTM("Probes") = 0
		RSTM("Agents") = 0
		RSTM("Mages") = 0
		RSTM("Time") = DateAdd("h", (Elite3Hours-1), Now())
		RSGSM("Capital") = PCCapitalM - RCapitalM
		RSGSM("ElementX") = PCElementXM - RElementXM
		RSGSM("Civillian") = PCCivillianM - RCivillianM
		RSDMM("Soldiers") = PCInfantryM - RInfantryM
		RSDMM("UsedFactorySlots") = TElite2M + TElite3M*2 + TTransportsM
		RSDMM("UsedBarracksSlots") = TInfantryM + TElite1M 
		RSGSM.Update
		RSDMM.Update
		RSTM.Update
		
		PCCapitalM = PCCapitalM - RCapitalM
		PCElementXM = PCElementXM - RElementXM
		PCCivillianM = PCCivillianM - RCivillianM
		PCInfantryM = PCInfantryM - RInfantryM
		PCUsedFSlots = PCUsedFSlots + RFactoryM
		PCUsedBSlots = PCUsedBSlots + RBarracksM	
		PCFactoryM = PCFactoryM - RFactoryM
		PCBarracksM = PCBarracksM - RBarracksM
		CapitalS = 0 - RCapitalM
		ElementXS = 0 - RElementXM

		
		Dim MessageM
		
		MessageM = "<p class = 'normal'>Training has begun at a cost of "&RCapitalM&"c Capital and "&RElementXM&"t ElementX "
	End If
End If
	
	Dim MInfantryM, MElite1M, MElite2M, MElite3M, MTransportsM, Num1, Num2, Num3, Num4
	
	Num1 = Int(PCCapitalM/50)
	Num2 = PCBarracksM
	Num3 = Int(PCCivillianM)
	MInfantryM = MIN(Num1, Num2, Num3, 1000000000000)

	Num1 = Int(PCCapitalM/Elite1Cost)
	Num2 = Int(PCElementXM/Elite1CostX)
	Num3 = PCInfantryM
	Num4 = PCBarracksM
	MElite1M = MIN(Num1, Num2, Num3, Num4)

	Num1 = Int(PCCapitalM/Elite2Cost)
	Num2 = Int(PCElementXM/Elite2CostX)
	Num3 = PCInfantryM
	Num4 = PCFactoryM
	MElite2M = MIN(Num1, Num2, Num3, Num4)
	
	Num1 = Int(PCCapitalM/Elite3Cost)
	Num2 = Int(PCElementXM/Elite3CostX)
	Num3 = Int(PCInfantryM/2)
	Num4 = PCFactoryM\2
	MElite3M = MIN(Num1, Num2, Num3, Num4)
	
	Num1 = Int(PCCapitalM/200)
	Num2 = PCInfantryM
	Num3 = PCFactoryM
	MTransportsM = MIN(Num1, Num2, Num3, 100000000000000000)

RSTM.Close
RSDMM.Close
%>
