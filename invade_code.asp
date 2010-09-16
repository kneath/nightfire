<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="include.asp"-->
<% Response.Buffer = True %>
<%

'*********Equations*********
Dim LandEquation, DeathEquation, x

Function Land(x)
Land =  (1-2.718281828^((100 * x)/(100/Log(1/3))))*.15
End Function


LandEquation = (1-2.718281828^((100*x)/(100/Log(1/3))))*.15
'((1-2.718281828^(100x/(100/log(1/3))))*.15)\1
DeathEquation = 0
'***************************


Dim strUserName, SelectedSystem
strUserName = Request.Cookies("NightFire")("UserName")

Call OpenConnection()

Dim RSTargetStats, objRSinvade, objRS2invade, RSUser,  RSTargetInfo, RSUserMilitary, RSTargetMilitary, RSUserAttackingMilitary, RSBuildLand, RSBuilding, RSLand, RSUserLand

Set RSUser = Server.CreateObject("ADODB.Recordset")
Set RSUserMilitary = Server.CreateObject("ADODB.Recordset")
Set RSTargetInfo = Server.CreateObject("ADODB.Recordset")
Set RSTargetMilitary = Server.CreateObject("ADODB.Recordset")
Set RSTargetStats = Server.CreateObject("ADODB.Recordset")
Set RSUserAttackingMilitary = Server.CreateObject("ADODB.Recordset")
Set objRSinvade = Server.CreateObject("ADODB.RecordSet")
Set objRS2invade = Server.CreateObject("ADODB.RecordSet")
Set RSBuildLand = Server.CreateObject("ADODB.RecordSet")
Set RSBuilding = Server.CreateObject("ADODB.RecordSet")
Set RSLand = Server.CreateObject("ADODB.RecordSet")
Set RSUserLand = Server.CreateObject("ADODB.RecordSet")

Dim RSUpdate, SQLUpdate
SQLUpdate = "SELECT LastUpdate FROM GeneralInfo WHERE UserName = '"&strUserName&"'"
Set RSUpdate = Conn.Execute(SQLUpdate)

Dim UpdateCalc, LastUpdate, CurrentUpdate
LastUpdate = RSUpdate("LastUpdate")
CurrentUpdate = Now()
If DateDiff("h", LastUpdate, CurrentUpdate) > 1 Then
	Response.Write"<!--#include file = 'update.asp'-->"
Else
	'nothin' no need to update
End If
Call OpenConnection()


objRSinvade.Open "SELECT * FROM GeneralStats WHERE UserName = '"&strUserName&"'", Conn
RSUserMilitary.Open "SELECT * FROM DoneMilitary WHERE UserName = '"&strUserName&"'",Conn, adOpenStatic, adLockPessimistic
RSUser.Open "SELECT * FROM GeneralInfo WHERE UserName = '" & strUserName & "'", Conn, adOpenStatic, adLockPessimistic
RSUserAttackingMilitary.Open "SELECT * FROM Attacking WHERE UserName = '"&strUserName&"'", Conn, adOpenStatic, adLockPessimistic

Dim AvailibleTransports, PlayerRace, PlayerSystem, AvailibleInfantry, AvailibleElite1, AvailibleElite2, AvailibleElite3
PlayerRace = RSUser("Race")
PlayerSystem = RSUser("SystemNumber")
Set AvailibleInfantry = RSUserMilitary("Soldiers")
Set AvailibleElite1 = RSUserMilitary("Elite1")
Set AvailibleElite2 = RSUserMilitary("Elite2")
Set AvailibleElite3 = RSUserMilitary("Elite3")
Set AvailibleTransports = RSUserMilitary("Transports")

Dim Elite1Energy, Elite1Desc, Elite2Energy, Elite2Desc, Elite3Energy, Elite3Desc, Elite1Offense, Elite2Offense, Elite3Offense, Elite1Defense, Elite2Defense, Elite3Defense, Elite1Name, Elite2Name, Elite3Name

Dim RSStats
Set RSStats = Server.CreateObject("ADODB.Recordset")

RSStats.Open "SELECT Type, Race, Unit, Name,EnergyCost, Description, Offense, Defense FROM COSTS WHERE Type = 'Military' AND Race = '" & PlayerRace & "'", Conn, adOpenStatic, adLockPessimistic
RSStats.Find "Unit = 'Elite1'"
Elite1Offense = RSStats("Offense")
Elite1Defense = RSStats("Defense")
Elite1Name = RSStats("Name")
Elite1Energy = RSStats("EnergyCost")
Elite1Desc = RSStats("Description")
RSStats.Movenext
RSStats.Find "Unit = 'Elite2'"
Elite2Offense = RSStats("Offense")
Elite2Defense = RSStats("Defense")
Elite2Name = RSStats("Name")
Elite2Energy = RSStats("EnergyCost")
Elite2Desc = RSStats("Description")
RSStats.Movenext
RSStats.Find "Unit = 'Elite3'"
Elite3Offense = RSStats("Offense")
Elite3Defense = RSStats("Defense")
Elite3Name = RSStats("Name")
Elite3Energy = RSStats("EnergyCost")
Elite3Desc = RSStats("Description")
RSStats.Close

Dim UserTransports, UserNation, UserSystem
UserNation = RSUser("Nation")
UserRace = RSUser("Race")
UserSystem = RSUser("SystemNumber")
UserLand = objRSinvade("Land")
UserInfantry = Request("Infantry") +0
UserElite1 = Request("Elite1") + 0
UserElite2 = Request("Elite2") + 0
UserElite3 = Request("Elite3") + 0
If PlayerRace = "Larxon" Then
UserTransports = UserInfantry + UserElite2 + (UserElite3 * 2)
Else
UserTransports = UserInfantry + UserElite1 + UserElite2 + (UserElite3 * 2)
End If
UserTransports = UserTransports\20 + 0

If Request.Form("System") <> "" Then
	Response.Cookies("NightFire")("SelectedSystem") = Request.Form("System")
End If
If Request.Cookies("NightFire")("SelectedSystem") = ""  Then
	Response.Cookies("NightFire")("SelectedSystem") = PlayerSystem
End If
'they have some other system in there that needs to stay.
SelectedSystem = Request.Cookies("NightFire")("SelectedSystem")

Dim SQLLL
SQLLL = "SELECT * FROM GeneralInfo WHERE SystemNumber = "
SQLLL = SQLLL & SelectedSystem & ""
objRS2invade.Open SQLLL, Conn

'Start Select Target info


Dim TargetSystem, TargetPlayer, PriorityTerrain
TargetSystem = Request.Form("SystemNumber")
TargetPlayer = Request.Form("Nation")
PriorityTerrain = Request.Form("Terrain")

Dim SuccessString
Dim ErrorMessage
If Request.Form("attack") = "yes" Then
If RSUserMilitary("Morale") <= 50 Then
	ErrorMessage = "<p align = 'center' class = 'small'>Unfortunately, your men's morale is far too low to mount another attack.</p>"
Else
	If UserInfantry>AvailibleInfantry or UserElite1>AvailibleElite1 or UserElite2>AvailibleElite2 or UserElite3>AvailibleElite3 or UserTransports>AvailibleTransports Then
	
			ErrorMessage = "<p align = 'center' class = 'small'>I'm sorry but you do not have that many troops or transports</p>"
	Else
	'they can attack... I guess...
     
	If isNull(UserInfantry) Then
		UserInfantry = 0
	End If
	If isNull(UserElite1) Then
		UserElite1 = 0
	End If
	If isNull(UserElite2) Then
		UserElite2 = 0
	End If
	If isNull(UserElite3) Then
		UserElite3 = 0
	End If
	'All targets have been selected the attacking can go through now

	'*****Race Information******

	Dim TargetElite1Name, TargetElite2Name, TargetElite3Name, TargetElite1Offense, TargetElite2Offense, TargetElite3Offense, TargetElite1Defense, TargetElite2Defense, TargetElite3Defense
	RSUserMilitary("Morale") = RSUserMilitary("Morale") - 5
	RSUserMilitary.Update

	

	'****************************
	'Start Target Information


	Dim TargetUserName
	TargetUserName = Request("Nation")
	Dim UserNameU
	UserNameU = TargetUserName
%>
<!--#Include File = "targetupdate.asp"-->
<%

	RSTargetMilitary.Open "SELECT * FROM DoneMilitary WHERE UserName = '"&TargetUserName&"'", Conn, adOpenStatic, adLockPessimistic
	RSTargetInfo.Open "SELECT * FROM GeneralInfo WHERE UserName = '"&TargetUserName&"'", Conn
	RSTargetStats.Open "SELECT * FROM GeneralStats WHERE UserName = '"&TargetUserName&"'", Conn, 3, 3
	
If DateDiff("h", Now(), RSTargetInfo("Protection")) > 0 Then
	ErrorMessage = "<p class = 'alert'>That nation is under protection, you are not allowed to attack it</p>"
Else
	Dim TargetSystemNumber, TargetNation, TargetRace, TargetLand, TargetInfantry, TargetElite1, TargetElite2, TargetElite3, TargetTotalDefense, TargetMods
	Set TargetRace = RSTargetInfo("Race")
	Set TargetLand = RSTargetStats("Land")
	Set TargetSystemNumber = RSTargetInfo("SystemNumber")
	Set TargetNation = RSTargetInfo("Nation")
	Set TargetInfantry = RSTargetMilitary("Soldiers")
	Set TargetElite1 = RSTargetMilitary("Elite1")
	Set TargetElite2 = RSTargetMilitary("Elite2")
	Set TargetElite3 = RSTargetMilitary("Elite3")

	Dim Elite1OffenseTarget, Elite1DefenseTarget, Elite2OffenseTarget, Elite2DefenseTarget, Elite3OffenseTarget, Elite3DefenseTarget
	RSStats.Open "SELECT Type, Race, Unit, Name,EnergyCost, Description, Offense, Defense FROM COSTS WHERE Type = 'Military' AND Race = '" & TargetRace & "'", Conn, adOpenStatic, adLockPessimistic
	RSStats.Find "Unit = 'Elite1'"
	Elite1OffenseTarget = RSStats("Offense")
	Elite1DefenseTarget = RSStats("Defense")
	RSStats.Movenext
	RSStats.Find "Unit = 'Elite2'"
	Elite2OffenseTarget = RSStats("Offense")
	Elite2DefenseTarget = RSStats("Defense")
	RSStats.Movenext
	RSStats.Find "Unit = 'Elite3'"
	Elite3OffenseTarget = RSStats("Offense")
	Elite3DefenseTarget = RSStats("Defense")
	RSStats.Close

	Dim DefendingInfantry, DefendingElite1, DefendingElite2, DefendingElite3

	If Elite1DefenseTarget = 0 Then
		TargetElite1 = 0
	ElseIf Elite2DefenseTarget = 0 Then
		TargetElite2 = 0
	ElseIf Elite3DefenseTarget = 0 Then
		TargetElite3 = 0
	End If

	DefendingInfantry = TargetInfantry
	DefendingElite1 = TargetElite1
	DefendingElite2 = TargetElite2
	DefendingElite3 = TargetElite3

	
	Dim RSRT
	Set RSRT = Server.CreateObject("ADODB.Recordset")
	RSRT.Open "SELECT Military, UserName FROM RESEARCH Where UserName = '"&TargetUserName&"'", Conn
	If RSRT("Military") = 2 Then
		TargetMods = .1
	End If
	RSRT.Close
	TargetMods = TargetMods + 1
	TargetTotalDefense = TargetInfantry +  TargetElite1*Elite1DefenseTarget + TargetElite2*Elite2DefenseTarget + TargetElite3*Elite3DefenseTarget
	TargetTotalDefense = TargetTotalDefense * TargetMods


	'Start User Info
	Dim UserRace, UserLand, UserInfantry, UserElite1, UserElite2, UserElite3, UserTotalOffense, UserMods



	'Let's send the boys out!(They'll Die later)
	Dim TransportsNeeded
	TransportsNeeded = UserTransports
	TransportsNeeded = TransportsNeeded\20
	RSUserMilitary("UserName") = strUserName
	RSUserMilitary("Soldiers") = AvailibleInfantry - UserInfantry
	RSUserMilitary("Elite1") = AvailibleElite1 - UserElite1
	RSUserMilitary("Elite2") = AvailibleElite2 - UserElite2
	RSUserMilitary("Elite3") = AvailibleElite3 - UserElite3
	RSUserMilitary("Transports") = AvailibleTransports - TransportsNeeded
	RSUserMilitary.Update


	'Start Calculations

	Dim RSRUI
	Set RSRUI = Server.CreateObject("ADODB.Recordset")
	RSRUI.Open "SELECT Military, UserName FROM RESEARCH WHERE UserName = '"&strUserName&"'", Conn
	If RSRUI("Military") > 0 Then
		UserMods = .1
	End If
	RSRUI.Close
	UserMods = UserMods + 1
	UserElite1 = UserElite1
	UserElite2 = UserElite2
	UserElite3 = UserElite3
	UserTotalOffense = UserInfantry + UserElite1*Elite1Offense + UserElite2*Elite2Offense + UserElite3*Elite3Offense
	UserTotalOffense = UserTotalOffense * UserMods
	'Finally done calculating Offense and Defense

	Dim Success
	If TargetTotalDefense < UserTotalOffense Or TargetTotalDefense = UserTotalOffense Then
		Success = "Yes"
	Else
		Success = "No"
	End If


	'Somehow we have to kill them.. 
	Dim DeadInfantry, DeadElite1, DeadElite2, DeadElite3, CasRate
	CasRate = TargetTotalDefense/(UserTotalOffense + 1)*.1
	If CasRate > .1 Then
		CasRate = .99887
	End If
	DeadInfantry = (TargetInfantry * CasRate)\1
	If TargetRace = "Human" Then
		DeadElite1 = (TargetElite1 * (CasRate + .1))\1
		DeadElite3 = (TargetElite3 * (CasRate - .05))\1
	Else
		DeadElite1 = (TargetElite1 * CasRate)\1
		DeadElite3 = (TargetElite3 * CasRate)\1
	End If
	DeadElite2 = (TargetElite2 * .1)\1
	RSTargetMilitary("Soldiers") = TargetInfantry - DeadInfantry
	RSTargetMilitary("Elite1") = TargetElite1 - DeadElite1
	RSTargetMilitary("Elite2") = TargetElite2 - DeadElite2
	RSTargetMilitary("Elite3") = TargetElite3 - DeadElite3
	RSTargetMilitary.Update


	'Somehow we have to kill them.. 
	CasRate = UserTotalOffense/(TargetTotalDefense + 1)*.1
	If CasRate < .1 Then
		CasRate = .10035
	ElseIf CasRate > .25 Then
		CasRate = .25
	End If
	DeadInfantry = (UserInfantry * CasRate)\1
	If UserRace = "Human" Then
		DeadElite1 = (UserElite1 * (CasRate + .1))\1
		DeadElite3 = (UserElite3 * (CasRate - .05))\1
	Else
		DeadElite1 = (UserElite1 * CasRate)\1
		DeadElite3 = (UserElite3 * CasRate)\1
	End If
	DeadElite2 = (UserElite2 * .1)\1
	RSUserAttackingMilitary.AddNew
	RSUserAttackingMilitary("UserName") = strUserName
	RSUserAttackingMilitary("Soldiers") = UserInfantry - DeadInfantry
	RSUserAttackingMilitary("Elite1") = UserElite1 - DeadElite1
	RSUserAttackingMilitary("Elite2") = UserElite2 - DeadElite2
	RSUserAttackingMilitary("Elite3") = UserElite3 - DeadElite3
 	RSUserAttackingMilitary("Transports") = TransportsNeeded
 	RSUserAttackingMilitary("Probes") = 0
	RSUserAttackingMilitary("Agents") = 0
	RSUserAttackingMilitary("Mages") = 0
	RSUserAttackingMilitary("Time") = DateAdd( "h", 8, Now())
	RSUserAttackingMilitary.Update

	Dim RSMessages
	Set RSMessages = Server.CreateObject("ADODB.Recordset")
	RSMessages.Open "SELECT * FROM MESSAGES", Conn, adOpenStatic, adLockPessimistic

	If Success = "Yes" Then




		'Begin taking land
		
		RSBuildland.Open "SELECT * FROM DoneBuildings WHERE UserName = '" & TargetUserName & "'", Conn, 3, 3
		RSBuilding.Open "SELECT * FROM Building WHERE UserName = '" & TargetUserName & "'", Conn, 3, 3
		RSLand.Open "SELECT * FROM Land WHERE UserName = '" & TargetUserName & "'", Conn, 3, 3
		

		Dim Barracks, Factories, SolarPanels, Consulates, ResearchCenters, ElementXMines, DefenseStations, AtomicPlants, PlutoniumPlants, AntiMatterLabs, NuclearSilos, Temples, Cities


		Dim TotalLandGained, LandRemaining, Ratio, TargetPlainLand, TargetSmoothLand, TargetRuggedLand, UserDland, TargetPlutoniumLand

		TargetPlainLand = RSLand("Plain")
		TargetRuggedLand = RSLand("Rugged")
		FreeLand = RSLand("FreePlain") + 0
		Ratio = TargetLand/UserLand + 0
		TotalLandGained = Land(Ratio) + 0
		TotalLandGained = (TotalLandGained*(TargetPlainLand + TargetRuggedLand))\1
		If TotalLandGained < 50 Then
			TotalLandGained = 50
		End If
		LandRemaining = TotalLandGained + 0
		UserDLand = ((TargetRuggedLand*TotalLandGained/TargetLand))\1
		TargetRuggedLand = TargetRuggedLand - UserDLand
		LandRemaining = LandRemaining -	UserDLand

		RSMessages.AddNew
		RSMessages("UserName") = TargetUserName
		RSMessages("Message") = "<p> &nbsp;&nbsp;&nbsp; Forces from "&UserNation&"(#"&UserSystem&") broke through your defenses capturing "&FormatNumber(TotalLandGained, 0)&" square kilometers!</p>"
		RSMessages("Type") = "MainAttack"
		RSMessages("Seen") = "No"
		RSMessages("Time") = Now()
		RSMessages.Update

		If TargetSystemNumber <> UserSystem Then
		RSMessages.AddNew
		RSMessages("UserName") = TargetUserName
		RSMessages("Message") = "<p> &nbsp;&nbsp;&nbsp; Forces from "&UserNation&"(#"&UserSystem&") were victorious in capturing "&FormatNumber(TotalLandGained, 0)&" square kilometers from fellow Nation "&TargetNation&"</p>"
		RSMessages("Type") = "SystemNews"&TargetSystemNumber
		RSMessages("Seen") = "No"
		RSMessages("Time") = Now()
		RSMessages.Update
		End If

		RSMessages.AddNew
		RSMessages("UserName") = TargetUserName
		RSMessages("Message") = "<p> &nbsp;&nbsp;&nbsp; Forces from "&UserNation&" were victorious in capturing "&FormatNumber(TotalLandGained, 0)&" square kilometers from "&TargetNation&"(#"&TargetSystemNumber&")</p>"
		RSMessages("Type") = "SystemNews"&UserSystem
		RSMessages("Seen") = "No"
		RSMessages("Time") = Now()
		RSMessages.Update

		RSMessages.Close
		Dim FreeLand


		If TotalLandGained > FreeLand Then
			LandRemaining = LandRemaining - FreeLand
 			RSLand("Plain") = RSLand("Plain") - FreeLand
 			RSLand("FreePlain") = 0
			RSLand.Update
		ElseIf TotalLandGained < FreeLand Then
			LandRemaining = 0
			RSLand("Plain") = TargetPlainLand - TotalLandGained
			RSLand("FreePlain") = FreeLand - TotalLandGained
			RSLand.Update
		End If
			RSLand("Rugged") = TargetRuggedLand
			RSLand.Update

		RSTargetStats("Land") = TargetLand - TotalLandGained
		If RSTargetStats("Land") <= 0 Then
			Dim UserNameO
			UserNameO = TargetUserName
			Dim RSTemp
			Set RSTemp = Server.CreateObject("ADODB.Recordset")
	
			RSTemp.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM SYSTEMS WHERE SystemNumber="&TargetSystemNumber, Conn, adOpenStatic, adLockPessimistic
			RSTemp("Nations") = RSTemp("Nations") - 1
			RSTemp.Close
			RSTemp.Open "SELECT * FROM DONEMILITARY WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
						Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM TRAINING WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
							RSTemp.Delete
					RSTemp.MoveNext
				Loop		
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM ATTACKING WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM DONEBUILDINGS WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM BUILDING WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
							RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM LAND WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM USERS WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM MESSAGES WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM NUKES WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM RESEARCH WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM RESEARCHING WHERE UserName = '"&UserNameO&"'", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
End If
		RSTargetStats.Update

		If LandRemaining > 0 Then

			Dim LandGain
			LandGain = LandRemaining/(TargetPlainLand - FreeLand)

			If RSBuilding.EOF = FALSE Then
				Do While Not RSBuilding.EOF
					RSBuilding("Barracks") = RSBuilding("Barracks") - (RSBuilding("Barracks")*LandGain)\1
					RSBuilding("Factories") = RSBuilding("Factories") - (RSBuilding("Factories")*LandGain)\1
					RSBuilding("SolarPanels") = RSBuilding("SolarPanels") - (RSBuilding("SolarPanels")*LandGain)\1
					RSBuilding("Consulates") = RSBuilding("Consulates") - (RSBuilding("Consulates")*LandGain)\1
					RSBuilding("ReasearchCenters") = RSBuilding("ReasearchCenters") - (RSBuilding("ReasearchCenters")*LandGain)\1
					RSBuilding("ElementXMines") = RSBuilding("ElementXMines") - (RSBuilding("ElementXMines")*LandGain)\1
					RSBuilding("DefenseStations") = RSBuilding("DefenseStations") - (RSBuilding("DefenseStations")*LandGain)\1
					RSBuilding("AtomicPlants") = RSBuilding("AtomicPlants") - (RSBuilding("AtomicPlants")*LandGain)\1
					RSBuilding("PlutoniumPlants") = RSBuilding("PlutoniumPlants") - (RSBuilding("PlutoniumPlants")*LandGain)\1
					RSBuilding("AntiMatterLabs") = RSBuilding("AntiMatterLabs") - (RSBuilding("AntiMatterLabs")*LandGain)\1
					RSBuilding("NuclearSilos") = RSBuilding("NuclearSilos") - (RSBuilding("NuclearSilos")*LandGain)\1
					RSBuilding("Temples") = RSBuilding("Temples") - (RSBuilding("Temples")*LandGain) \1
					RSBuilding("Cities") = RSBuilding("Cities") - (RSBuilding("Cities")*LandGain) \1
					LandRemaining = LandRemaining - ((RSBuilding("Barracks")*LandGain)\1 + (RSBuilding("Factories")*LandGain)\1 + (RSBuilding("SolarPanels")*LandGain)\1 + (RSBuilding("Consulates")*LandGain)\1 + (RSBuilding("ReasearchCenters")*LandGain)\1 + (RSBuilding("ElementXMines")*LandGain)\1 + (RSBuilding("ElementXMines")*LandGain)\1 + (RSBuilding("DefenseStations")*LandGain)\1 + (RSBuilding("AtomicPlants")*LandGain)\1 + (RSBuilding("PlutoniumPlants")*LandGain)\1 + (RSBuilding("AntiMatterLabs")*LandGain)\1 + (RSBuilding("NuclearSilos")*LandGain)\1 + (RSBuilding("Temples")*LandGain) \1 + (RSBuilding("Cities")*LandGain) \1)
					RSBuilding.Update
					RSBuilding.MoveNext
				Loop
			Else
			'nothing building
			End If
			If LandRemaining > 0 Then
				RSBuildLand("Barracks") = RSBuildLand("Barracks") - (RSBuildLand("Barracks")*LandGain) \1
				RSBuildLand("Factories") = RSBuildLand("Factories") - (RSBuildLand("Factories")*LandGain)\1
				RSBuildLand("SolarPanels") = RSBuildLand("SolarPanels") - (RSBuildLand("SolarPanels")*LandGain)\1
				RSBuildLand("Consulates") = RSBuildLand("Consulates") - (RSBuildLand("Consulates")*LandGain)\1
				RSBuildLand("ReasearchCenters") = RSBuildLand("ReasearchCenters") - (RSBuildLand("ReasearchCenters")*LandGain)\1
				RSBuildLand("ElementXMines") = RSBuildLand("ElementXMines") - (RSBuildLand("ElementXMines")*LandGain)\1
				RSBuildLand("DefenseStations") = RSBuildLand("DefenseStations") - (RSBuildLand("DefenseStations")*LandGain)\1
				RSBuildLand("AtomicPlants") = RSBuildLand("AtomicPlants") - (RSBuildLand("AtomicPlants")*LandGain)\1
				RSBuildLand("PlutoniumPlants") = RSBuildLand("PlutoniumPlants") - (RSBuildLand("PlutoniumPlants")*LandGain)\1
				RSBuildLand("AntiMatterLabs") = RSBuildLand("AntiMatterLabs") - (RSBuildLand("AntiMatterLabs")*LandGain)\1
				RSBuildLand("NuclearSilos") = RSBuildLand("NuclearSilos") - (RSBuildLand("NuclearSilos")*LandGain)\1
				RSBuildLand("Temples") = RSBuildLand("Temples") - (RSBuildLand("Temples")*LandGain) \1
				RSBuildLand("Cities") = RSBuildLand("Cities") - (RSBuildLand("Cities")*LandGain) \1
				RSBuildLand.Update
			End If
		Else
			'all barren was taken
		End If


		''let's give the user thier land back now 
		UserDLand = UserDLand + (TotalLandGained*.1)
		RSUserLand.Open "SELECT * FROM Building WHERE UserName = '" & strUserName & "'", Conn, 3, 3
		RSUserLand.AddNew
		RSUserLand("UserName") = strUserName
		RSUserLand("Barracks") = 0
		RSUserLand("Factories") = 0
		RSUserLand("SolarPanels") = 0
		RSUserLand("Consulates") = 0
		RSUserLand("ReasearchCenters") = 0
		RSUserLand("ElementXMines") = 0
		RSUserLand("Cities") = 0
		RSUserLand("DefenseStations") = 0
		RSUserLand("AtomicPlants") = 0
		RSUserLand("PlutoniumPlants") = 0
		RSUserLand("AntimatterLabs") = 0
		RSUserLand("NuclearSilos") = 0
		RSUserLand("Temples")  = 0
		RSUserLand("AntimatterPlants") = 0
		RSUserLand("Cities") = 0
		RSUserLand("ElementXLand") = 0
		RSUserLand("PlutoniumLand") = 0
		RSUserLand("AntimatterLand") = 0
		RSUserLand("SmoothLand") = 0
		RSUserLand("RuggedLand") = UserDland
		RSUserLand("PlainLand") = TotalLandGained - UserDland
		RSUserLand("Time") = DateAdd("h", 8, Now())
		RSUserLand.Update

		RSUserLand.Close
		RSLand.Close
		RSBuildLand.Close
		RSBuilding.Close

		SuccessString = "<p class = 'small' align = 'center'>Your Forces Reach the Nation of <span class = 'normal'>"&TargetNation&"</span> and the battle begins...</p><p class = 'small' align = 'center'><br>You smile proudly as you hear screams of victory.  You lost <span class = 'alert'>"&FormatNumber(DeadInfantry, 0)&"</span> Infantry, <span class = 'alert'>"&FormatNumber(DeadElite1, 0)&"</span> "&Elite1Name&" , <span class = 'alert'>"&FormatNumber(DeadElite2, 0)&"</span> "&Elite2Name&" , <span class = 'alert'>"&FormatNumber(DeadElite3, 0)&"</span> "&Elite3Name&". <br> Your Forces took  <span class = 'alert'>"&FormatNumber(TotalLandGained - UserDland, 0)&"</span> km of Plain land, and <span class = 'alert'>"&FormatNumber(UserDland, 0)&"</span>km of land was destroyed in the battle</p>"
	Else
		RSMessages.AddNew		
		RSMessages("UserName") = TargetUserName
		RSMessages("Message") = "<p> &nbsp;&nbsp;&nbsp; Forces from "&UserNation&"(#"&UserSystem&") attempted to sieze your territories, but we managed to fight them off!</p>"
		RSMessages("Type") = "MainAttack"
		RSMessages("Seen") = "No"
		RSMessages("Time") = Now()
		RSMessages.Update
		RSMessages.AddNew
		RSMessages("UserName") = TargetUserName
		RSMessages("Message") = "<p> &nbsp;&nbsp;&nbsp; Forces from "&UserNation&" attempted to sieze territories from "&TargetNation&"(#"&TargetSystemNumber&") but were overwhelmed and were forced to retreat</p>"
		RSMessages("Type") = "SystemNews"&UserSystem
		RSMessages("Seen") = "No"
		RSMessages("Time") = Now()
		RSMessages.Update
		If TargetSystemNumber <> UserSystem Then
		RSMessages.AddNew
		RSMessages("UserName") = TargetUserName
		RSMessages("Message") = "<p> &nbsp;&nbsp;&nbsp; Forces from "&UserNation&"(#"&UserSystem&") attempted to sieze territories from "&TargetNation&".</p>"
		RSMessages("Type") = "SystemNews"&TargetSystemNumber
		RSMessages("Seen") = "No"
		RSMessages("Time") = Now()
		RSMessages.Update
		End If

		RSMessages.Close
		SuccessString = "<p class = 'small' align = 'center' >Your Forces Reach the Nation of <span class = 'normal'>"&TargetNation&"</span> and the battle begins...</p><p align = 'center' class = 'small'><br>You hear dissapointing calls of defeat. You lost <span class = 'alert'>"&FormatNumber(DeadInfantry, 0)&"</span> Infantry, <span class = 'alert'>"&FormatNumber(DeadElite1, 0)&"</span> "&Elite1Name&" , <span class = 'alert'>"&FormatNumber(DeadElite2,0)&"</span> "&Elite2Name&" , <span class = 'alert'>"&FormatNumber(DeadElite3,0)&"</span> "&Elite3Name&".</p>"
	End If
End If
End If
End If
End If

%>
