<%
'Set the recordsets
Dim RSDMM, RSTM, RSAM, RSSM
Set RSDMM = Server.CreateObject("ADODB.Recordset")
Set RSTM = Server.CreateObject("ADODB.Recordset")
Set RSAM = Server.CreateObject("ADODB.Recordset")
Set RSSM = Server.CreateObject("ADODB.Recordset")

'Get their race
Dim RaceM, RSTempM
Set RSTempM = Conn.Execute("SELECT UserName, Race FROM GENERALINFO WHERE UserName = '" & UserNameM & "'")
RaceM = RSTempM("Race")
RSTempM.Close

'Get the unit info
Dim Elite1NM, Elite2NM, Elite3NM
RSSM.Open "SELECT Type, Unit, Race, Name FROM COSTS WHERE Type = 'Military' AND Race = '" & RaceM & "'", Conn, adOpenStatic
RSSM.Find "Unit = 'Elite1'"
Elite1NM = RSSM("Name")
RSSM.Find "Unit = 'Elite2'"
Elite2NM = RSSM("Name")
RSSM.Find "Unit = 'Elite3'"
Elite3NM = RSSM("Name")
RSSM.Close

'Get all finished at-home units
RSDMM.Open "SELECT * FROM DONEMILITARY WHERE UserName = '" & UserNameM & "'", Conn
Dim DInfantry, DElite1, DElite2, DElite3, DTransports
DInfantry = RSDMM("Soldiers")
DElite1 = RSDMM("Elite1")
DElite2 = RSDMM("Elite2")
DElite3 = RSDMM("Elite3")
DTransports = RSDMM("Transports")
RSDMM.Close

'Get info for training units
Dim iIndex, TTime, TInfantry, TElite1, TElite2, TElite3, arrInfantry(17), arrElite1(17), arrElite2(17), arrElite3(17), arrTransports(17)
RSTM.Open "SELECT * FROM TRAINING WHERE UserName = '" & UserNameM & "'", Conn, adOpenStatic, adLockReadOnly
If RSTM.EOF = FALSE Then
	Do While RSTM.EOF = FALSE
		iIndex = DateDiff("h", Now(), RSTM("Time"))
		arrInfantry(iIndex) = arrInfantry(iIndex) + RSTM("Soldiers")
		arrElite1(iIndex)  = arrElite1(iIndex) + RSTM("Elite1")
		arrElite2(iIndex)  = arrElite2(iIndex) + RSTM("Elite2")
		arrElite3(iIndex)  = arrElite3(iIndex) + RSTM("Elite3")
		arrTransports(iIndex) = arrTransports(iIndex) + RSTM("Transports")
		RSTM.MoveNext
	Loop
	RSTM.MoveFirst
	RSTM.Close
End If

'Fill in empty spaces and add up totals
Dim TotInfantry, TotElite1, TotElite2, TotElite3, TotTransports
TotInfantry = 0
TotElite1 = 0
TotElite2 = 0
TotElite3 = 0
TotTransports = 0
For iIndex = 0 to 17
	If arrInfantry(iIndex) = "" OR arrInfantry(iIndex) = "0" Then
		arrInfantry(iIndex) = "-"
	Else
		TotInfantry = TotInfantry + arrInfantry(iIndex) + 0
	End If
	If arrElite1(iIndex) = "" OR arrElite1(iIndex) = "0" Then
		arrElite1(iIndex) = "-"
	Else
		TotElite1 = TotElite1 + arrElite1(iIndex) + 0
	End If
	If arrElite2(iIndex) = "" OR arrElite2(iIndex) = "0" Then
		arrElite2(iIndex) = "-"
	Else
		TotElite2 = TotElite1 + arrElite2(iIndex) + 0
	End If
	If arrElite3(iIndex) = "" OR arrElite3(iIndex) = "0" Then
		arrElite3(iIndex) = "-"
	Else
		TotElite3 = TotElite3 + arrElite3(iIndex) + 0
	End If
	If arrTransports(iIndex) = "" OR arrTransports(iIndex) = "0" Then
		arrTransports(iIndex) = "-"
	Else
		TotTransports = TotTransports + arrTransports(iIndex) + 0
	End If
Next
%>