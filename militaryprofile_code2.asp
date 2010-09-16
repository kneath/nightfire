<%
'Get rid of old values
Erase arrInfantry
Erase arrElite1
Erase arrElite2
Erase arrElite3
Erase arrTransports

'Get info for attacking units
RSAM.Open "SELECT * FROM ATTACKING WHERE UserName = '" & UserNameM & "'", Conn, adOpenStatic, adLockReadOnly
If RSAM.EOF = FALSE Then
	Do While RSAM.EOF = FALSE
		iIndex = DateDiff("h", Now(), RSAM("Time"))
		arrInfantry(iIndex) = arrInfantry(iIndex) + RSAM("Soldiers")
		arrElite1(iIndex)  = arrElite1(iIndex) + RSAM("Elite1")
		arrElite2(iIndex)  = arrElite2(iIndex) + RSAM("Elite2")
		arrElite3(iIndex)  = arrElite3(iIndex) + RSAM("Elite3")
		arrTransports(iIndex) = arrTransports(iIndex) + RSAM("Transports")
		RSAM.MoveNext
	Loop
	RSAM.MoveFirst
	RSAM.Close
End If

'Fill in empty spaces
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