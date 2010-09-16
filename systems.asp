<%
Dim RS
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open "SYSTEMS", Conn, adOpenStatic, adLockPessimistic

Dim i, x
x = 1

For i = 1 to 400
	Response.Write i & ". . ."
	RS.AddNew
	RS("SystemNumber") = i
	RS("Nations") = 0
	RS("SystemName") = "System #" & i
	Select Case x
		Case 1
			RS("SystemType") = "Generic Star"
			x = 2
		Case 2
			RS("SystemType") = "Black Hole"
			x = 3
		Case 3
			RS("SystemType") = "Red Giant"
			x = 4
		Case 4
			RS("SystemType") = "Brown Dwarf"
			x = 1
	End Select
	RS.Update
Next
RS.Close
%>