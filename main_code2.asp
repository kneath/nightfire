<%
Dim RSMessages, y
y = 1

Set RSMessages = Server.CreateObject("ADODB.Recordset")
RSMessages.Open "SELECT * FROM MESSAGES WHERE UserName = '"&UserNameMM&"' AND Type = 'MainAttack' AND Seen = 'No'", ConnMM, adOpenStatic, adLockPessimistic

If RSMessages.EOF = False Then
Response.Write "<table><tr bgcolor='#ee5900'><td align = 'center' colspan=2>Recent News</td></tr>"
Do While RSMessages.EOF = FALSE
If y = 1 Then
Response.Write "<tr> <td bgcolor = '#000066'><div>Time: "&RSMessages("Time")&" </div></td><td bgcolor = '#000033'><div class = 'alert'>"&RSMessages("Message")&"</div></td></tr>"
Else
Response.Write "<tr> <td bgcolor = '#000033'><div>Time: "&RSMessages("Time")&" </div></td><td bgcolor = '#000066'><div class = 'alert'>"&RSMessages("Message")&"</div></td></tr>"
End If
RSMessages("Seen") = "Yes"
RSMessages.Update
RSMessages.MoveNext
Loop
Response.Write "</table>"
Else
Response.Write "<span class = 'normal'><center>No new recent news</center></span><br>"
End If

If Request.Querystring("news") = "yes" Then
	RSMessages.Close
	RSMessages.Open "SELECT * FROM MESSAGES WHERE UserName = '"&UserNameMM&"' AND Type = 'MainAttack'", ConnMM, adOpenStatic, adLockPessimistic

	If RSMessages.EOF = False Then
		Response.Write "<table><tr bgcolor='#ee5900'><td align = 'center' colspan=2>Recent News</td></tr>"
		Do While RSMessages.EOF = FALSE
			If y = 1 Then
				Response.Write "<tr> <td bgcolor = '#000066'><div>Time: "&RSMessages("Time")&" </div></td><td bgcolor = '#000033'><div class = 'alert'>"&RSMessages("Message")&"</div></td></tr>"
			Else
				Response.Write "<tr> <td bgcolor = '#000033'><div>Time: "&RSMessages("Time")&" </div></td><td bgcolor = '#000066'><div class = 'alert'>"&RSMessages("Message")&"</div></td></tr>"
			End If
			RSMessages.MoveNext
		Loop
		Response.Write "</table>"
	End If
Else
		Response.Write "<div><a href = 'main.asp?news=yes' class='menu_main' style='width:200pt'>See All Recent News</a></div>"
End If
RSMessages.Close
RSMessages.Open "SELECT * FROM MESSAGES WHERE UserName = '"&UserNameMM&"' AND Type = 'PersonalMessage' AND Seen = 'No'", ConnMM, adOpenStatic, adLockPessimistic

If RSMessages.EOF = FALSE Then
Response.Write "<a href = 'messages.asp'><p align = 'center' class = 'normal'>New messages</p></a>"
Else
Response.Write "<p class = 'normal' align = 'center'>No new messages</p>"
End If

Timer2 = Timer()
 TotalTime = Timer2 - Timer1

Response.Write "<p class = 'small'> This page took "&FormatNumber(TotalTime, 5)&" seconds</p>"

Set ConnMM = nothing
Set RSGIMM = nothing
Set RSGSMM = nothing
Set RSDMMM = nothing
Set RSMessages = nothing
%>