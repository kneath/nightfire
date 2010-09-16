<!--#Include File = "include.asp"-->
<%
Dim RSTemp
Set RSTemp = Server.CreateObject("ADODB.Recordset")
If Request.Form("password") = "Warpspire" Then
	
			RSTemp.Open "SELECT * FROM GENERALSTATS", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM GENERALINFO", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM DONEMILITARY", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
						Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM TRAINING", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
							RSTemp.Delete
					RSTemp.MoveNext
				Loop		
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM ATTACKING", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM DONEBUILDINGS", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM BUILDING", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
							RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM LAND", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM USERS", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM NUKES", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM RESEARCH", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM RESEARCHING", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close
			RSTemp.Open "SELECT * FROM MESSAGES", Conn, adOpenForwardOnly, adLockPessimistic
			If RSTemp.EOF = FALSE Then
				Do While RSTemp.EOF = FALSE
					RSTemp.Delete
					RSTemp.MoveNext
				Loop
			End If
			RSTemp.Close	
			Conn.Execute("DELETE * FROM MEETING")
			Conn.Execute("DELETE * FROM SYSTEMS")
%>
<!--#Include File = "systems.asp"-->
<%
End If

%>
<html>
<body>
<form action = "reset.asp" method = "post">
Reset the round!
<input name = "password" length=5>
<input type = "submit">
</form>
</html>
