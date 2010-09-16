<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="include.asp"-->
<% Response.Buffer = True %>


<%
Dim UserNameO
UserNameO = Request.Cookies("NightFire")("UserName")

Dim ConnO

Set ConnO = Server.CreateObject("ADODB.Connection")

ConnO.ConnectionString="DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & Server.MapPath("\NightFire\db\mydb.mdb")

ConnO.Open

Dim RSO

Set RSO = Server.CreateObject("ADODB.Recordset")

RSO.Open "SELECT * FROM Users WHERE UserName = '" & UserNameO & "'", ConnO, 3, 3

Dim SuccessCodeO, ErrorCodeO
If Request("NewPassWord") <> "" Then
	If Request("OldPassWord") = RSO("Password") Then
		RSO.MoveFirst
		RSO("Password") = Request("NewPassWord")
		RSO.Update
		SuccessCodeO = "Congratulations, you password has been successfully changed."
	Else
		ErrorCodeO = "Sorry, but your Old passwords do not match.  Please try again."
	End If
End If

Dim DeleteOption, SystemNumber

If Request.Form("WhatDo") = "Delete1" Then
	DeleteOption = "Delete2"
	ErrorCodeO = "Notice: Pressing Delete Account again will delete your account, there is no going back from this option"
Else
	DeleteOption = "Delete1"
End If

If Request.Form("WhatDo") = "Delete2" Then
	Dim RSTemp
	Set RSTemp = Server.CreateObject("ADODB.Recordset")
	
	RSTemp.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		SystemNumber = RSTemp("SystemNumber")
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM SYSTEMS WHERE SystemNumber="&SystemNumber, Conn, adOpenStatic, adLockPessimistic
	RSTemp("Nations") = RSTemp("Nations") - 1
	RSTemp.Close
	RSTemp.Open "SELECT * FROM DONEMILITARY WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM TRAINING WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM ATTACKING WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM DONEBUILDINGS WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM BUILDING WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM LAND WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close
	RSTemp.Open "SELECT * FROM USERS WHERE UserName = '"&UserNameO&"'", ConnO, adOpenForwardOnly, adLockPessimistic
	If RSTemp.EOF = FALSE Then
		Do While RSTemp.EOF = FALSE
			RSTemp.Delete
			RSTemp.MoveNext
		Loop
	End If
	RSTemp.Close

	Response.Redirect "index.asp"
End If
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

<body bgcolor="#000000" text="#ffffff" link="#ffffff" vlink="#ffffff" alink="ffcc33" >

<div>
<CENTER>Account Options</CENTER><BR>
<CENTER>
<FORM ACTION="options.asp" METHOD="POST">
<%
If ErrorCodeO <> "" Then
	Response.Write "<div>" & ErrorCodeO & "</div><br>"
End If
If SuccessCodeO <> "" Then
	Response.Write "<div>" & SuccessCodeO & "</div><br>"
End If
%>
<table>
<tr><td bgcolor="#0000CC" colspan=2><div><center>Change Password</center></div></td></tr>

<tr>
<td>Old Password: </td>
<td><INPUT MAXLENGTH="25" NAME="OldPassWord" SIZE="20" TYPE="PASSWORD" class = "numbar"></td>

</tr>
<tr>
<td>New Password: </td>
<td><INPUT MAXLENGTH="25" NAME="NewPassWord" SIZE="20" TYPE="PASSWORD" class = "numbar"></td>

</tr>
<tr>
<td align="center" colspan="2"><INPUT TYPE="submit" VALUE="Change" class = "button"><br><br></td>

</tr>
</form>
<form action=options.asp method=post>
<tr>
<td bgcolor = "#0000CC" colspan=2><div><center>Delete Account</center</div></td>
</tr>
<tr>
<td align=center colspan=2><input type=hidden name=WhatDo value="<%=DeleteOption%>"><input type=submit value="Delete Account" class = "button"></td>
</tr>
</form>
</table>

</center>
</div>
</body>
</html>
