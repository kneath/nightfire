<%@ Language=VBScript %>

<% Option Explicit %>

<!--#include file="adovbs.inc"-->

<% Response.Buffer = True



Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")

objConn.ConnectionString="DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & Server.MapPath("\NightFire\db\mydb.mdb")

objConn.Open


Dim strSQL
Dim objRS
Dim strUserName
Dim strPassword
Dim IP, LoginTime

IP = Request.Servervariables("REMOTE_ADDR")
LoginTime = Now()

strUserName = Request.Form("UserName")
strPassword = Request.Form("Password")

If strUserName = "" Then

strUserName = Request.Cookies("NightFire")("UserName")
strPassword = Request.Cookies("NightFire")("Password")
End If

Set objRS = Server.CreateObject("ADODB.Recordset")

strSQL = "SELECT * FROM Users WHERE UserName = '" & strUserName & "' AND Password = '" & strPassword & "'"

objRS.Open strSQL, objConn, adOpenStatic, adLockPessimistic

%>

<HTML>
<HEAD>
 <TITLE>NightFire</TITLE>
<link rel=stylesheet href="game.css" type="text/css"> 
<%
IF objRS.EOF = True THEN
%>
	</HEAD>
	<DIV>
	</DIV><BR><BR>
    <BODY BGCOLOR='#000000' text='#ffffff' link='#ffffff' vlink='#ffffff' alink='ffcc33' MARGINWIDTH=0 MARGINHEIGHT=0 TOPMARGIN=0 BOTTOMMARGIN=0 LEFTMARGIN=0 RIGHTMARGIN=0><div><CENTER><div class = "normal">Sorry, but your request to login failed.  Please try again.</div><FORM ACTION='start.asp' METHOD='POST'><table><tr><td><div>Username: </div></td><td><INPUT MAXLENGTH='25' NAME='UserName'SIZE='20' TYPE='TEXT' class = 'numbar'></td></tr><tr><td><div>Password: </div></td><td><INPUT MAXLENGTH='25' NAME='Password' SIZE='20' TYPE='PASSWORD' class = 'numbar'></td></tr><tr><td align='center' colspan='2'><INPUT TYPE='submit' VALUE='Login' class = 'button'></td></tr></table></body>
<%
ELSE

Response.Cookies("NightFire")("UserName") = strUserName
Response.Cookies("NIghtFire")("Password") = strPassword
Response.Cookies("NightFire").Expires = DateAdd("h", 2, Now())
objRS("IP") = IP
objRS("LoginTIme") = LoginTime
objRS.Update

Dim RS
Set RS = objConn.Execute("SELECT UserName, SystemNumber FROM GENERALINFO WHERE UserName = '"&strUserName&"'")
Dim SystemNumber
SystemNumber = RS("SystemNumber")
RS.Close
Response.Cookies("NightFire")("SelectedSystem") = SystemNumber

Response.Redirect "main.asp"
END IF

objRS.Close
objConn.Close
Set objRS = Nothing
Set objConn = Nothing
%>
</HTML>
