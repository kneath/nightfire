<%Option Explicit%>
<!--#Include File = "include.inc"-->

<%
Server.Scripttimeout = 10
'Coded entirely by Duke Brak by hand
'This code may not be used in any other games than NightFire, a game by WarpSpire.
'www.warpspire.com/nightfire/
'Version 1.1 Alpha

'This script updates The User

'Get Thier UserName

Dim UserNameU
UserNameU = Request.Cookies("NightFire")("UserName")

%>
<!--#Include File = "update_code.asp"-->
<%
'Send the user back to the main page
Response.Redirect "main.asp"

%>
