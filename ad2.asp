
<%

Dim Conn, OpenConn
Sub OpenConnection()
	If OpenConn <> TRUE Then
		OpenConn = TRUE
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.ConnectionString="PROVIDER=MSDASQL;DRIVER={SQL Server};SERVER=127.0.0.1;DATABASE=warpspire;UID=sa;PWD=Kt6261;"
		Conn.Open
	End If
End Sub

If Request.Querystring("f") = "landbonus" Then

	Dim UserNameLLL
	UserNameLLL = Request.Cookies("NightFire")("UserName")
	Call OpenConnection()
	Dim RSLLL
	Set RSLL = Conn.Execute("SELECT LandBonus, LandBonTime, UserName FROM Land WHERE UserName = '"&UserNameLLL&"'")
	If DateDiff("d", RSLL("LandBonTime"), Now()) >= 1 Then	
		RSLL.Close
		Set RSLLL = Server.CreateObject("ADODB.Recordset")
		RSLL.Open "SELECT * FROM LAND WHERE USERNAME = '"&UserNameLLL&"'", Conn, 3, 3
		RSLL("LandBonTime") = Now()
		RSLL("FreePlain") = RSLL("FreePlain") + 25
		RSLL("Plain") = RSLL("Plain") + 25
		RSLL.Update
		RSLL.Close
		RSLL.Open "SELECT Land, UserName FROM GENERALSTATS WHERE UserName = '"&UserNameLLL&"'", Conn, 3, 3
		RSLL("Land") = RSLL("Land") + 25
		RSLL.Update
		RSLL.Close
		Response.Write "<p class = 'normal'> You just gained 25km of Plain land! Thank you for supporting NightFire</p>"
	End If
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

	<body bgcolor="#000000" text="#ffffff" link="#ffffff" vlink="#ffffff" alink="ffcc33" MARGINWIDTH=0 MARGINHEIGHT=0 TOPMARGIN=0 BOTTOMMARGIN=0 LEFTMARGIN=0 RIGHTMARGIN=0>
<center>
<p> No banners yet</p>
<a href = "ad2.asp?f=landbonus" target = "new">Click here to get land bonus</a>
</center>
