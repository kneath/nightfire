
<%



If Request.Querystring("f") = "landbonus" Then

	Dim UserNameLLL
	UserNameLLL = Request.Cookies("NightFire")("UserName")
	Call OpenConnection()
	Dim RSLLL
	Set RSLL = Conn.Execute("SELECT LandBonus, UserName FROM Land WHERE UserName = '"&UserNameLLL&"'")
	If RSLL("LandBonus") = 1 Then
		RSLL.Close
		Set RSLLL = Server.CreateObject("ADODB.Recordset")
		RSLL.Open "Land", Conn, 3, 3
		RSLL("LandBonus") = 0
		RSLL("FreePlain") = RSLL("FreePlain") + 25
		RSLL("Plain") = RSLL("Plain") + 25
		RSLL.Update
		RSLL.Close
		RSLL.Open "SELECT Land, UserName FROM GENERALINFO WHERE UserName = '"&UserNameLLL&"'", Conn, 3, 3
		RSLL("Land") = RSLL("Land") + 25
		RSLL.Update
		RSLL.Close
		Response.Write "<p class = 'normal'> You just gained 25km of Plain land! Thank you for supporting NightFire</p>"
	End If
End If
%>
<center>
<p> No banners yet</p>
<a href = "#" onclick = "window.open('ad2.asp?f=landbonus','admessage',
'width=270,height=120,status=0,location=0,left=1,top=1,x=1,y=1')" >Click here to get land bonus</a>
</center>