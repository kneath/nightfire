<!--#Include File = "military_code.asp"-->
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



<body bgcolor="#000000" text="#ffffff">
	<SCRIPT SRC="fieldcheck.js"></SCRIPT>
	<center>
		<!--#Include File = "ad.asp"-->
		<!--#Include File = "statbar.asp"-->
		<%=MessageM%>
		<%=ErrorMessage%>

	<table cellspacing=0 cellpadding=5 width="100%">

		<tr>
			<td bgcolor="#0000CC" colspan=5><div align=center class="large">Military - Train</div></td>
		</tr>

		<tr>
			<td bgcolor="#000066" colspan=5><div align=center><a href="military.asp" class="menu_main" style="width:40pt">Train</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="buymilitary.asp" class="menu_main" style="width:40pt">Buy</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="automilitary.asp" class="menu_main" style="width:40pt">Auto</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="sellmilitary.asp" class="menu_main" style="width:40pt">Sell</a>&nbsp;&nbsp;&nbsp;&nbsp;
				</div>
			</td>
		</tr>

		<tr>
			<form action="military.asp" method="post" name = "statbar">
			<td bgcolor="#000066" colspan=5>
				<div align=center><b>Cost:</b> <input name = "statbar1" type=text size=50 value="put mouse over unit to see costs" class = "statbar"><br><b>Stats:<input name = "statbar2" type=text size=50 value="put mouse over unit to see stats" class = "statbar"></b></div>
				<br>
				Total Factory Slots: <span class="normal"><%=(PCFactoryM + PCUsedFSlots)%></span> &nbsp;&nbsp; Availible Factory Slots: <span class="normal"><%=PCFactoryM%></span>
				<br>
				Total Barracks Slots: <span class="normal"><%=(PCBarracksM + PCUsedBSlots)%></span> &nbsp;&nbsp; Availible Barracks Slots: <span class="normal"><%=PCBarracksM%></span>
			</td>
			</form>
			<form action="military.asp" method="post">
		</tr>

		<tr bgcolor="#000044">
			<td>
				<div>Unit
				</div>
			</td>
			<td>
				<div>Processed</div>
			</td>
			<td>
				<div>Ordering</div>
			</td>

			<td><div>Max</div></td>

			<td><div>Order</div></td>

		</tr>
		<tr bgcolor="#000044">

			<td>
				<div><a style="width: 160; padding: 4" href="#" onMouseOver="statbar1('<%=InfantryStat%>', ' <%=InfantryDesc%> ')">Infantry:</a></div>
			</td>

			<td>
				<div><%=AInfantryM%></div>
			</td>

			<td>
				<div><%=TInfantryM%></div>
			</td>

			<td>
				<div><%=MInfantryM %></div>
			</td>

			<td>
				<div><input type="text" size=10 value="0" name="Soldiers" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></div>
			</td>

		</tr>

		<tr bgcolor="#000044">

			<td>
				<div><a style="width: 160; padding: 4" href="#" onMouseOver="statbar1('<%=Elite1Stat%>', '<%=Elite1Desc%>')"><%=Elite1Name%>:</a></div>
			</td>
			<td>
				<div><%=AElite1M%></div>
			</td>
			<td>
				<div><%=TElite1M%></div>
			</td>
			<td>
				<div><%=MElite1M%></div>
			</td>
			<td>
				<div><input type="text" size=10 value="0" name="Elite1" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></div>
			</td>

		</tr>
		<tr bgcolor="#000044">
			<td>
				<div><a style="width: 160; padding: 4" href="#" onMouseOver="statbar1('<%=Elite2Stat%>', '<%=Elite2Desc%>')"><%=Elite2Name%>:</a></div>
			</td>
			<td>
				<div><%=AElite2M%></div>
			</td>
			<td>
				<div><%=TElite2M%></div>
			</td>
			<td>
				<div><%=MElite2M%></div>
			</td>
			<td>
				<div><input type="text" size=10 value="0" name="Elite2" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></div>
			</td>
		</tr>

		<tr bgcolor="#000044">
			<td>
				<div><a style="width: 160; padding: 4" href="#" onMouseOver="statbar1('<%=Elite3Stat%>', '<%=Elite3Desc%>')"><%=Elite3Name%>:</a></div>
			</td>
			<td>
				<div><%=AElite3M%></div>
			</td>
			<td><div><%=TElite3M%></div>
			</td>
			<td><div><%=MElite3M%></div>
			</td>
			<td>
				<div><input type="text" size=10 value="0" name="Elite3" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></div>
			</td>
		</tr>

		<tr bgcolor="#000044">
			<td>
				<div><a style="width: 160; padding: 4" href="#" onMouseOver="statbar1('200c | 1infantry | 6hours', 'Carries 200 infantry class units')">Transports:</a></div>
			</td>
			<td>
				<div><%=ATransportsM%></div>
			</td>
			<td>
				<div><%=TTransportsM%></div>
			</td>
			<td>
				<div><%=MTransportsM%></div>
			</td>
			<td>
				<div><input type="text" size=10 value="0" name="Transports" class="numbar" ONFOCUS ="LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR="LightOut(this);"></div>
			</td>
		</tr>

		<tr bgcolor="#000066">
			<td colspan=4>
					<a href="militaryprofile.asp" class="menu_main">&nbsp;&nbsp;Military Profile&nbsp;&nbsp;</a>*
					<a href="build" target="new" class="menu_main">&nbsp;&nbsp;Military Guide&nbsp;&nbsp;</a>
			</td>
			<td align=right>
					<input type = "hidden" name = "arrived" value = "yes">
					<input type=SUBMIT value="Order" class = "button" onMouseOver="ButtonLight(this);" onMouseOut="ButtonDark(this);">
			</td>
		</tr>

	</table>

</form>

</center>
<%
Timer2 = Timer()
TotalTime = Timer2 - Timer1

Response.Write "<p class = 'small'>This took "&FormatNumber(TotalTime, 5)&" seconds</p"
%>

</body>

</html>
<%
Conn.Close
Set RSTM = nothing
Set RSDMM = nothing
Set RSGSM = nothing
Set Conn = nothing
%>
