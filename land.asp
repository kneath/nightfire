<!--#Include File = "land_code.asp"-->
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
<body bgcolor="#000000" text="#ffffff" link="#ffffff" vlink="#ffffff" alink="ffcc33">
<SCRIPT SRC="fieldcheck.js"></SCRIPT>


<!--#Include File = "ad.asp"-->
<!--#Include File = "statbar.asp"-->

<%=MessageL%>
<%=MessageEL%>

<table cellspacing=0 cellpadding=5  width="100%">

<tr><td bgcolor="#0000CC" colspan=8><div align=center class="large">Land</div></td></tr>

<form action = "land.asp" method = "post" name = "statbar1">

</table><table cellspacing=0 cellpadding=2  width = "100%"><tr><td bgcolor="#000066" colspan=8><div align=center><b>Cost:</b> 

<input type=text size=40 value="put mouse over building to see cost" class = "statbar" name = "statbar1"></div></td></form>

</tr>

<tr>
<td bgcolor="#000066" colspan=8 align="center">
Plain Land: <span class = 'normal'> <%=TLandL%> (<%=FormatNumber(TLandL*100/TotLandL, 1)%>%)</span> &nbsp; &nbsp; Destroyed Land: <span class = 'normal'><%=RuggedLandL%>(<%=FormatNumber(RuggedLandL*100/TotLandL, 1)%>%)</span>
</td>
</tr>
<form action = "land.asp" method = "post" name = "land">
<tr bgcolor="#000044">

<td><div>Building</div></td>

<td><div>Built</div></td>

<td><div>Max</div></td>

<td><div>Order</div></td>

<td><div>Building</div></td>

<td><div>Built</div></td>

<td><div>Max</div></td>

<td><div>Order</div></td>

</tr>

<tr bgcolor="#000044">

<td><div><a href="#" style="width: 160; padding: 4" onMouseOver="document.forms[0].elements[0].value='1500c | 150t | 1km';">Barracks:

<br><font class = "small">(<%=FormatNumber(BBarracksL, 0)%> building, <%=FormatNumber(DBarracksL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DBarracksL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MBarracksL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="Barracks" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='1300c | 140t | 1km';">Defense Stations:

<br><font class = "small">(<%=FormatNumber(BDefenseStationsL, 0)%> building, <%=FormatNumber(DDefenseStationsL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DDefenseStationsL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MDefenseStationsL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="DefenseStations" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='2000c | 200t | 1km';">Factories:

<br><font class = "small">(<%=FormatNumber(BFactoriesL, 0)%> building, <%=FormatNumber(DFactoriesL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DFactoriesL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MFactoriesL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="Factories" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='1200c | 120t | 1km';">Atomic Plants:

<br><font class = "small">(<%=FormatNumber(BAtomicPlantsL, 0)%> building, <%=FormatNumber(DAtomicPlantsL*100/TLandL, 2)%>%)</a></div></td>

<td><div class = "small"><%=FormatNumber(DAtomicPlantsL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MAtomicPlantsL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="AtomicPlants" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='300c | 20t | 1km';">Solar Panels:

<br><font class = "small">(<%=FormatNumber(BSolarPanelsL, 0)%> building, <%=FormatNumber(DSolarPanelsL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DSolarPanelsL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MSolarPanelsL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="SolarPanels" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='No Tech';">Antimatter Plants:

<br><font class = "small">(<%=FormatNumber(BAntimatterPlantsL, 0)%> building, <%=FormatNumber(DAntimatterPlantsL*100/TLandL, 2)%>%)</a></div></td>

<td><div class = "small"><%=FormatNumber(DAntimatterPlantsL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MAntimatterPlantsL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="AntimatterPlants" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='1000c | 50t | 1km';">Consulates:

<br><font class = "small">(<%=FormatNumber(BConsulatesL, 0)%> building, <%=FormatNumber(DConsulatesL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DConsulatesL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MConsulatesL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="Consulates" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='1300c | 100t | 1km';">Plutonium Mines:
<br><font class = "small">(<%=FormatNumber(BPlutoniumPlantsL, 0)%> building, <%=FormatNumber(DPlutoniumPlantsL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DPlutoniumPlantsL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MPlutoniumPlantsL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="PlutoniumMines" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='200,000c | 100t | 1km | +4 hour build time';">Research Centers:

<br><font class = "small">(<%=FormatNumber(BReasearchCentersL, 0)%> building, <%=FormatNumber(DReasearchCentersL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumbeR(DReasearchCentersL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MReasearchCentersL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="ResearchCenters" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='No Tech';">Antimatter Labs:
<br><font class = "small">(<%=FormatNumber(BAntimatterLabsL, 0)%> building, <%=FormatNumber(DAntimatterLabsL*100/TLandL, 2)%>%)</a></div></td>

<td><div class = "small"><%=FormatNumber(DAntimatterLabsL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MAntimatterLabsL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="AntimatterLabs" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='1000c | 100t | 1km';">Element-X Mines:
<br><font class = "small">(<%=FormatNumber(BElementXMinesL, 0)%> building, <%=FormatNumber(DElementXMinesL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DElementXMinesL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MElementXMinesL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="ElementXMines" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='No Tech';">Nuclear Silos:
<br><font class = "small">(<%=FormatNumber(BNuclearSilosL, 0)%> building, <%=FormatNumber(DNuclearSilosL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DNuclearSilosL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MNuclearSilosL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="NuclearSilos" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

</tr>



<tr bgcolor="#000044">

<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='Free | Free | 1km | House your population';">Cities:
<br><font class = "small">(<%=FormatNumber(BCitiesL, 0)%> building, <%=FormatNumber(DCitiesL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DCitiesL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(FreePlainL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="Cities" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>
<td><div><a style="width: 160; padding: 4" href="#" onMouseOver="document.forms[0].elements[0].value='1000c | 10t | 1km';">Psychic Centers:
<br><font class = "small">(<%=FormatNumber(BTemplesL, 0)%> building, <%=FormatNumber(DTemplesL*100/TLandL, 2)%>%)</div></a></td>

<td><div class = "small"><%=FormatNumber(DTemplesL, 0)%></div></td>

<td><div class = "small"><%=FormatNumber(MTemplesL, 0)%></div></td></font>

<td><div><input type="text" size=4 maxlength=10 value="0" name="Temples" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"></div></td>

</tr>


<tr bgcolor="#000044">

<td colspan=8><div align=center>
<%
Dim BuiltLandL
BuiltLandL = TLandL - FreePlainL
%>
Cities Allotocated: <span class="normal"><%=FormatNumber(AllotCitiesL, 0)%></span>
Total Plain Land: <span class="normal"><%=FormatNumber(TLandL, 0)%></span>&nbsp;&nbsp;
Free Plain Land: <span class="normal"><%=FormatNumber(FreePlainL, 0)%></span>&nbsp;&nbsp;
Built Land: <span class="normal"><%=FormatNumber(BuiltLandL, 0)%></span>


</div></td></tr>
<tr bgcolor="#000066"><td colspan=8><div align="center">Incoming Land: <span class="normal"><%=FormatNumber(BuildingLandL, 0)%></span>&nbsp;&nbsp;
<input type=submit value="Explore" class = "button" onMouseOver="ButtonLight(this);" onMouseOut="ButtonDark(this);">&nbsp;&nbsp;
<input type="text" size=10 value="0" name="ExploreAcres" class = "numbar" ONFOCUS = "LightUp(this);" ONCHANGE="checkField(this.value, this);" ONBLUR = "LightOut(this);"> &nbsp;&nbsp;
<input type=hidden name = "LandType" value = "Plain">&nbsp;&nbsp;
Max: <span class="normal"><%=MLandL%></span></div></td></tr>

<tr bgcolor="#000066">

<td colspan=4>

<a href="landprofile.asp" class="menu_main">Land Profile</a>*

<a href="/nightfire/guide/buildings.html" class="menu_main" target="new">Land Guide</a>*

</td><td colspan=4 align=right>

<INPUT TYPE=RADIO NAME=WhatDo VALUE="Build" class = "button" checked>Build&nbsp;&nbsp;&nbsp;
<INPUT TYPE=RADIO NAME=WhatDo VALUE="Raze" class = "button">Raze&nbsp;&nbsp;
<Input type = "hidden" value = "yes" name = "Arrived">
<INPUT TYPE=SUBMIT value="Order" class = "button" onMouseOver="ButtonLight(this);" onMouseOut="ButtonDark(this);">

</td></tr>

</form>

</tr>
</table>
</form>
</center>


</body>

</html>
<%

%>

