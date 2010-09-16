<%


Dim UserNameS, ConnS
Set ConnS = Server.CreateObject("ADODB.Connection")
ConnS.ConnectionString="PROVIDER=MSDASQL;DRIVER={SQL Server};SERVER=127.0.0.1;DATABASE=warpspire;UID=sa;PWD=Kt6261;"
ConnS.Open

UserNameS = Request.Cookies("NightFire")("UserName")


Dim RSS

Set RSS = Server.CreateObject("ADODB.Recordset")

RSS.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&UserNameS&"'", ConnS, adOpenDynamic, adLockOptimistic


NetworthS = NetworthS + RSS("Networth")
CapitalS = CapitalS + RSS("Capital")
ElementXS = ElementXS + RSS("ElementX")
EnergyS = EnergyS + RSS("Energy")
ManaS = ManaS + RSS("Mana")
PlutoniumS = RSS("Plutonium")
AntimatterS = RSS("Antimatter")

RSS.Close

%>
<center>
<table cellspacing=0 cellpadding=3><tr>

<td bgcolor="#000033" align=right><div class="small">Capital [c]</div></td>

<td bgcolor="#000066"><div class="small_orange"><%=FormatNumber(CapitalS, 0)%></div></td>



<td bgcolor="#000033" align=right><div class="small">Element-X [t]</div></td>

<td bgcolor="#000066"><div class="small_orange"><%=FormatNumber(ElementXS, 0)%></div></td>



<td bgcolor="#000033" align=right><div class="small">Energy [j]</div></td>

<td bgcolor="#000066"><div class="small_orange"><%=FormatNumber(EnergyS, 0)%></div></td></tr>



<tr><td bgcolor="#000033" align=right><div class="small">Psi [s]</div></td>

<td bgcolor="#000066"><div class="small_orange"><%=FormatNumber(ManaS, 0)%></div></td>



<td bgcolor="#000033" align=right><div class="small">Plutonium [p]</div></td>

<td bgcolor="#000066"><div class="small_orange"><%=FormatNumber(PlutoniumS, 0)%></div></td>



<td bgcolor="#000033" align=right><div class="small">Antimatter [a]</div></td>

<td bgcolor="#000066"><div class="small_orange"><%=FormatNumber(AntimatterS, 0)%></div></td>



</tr></table><br>
</center>