<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="include.asp"-->
<% Response.Buffer = True
Response.Expires = -1 %>


<%
Dim Timer1, Timer2, TotalTime

Timer1 = Timer()

Dim UserNameMM
UserNameMM = Request.Cookies("NightFire")("UserName")

Dim ConnMM

Set ConnMM = Server.CreateObject("ADODB.Connection")

ConnMM.ConnectionString="Provider=SQLOLEDB;User ID=sa;Password=Kt6261;Initial Catalog=warpspire"

ConnMM.Open 

Dim RSGSMM, RSGIMM, RSDMMM, RSSMM

Set RSGSMM = Server.CreateObject("ADODB.Recordset")
Set RSSMM = Server.CreateObject("ADODB.Recordset")
Set RSGIMM = Server.CreateObject("ADODB.Recordset")
Set RSDMMM = Server.CreateObject("ADODB.Recordset")

RSGSMM.Open "SELECT * FROM GENERALSTATS WHERE UserName = '"&UserNameMM&"'", ConnMM, adOpenForwardOnly, adLockPessimistic
RSGIMM.Open "SELECT * FROM GENERALINFO WHERE UserName = '"&UserNameMM&"'", ConnMM, adOpenForwardOnly, adLockPessimistic
RSDMMM.Open "SELECT * FROM DONEMILITARY WHERE UserName = '"&UserNameMM&"'", ConnMM, adOpenForwardOnly, adLockPessimistic

'If RSGSMM.EOF = TRUE Then
'Response.Redirect "login.html"
'Else
'Response.Cookies("NightFire").Expires = DateAdd("h", 2, Now())
'End If

Dim PopulationMM, CivillianMM, MilitaryMM, LandMM, NetworthMM, CapitalMM, ElementXMM, EnergyMM, ManaMM, PlutoniumMM, AntimatterMM

PopulationMM = RSGSMM("Population")
CivillianMM = RSGSMM("Civillian")
MilitaryMM = RSGSMM("Military")
LandMM = RSGSMM("Land")
NetworthMM = RSGSMM("Networth")

RSGSMM.Close

Dim KingMM, ProtectionMM, NationMM, MinisterMM, Race, SystemTypeMM, SystemNumberMM, HealthMM, LastUpdateMM

NationMM = RSGIMM("Nation")
MinisterMM = RSGIMM("Minister")
Race = RSGIMM("Race")
SystemTypeMM = RSGIMM("SystemType")
SystemNumberMM = RSGIMM("SystemNumber")
HealthMM = RSGIMM("Health")
LastUpdateMM = RSGIMM("LastUpdate")
ProtectionMM = RSGIMM("Protection")
KingMM = RSGIMM("King")
If KingMM = "Yes" Then
	MinisterMM = MinisterMM & " *ESR*"
End If

RSGIMM.Close

RSSMM.Open "SELECT * FROM SYSTEMS WHERE SystemNumber = "&SystemNumberMM, Conn
Dim SystemNameMM
SystemNameMM = RSSMM("SystemName")
RSSMM.Close

Dim InfantryMM, Elite1MM, Elite2MM, Elite3MM, TransportsMM, MagesMM, ProbesMM, AgentsMM, MoraleMM, BarracksSlotsMM, FactorySlotsMM, UsedBarracksSlotsMM, UsedFactorySlotsMM

InfantryMM = RSDMMM("Soldiers")
Elite1MM = RSDMMM("Elite1")
Elite2MM = RSDMMM("Elite2")
Elite3MM = RSDMMM("Elite3")
TransportsMM = RSDMMM("Transports")
MagesMM = RSDMMM("Mages")
ProbesMM = RSDMMM("Probes")
AgentsMM = RSDMMM("Agents")
MoraleMM = RSDMMM("Morale")
BarracksSlotsMM = RSDMMM("BarracksSlots")
FactorySlotsMM = RSDMMM("FactorySlots")
UsedBarracksSlotsMM = RSDMMM("UsedBarracksSlots")
UsedFactorySlotsMM = RSDMMM("UsedFactorySlots")

RSDMMM.Close

%>
<%


Dim Elite1Defense, Elite1Offense, Elite2Defense, Elite2Offense, Elite3Defense, Elite3Offense, Elite1NameMM, Elite2NameMM, Elite3NameMM, IncomeMM, ElementXIncomeMM, EnergyIncomeMM, ManaIncomeMM, PlutoniumIncomeMM, AntimatterIncomeMM, DPAMM, OPAMM, OffenceMM, DefenceMM, ModsOMM, ModsDMM
Dim RSStats
Set RSStats = Server.CreateObject("ADODB.Recordset")
RSStats.Open "SELECT Type, Race, Unit, Name, Offense, Defense FROM COSTS WHERE Type = 'Military' AND Race = '"&Race&"'", ConnMM, adOpenStatic
RSStats.Find "Unit = 'Elite1'"
Elite1NameMM = RSStats("Name")
Elite1Offense = RSStats("Offense")
Elite1Defense = RSStats("Defense")
RSStats.Find "Unit = 'Elite2'"
Elite2NameMM = RSStats("Name")
Elite2Offense = RSStats("Offense")
Elite2Defense = RSStats("Defense")
RSStats.Find "Unit = 'Elite3'"
Elite3NameMM = RSStats("Name")
Elite3Offense = RSStats("Offense")
Elite3Defense = RSStats("Defense")
RSStats.Close

IncomeMM = 10000
ElementXIncomeMM = 10000
EnergyIncomeMM = 10000
ManaIncomeMM = 10000
PlutoniumIncomeMM = 10000
AntimatterIncomeMM = 10000
ModsOMM = 0
ModsDMM = 0
DefenceMM = (Elite1MM * Elite1Defense) + (Elite2MM * Elite2Defense) + (Elite3MM * Elite3Defense) + InfantryMM
DefenceMM = (DefenceMM + (DefenceMM * ModsDMM) )\1
OffenceMM = (Elite1MM * Elite1Offense) + (Elite2MM * Elite2Offense) + (Elite3MM * Elite3Offense) + InfantryMM
OffenceMM = (OffenceMM + (OffenceMM * ModsOMM) )\1
DPAMM = (DefenceMM/LandMM)
OPAMM = (OffenceMM/LandMM)

If  DateDiff("h", LastUpdateMM, Now()) >0 Then 
Response.Redirect "update.asp?page=main"
End If
%>
