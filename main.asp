<!--#Include File = "main_code.asp"-->
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
		<center>
		<!--#Include File = "ad.asp"-->
		<!--#Include File = "statbar.asp"-->

		<table cellspacing=0 cellpadding=5 width=600>
			<tr>
				<td bgcolor="#0000CC" colspan=2>
					<div align=center class="large">Status</div>
				</td>
			</tr>
			<tr bgcolor="#000044">
				<td align=left rows=400>
					<div class = "small">Nation: <%=NationMM%> &nbsp;(<%=SystemNumberMM%>) of <%=SystemNameMM%></div>
				</td>
				<td align=left rows=200>
					<div class = "small">Dictator: <%=MinisterMM%></div>
				</td>
			</tr>
			<tr>
				<td bgcolor="#000044" colspan=2>
					<table border=0 bgcolor="#000044" cellpadding=5 width=100%>
						<tr>
							<td align=right valign=top>
								<div class="small">
									Population:<br>
									Civilian:<br>
									Military:<br>
									Land:<br>
									Networth:<br>
								</div>
							</td>
							<td valign=top width=60>
								<div class="small">
									<%=FormatNumber(PopulationMM, 0)%><br>
									<%=FormatNumber(CivillianMM, 0)%><br>
									<%=FormatNumber(MilitaryMM, 0)%><br>
									<%=FormatNumber(LandMM, 0)%><br>
									<%=FormatNumber(NetworthMM, 0)%><br>
								</div>
							</td>

							<td align=right valign=top>
								<div class="small">
									Transports:<br>
									Infantry:<br>
									<%=Elite1NameMM%>:<br>
									<%=Elite2NameMM%>:<br>
									<%=Elite3NameMM%>:<br>
									Military Morale:<br>
								</div>
							</td>
							<td valign=top cols=70 >
								<div class="small">
									<%=FormatNumber(TransportsMM, 0)%><br>
									<%=FormatNumber(InfantryMM, 0)%><br>
									<%=FormatNumber(Elite1MM, 0)%><br>
									<%=FormatNumber(Elite2MM, 0)%><br>
									<%=FormatNumber(Elite3MM, 0)%><br>
									<%=MoraleMM%>%<br>
								</div>
								</td>
								<td align=right valign=top>
									<div class="small">
										Race:<br>
										System Type:<br>
										Civilian Health:<br>
										Agents:<br>
										Probes:<br>
										Mages:<br>
									</div>
								</td>
								<td align=left valign=top>
									<div class="small">
										<%=Race%><br>
										<%=SystemTypeMM%><br>
										<%=HealthMM%>%<br>
										<%=FormatNumber(AgentsMM, 0)%><br>
										<%=FormatNumber(ProbesMM, 0)%><br>
										<%=MagesMM%><br>
									</div>
								</td>
							</tr>
						</table>

					</td>
				</tr>

				<tr>
					<td bgcolor="#0000CC" colspan=2>
						<div align=center class="large">Profiles</div>
					</td>
				</tr>
				<tr>
					<td colspan=2 bgcolor="#000066">
						<div align=center>
							<a href="ecoprofile.asp" class="menu_main" style="width:80pt">Economy <br>Profile</a>*
							<a href="militaryprofile.asp" class="menu_main" style="width:80pt">Military<br>Profile</a>*
							<a href="landprofile.asp" class="menu_main" style="width:80pt">Land<br>Profile</a>*
							<a href="" class="menu_main" style="width:80pt">Magic <br>Profile</a>
							<a href="/NightFire/Guide/guide.html" class="menu_main" target="new" style="width:80pt">NightFire<br>Guide</a>*
						</div>
					</td>
				</tr>
			</table>
			<br>
			<div class = "small">
				Current Time: <%=Now()%><br>
				Time since your last update: <%=LastUpdateMM%><br>
				<%
				If DateDiff("h", Now(), ProtectionMM) > 0 Then
					Response.Write "Your nation has been granted protection from outside forces in order to prove yourself. There are " & DateDiff("h", Now(), ProtectionMM) & " hours left to prove your worthiness."
				End If
				%>
				<br>
			</div>
<!--#Include File = "main_code2.asp"-->
</center>
	</body>
</html>
