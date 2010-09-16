<!-- #include file="data2.asp" -->
<%
set data = server.createobject("ADODB.RecordSet")
data.activeconnection = strDB
data.source = "SELECT * FROM counter_total WHERE id = 1"
data.open
	
' Get Stats for Totals
total_sessions = data("sessions")
total_pages = data("pages")
monday = data("monday")
tuesday = data("tuesday")
wednesday = data("wednesday")
thursday = data("thursday")
friday = data("friday")
saturday = data("saturday")
sunday = data("sunday")
january = data("january")
febuary = data("febuary")
march = data("march")
april = data("april")
may = data("may")
june = data("june")
july = data("july")
august = data("august")
september = data("september")
october = data("october")
november = data("november")
december = data("december")
explorer = data("explorer")
netscape = data("netscape")
days = data("days")
' Do Calculations
if total_pages > 0 and total_sessions > 0 then
pages_per_total = total_pages / total_sessions
else
pages_per_total = 0
end if
pages_per_total = formatnumber(pages_per_total,2)
data.close

' Get Stats for Today
data.source = "SELECT * FROM counter_today WHERE id = 1"
data.open
today_sessions = data("sessions")
today_pages = data("pages")
' Do Calculations
if today_sessions > 0 then
pages_per_today = today_pages / today_sessions
else
pages_per_today = 0
end if
pages_per_today = formatnumber(pages_per_today,2)
data.close

' Get Stats for Yesterday
data.source = "SELECT * FROM counter_yesterday WHERE id = 1"
data.open
yesterday_sessions = data("sessions")
yesterday_pages = data("pages")
' Do Calculations
if yesterday_sessions > 0 then
pages_per_yesterday = yesterday_pages / yesterday_sessions
else
pages_per_yesterday = 0
end if
pages_per_yesterday = formatnumber(pages_per_yesterday,2)
data.close
%>
<div align=center><p><a href="stats.asp">Back to the Download page</a></p></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align=center>
	<table width="95%" border="1" cellspacing="1" cellpadding="1">
		<tr>
		    <td colspan="2" align=center bgcolor="#336699"><b>Total</b></td>
		    </tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Sessions:</b></font></td>
		    <td bgcolor="#CCCC99"><%= total_sessions %></td>
		</tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Page Views:</b></font></td>
		    <td bgcolor="#CCCC99"><%= total_pages %></td>
		</tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Pages Per Session:</b></font></td>
		    <td bgcolor="#CCCC99"><%= pages_per_total %></td>
		</tr>
	</table>
	</td>
    <td align=center>
	<table width="95%" border="1" cellspacing="1" cellpadding="1">
  		<tr>
   			<td colspan="2" align=center bgcolor="#336699"><b>Today's Stats</b></td>
 	    </tr>
  		<tr>
   			<td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Sessions:</b></font></td>
  			<td bgcolor="#CCCC99"><%= today_sessions %></td>
	    </tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Page Views:</b></font></td>
		    <td bgcolor="#CCCC99"><%= today_pages %></td>
		</tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Pages Per Session:</b></font></td>
		    <td bgcolor="#CCCC99"><%= pages_per_today %></td>
		</tr>
	</table>
	</td>
  <td align=center>
	<table width="95%" border="1" cellspacing="1" cellpadding="1">
		<tr>
		    <td colspan="2" align=center bgcolor="#336699"><b>Yesterday's Stats</b></td>
		    </tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Sessions:</b></font></td>
		    <td bgcolor="#CCCC99"><%= yesterday_sessions %></td>
		</tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Page Views:</b></font></td>
		    <td bgcolor="#CCCC99"><%= yesterday_pages %></td>
		</tr>
		<tr>
		    <td align=right width="50%" bgcolor="#CCCC99"><font size=2><b>Pages Per Session:</b></font></td>
		    <td bgcolor="#CCCC99"><%= pages_per_yesterday %></td>
		</tr>
	</table>
	</td>
  </tr>
</table>
<br><br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<table width="50%" border="0" cellspacing="1" cellpadding="1" align=center>
	  <tr>
    	<td colspan="2" align=center><font size=5>Browser Type</font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Explorer</b></font></td>
	    <td>
		<%
		pic_size = (explorer / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/blue.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= explorer %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Netscape</b></font></td>
	    <td>
		<%
		pic_size = (netscape / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/red.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= netscape %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
	</table>
	<br><br>
	</td>
	</tr>
  <tr>
    <td>
	<table width="50%" border="0" cellspacing="1" cellpadding="1" align=center>
	  <tr>
    	<td colspan="2" align=center><font size=5>Sessions Per Day</font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Monday</b></font></td>
	    <td>
		<%
		pic_size = (monday / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/red.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= monday %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Tuesday</b></font></td>
	    <td>
		<%
		pic_size = (tuesday / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/green.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= tuesday %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Wednesday</b></font></td>
		<td>
		<%
		pic_size = (wednesday / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/blue.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= wednesday %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Thursday</b></font></td>
		<td>
		<%
		pic_size = (thursday / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/yellow.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= thursday %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Friday</b></font></td>
		<td>
		<%
		pic_size = (friday / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/orange.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= friday %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Saturday</b></font></td>
		<td>
		<%
		pic_size = (saturday / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/red.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= saturday %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Sunday</b></font></td>
		<td>
		<%
		pic_size = (sunday / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/blue.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= sunday %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
   	  <tr>
    	  <td width="20%" align=right><font size="2"><b>Total</b></font></td>
	    <td><img src="../images/green.gif" width="90%" height=5> <font size=2><b><%= total_sessions %></b></font>
		</td>
	  </tr>
	</table>
	<br><br>
	</td>
	</tr>
	<tr>
    <td>
	<table width="50%" border="0" cellspacing="1" cellpadding="1" align=center>
	  <tr>
    	<td colspan="2" align=center><font size=5>Sessions Per Month</font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>January</b></font></td>
	    <td>
		<%
		pic_size = (january / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/red.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= january %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>Febuary</b></font></td>
	    <td>
		<%
		pic_size = (febuary / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/green.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= febuary %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>March</b></font></td>
		<td>
		<%
		pic_size = (march / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/blue.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= march %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>April</b></font></td>
		<td>
		<%
		pic_size = (april / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/yellow.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= april %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>May</b></font></td>
		<td>
		<%
		pic_size = (may / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/orange.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= may %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>June</b></font></td>
		<td>
		<%
		pic_size = (june / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/red.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= june %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
  	  <tr>
    	<td width="20%" align=right><font size="2"><b>July</b></font></td>
		<td>
		<%
		pic_size = (july / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/blue.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= july %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
	  <tr>
    	<td width="20%" align=right><font size="2"><b>August</b></font></td>
		<td>
		<%
		pic_size = (august / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/green.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= august %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
	  <tr>
    	<td width="20%" align=right><font size="2"><b>September</b></font></td>
		<td>
		<%
		pic_size = (september / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/yellow.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= september %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
	  <tr>
    	<td width="20%" align=right><font size="2"><b>October</b></font></td>
		<td>
		<%
		pic_size = (october / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/orange.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= october %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
	  <tr>
    	<td width="20%" align=right><font size="2"><b>November</b></font></td>
		<td>
		<%
		pic_size = (november / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/red.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= november %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
	  <tr>
    	<td width="20%" align=right><font size="2"><b>December</b></font></td>
		<td>
		<%
		pic_size = (december / total_sessions) * 100
		pic_size = formatnumber(pic_size,2)
		%>
		<img src="../images/blue.gif" height=5 width="<%= pic_size %>%"><font size=2> <b><%= december %></b> Sessions or <b><%= pic_size %>%</b></font></td>
	  </tr>
   	  <tr>
    	  <td width="20%" align=right><font size="2"><b>Total</b></font></td>
	    <td><img src="../images/green.gif" width="90%" height=5> <font size=2><b><%= total_sessions %></b></font></td>
	  </tr>
	</table>		
	</td>
  </tr>
</table>
<br><br>
<%
data.source = "SELECT * FROM counter_refer ORDER BY hits DESC"
data.open
response.write "<table width=""50%"" border=""1"" cellspacing=""1"" cellpadding=""1"" align=center>"
response.write "<tr><td colspan=""2"" align=center bgcolor=""#336699""><font size=5 color=""#000000"">Top 20 Refering Links</font></td></tr>"
response.write "<tr bgcolor=""#CCCC99""><td align=center width=""85%""><b>Link</b></td><td align=center><b>Clicks</b></td></tr>"
for i = 0 to 20
	if data.eof then
		'do nothing
	else
	response.write "<tr>"
	response.write "<td align=right width=""85%"" bgcolor=""#CCCC99""><a href=""" & data("refer") & """ target=""_blank"">" & data("refer") & "</a></td><td align=center bgcolor=""#CCCC99"">" & data("hits") & "</td>"
	response.write "</tr>"
	data.movenext
	end if
next
response.write "</table>"
data.close
set data = nothing
%>
<br><br>
<table width="50%" border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td width=50% align=right><img src="../images/primal.gif"></td>
    <td><font size="2">&nbsp;© Copyright Primal Blue</font></td>
  </tr>
</table>
<!-- #include file="count.asp" -->