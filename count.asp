<!-- #include file="data.asp" -->
<%
Dim counter, id, refer, data, total_pages, total_sessions, text_day, day_name, day_counter, text_month, month_name, month_counter, month_hits, days, netscape, explorer, today_pages, today_sessions, today_date, date_now, edit
set data = server.createobject("ADODB.recordset")
	data.activeconnection = strDB
	data.source = "SELECT * FROM counter_total WHERE ID = 1"
	data.open
' Totals for Pages & Sessions
total_pages = data("pages")
total_sessions = data("sessions")
' Totals for Days
text_day = weekday(date)
day_name = weekdayName(text_day)
day_counter = data(day_name)
' Totals for Months
text_month = month(date)
month_name = monthname(text_month)
month_hits = data(month_name)
' Total Days
days = data("days")
' Browsers
netscape = data("netscape")
explorer = data("explorer")
data.close

data.source = "SELECT * FROM counter_today WHERE id = 1"
data.open
' Get Todays Totals
today_pages = data("pages")
today_sessions = data("sessions")
today_date = data("date")
today_date = cdate(today_date)
data.close

set data = nothing

' If New Day Change Todays Stats to Yesterdays and Reset Todays Stats
date_now = date
if today_date <> date_now then
	days = days + 1
	set edit = server.createobject("ADODB.Command")
	edit.activeconnection = strDB
	edit.commandtext = "UPDATE counter_yesterday SET counter_yesterday.pages = " & today_pages & " WHERE counter_yesterday.id = 1"
	edit.execute
	edit.commandtext = "UPDATE counter_yesterday SET counter_yesterday.sessions = " & today_sessions & " WHERE counter_yesterday.id = 1"
	edit.execute
	edit.commandtext = "UPDATE counter_yesterday SET counter_yesterday.date = '" & today_date & "' WHERE counter_yesterday.id = 1"
	edit.execute
	edit.commandtext = "UPDATE counter_today SET counter_today.date = '" & date & "' WHERE counter_today.id = 1"
	edit.execute
	edit.commandtext = "UPDATE counter_today SET counter_today.sessions = 0 WHERE counter_today.id = 1"
	edit.execute
	edit.commandtext = "UPDATE counter_today SET counter_today.pages = 0 WHERE counter_today.id = 1"
	edit.execute
	edit.commandtext = "UPDATE counter_total SET counter_total.days = " & days & " WHERE counter_total.id = 1"
	edit.execute
	edit.activeconnection.close
	set edit = nothing

	today_pages = 0
	today_sessions = 0			   
end if

' If New Session Add Stats for it
if session("count") = "" then
	set edit = server.createobject("ADODB.Command")
	edit.activeconnection = strDB
	' Find out who refered the user to this site
	refer = request.servervariables("HTTP_REFERER")
	if refer <> "" then
		set data = server.createobject("ADODB.RecordSet")
			data.activeconnection = strDB
			data.source = "SELECT * FROM counter_refer WHERE refer = '" & refer & "'"
			data.open
			' If the refering page is new, create the record
			if data.eof then
				data.close
				set data = nothing
				
				edit.commandtext = "INSERT INTO counter_refer(refer, hits) VALUES('" & refer & "', 1)"
				edit.execute
			else
				counter = data("hits") + 1
				id = data("id")
				data.close
				set data = nothing
				
				edit.commandtext = "UPDATE counter_refer SET counter_refer.hits = " & counter & " WHERE id = " & id
				edit.execute		
			end if
	end if
	
	' Update Todays Stats
	today_sessions = today_sessions + 1
	today_pages = today_pages + 1
	edit.commandtext = "UPDATE counter_today SET counter_today.sessions = " & today_sessions & " WHERE id = 1"
	edit.execute		
	edit.commandtext = "UPDATE counter_today SET counter_today.pages = " & today_pages & " WHERE id = 1"
	edit.execute		

	' Update Total Stats
	total_sessions = total_sessions + 1
	total_pages = total_pages + 1
	month_hits = month_hits + 1
	day_counter = day_counter + 1
	edit.commandtext = "UPDATE counter_total SET counter_total.sessions = " & total_sessions & " WHERE id = 1"
	edit.execute		
	edit.commandtext = "UPDATE counter_total SET counter_total.pages = " & total_pages & " WHERE id = 1"
	edit.execute		
	edit.commandtext = "UPDATE counter_total SET counter_total." & month_name & " = " & month_hits & " WHERE id = 1"
	edit.execute		
	edit.commandtext = "UPDATE counter_total SET counter_total." & day_name & " = " & day_counter & " WHERE id = 1"
	edit.execute

	' Update Browser stats
	if InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Netscape") > 0 then
		netscape = netscape + 1
 		edit.commandtext = "UPDATE counter_total SET counter_total.netscape = " & netscape & " WHERE id = 1"
		edit.execute		
	else 
		explorer = explorer + 1
 		edit.commandtext = "UPDATE counter_total SET counter_total.explorer = " & explorer & " WHERE id = 1"
		edit.execute		
	end if
	session("count") = "yes"
else
	'If not a new session just update the page counts
	total_pages = total_pages + 1
	today_pages = today_pages + 1
	set edit = server.createobject("ADODB.Command")
	edit.activeconnection = strDB
	edit.commandtext = "UPDATE counter_total SET counter_total.pages = " & total_pages & " WHERE id = 1"
	edit.execute
	edit.commandtext = "UPDATE counter_today SET counter_today.pages = " & today_pages & " WHERE id = 1"
	edit.execute		
end if
edit.activeconnection.close
set edit = nothing
%>

