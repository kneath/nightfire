<%@ Language=VBScript %>
<% Option Explicit %>

<%
Response.Cookies("NightFire")("UserName") = ""
Response.Cookies("NightFire").Expires = Date - 365

Response.Redirect "index.asp"
%>