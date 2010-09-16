
<%
Dim Browser, Version, Platform, Name
Set Browser = Server.CreateObject("MSWC.BrowserType")
Version = Browser.Version
Platform = Browser.Platform
Name = Browser.Browser

Dim IP, Path
IP = Request.ServerVariables("REMOTE_ADDR")
Path = Request.ServerVariables("URL")


Dim ConnCC
Set ConnCC = Server.CreateObject("ADODB.Connection")
ConnCC.ConnectionString="DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & Server.MapPath("\NightFire\db\mydb.mdb")
ConnCC.Open

Dim RSCC, PageNameCC
Path = Replace(Path, "/nightfire/", "")
Select Case Path
	Case "index.asp"
		PageNameCC = "FrontPage"
	Case Else
		PageNameCC = "GamePages"
End Select


Set RSCC = Server.CreateObject("ADODB.RecordSet")

RSCC.Open "SELECT * FROM PAGECOUNT WHERE Page = '"&PageNameCC&"'", ConnCC, 3, 3

Dim HitsCC

If RSCC.EOF = FALSE Then
	HitsCC = RSCC.RecordCount + 1
Else
	HitsCC = 1
End If

RSCC.AddNew
RSCC("Page") = PageNameCC
RSCC("TimeOfHit") = Now()
RSCC("Browser") = Name 
RSCC("IP") = IP
RSCC("Platform") = Platform
RSCC.Update
RSCC.Close
If PageNameCC = "FrontPage" Then
	Response.Write "<p class = 'normal' align = 'center'>There have been "&HitsCC&" hits to this page since <i>6/7/01</i></p>"
End If
ConnCC.Close
%>

