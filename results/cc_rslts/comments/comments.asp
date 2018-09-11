<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i
Dim lRaceID, lEventID
Dim sRaceName, sEventName
Dim dEventDate

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_comment") = "submit_comment" Then
End If

'log this user if they are just entering the site
If Session("access_results") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'fitness_results')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(1).Value
Set rs = Nothing

%>
<!--#include file = "../../../includes/clean_input.asp" -->
<%
%>
<!DOCTYPE html>
<html>
<head>
<title>Gopher State Events Results Comments Page</title>
<!--#include file = "../../../includes/meta2.asp" -->
</head>

<body style="background-none;background-color: #fff;">
<div style="margin: 5px;padding: 5px;background-color: #fff;font-size:0.8em;text-align: left;">
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>