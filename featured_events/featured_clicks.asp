<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lFeaturedEventsID
Dim sWebURL, sClickPage, sIPAddress

lFeaturedEventsID = Request.QueryString("featured_events_id")
sClickPage = Request.QueryString("click_page")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT WebURL, Clicks FROM FeaturedEvents WHERE FeaturedEventsID = " & lFeaturedEventsID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sWebURL = rs(0).Value
    rs(1).Value = CInt(rs(1).Value) + 1
    rs.Update
End If
rs.Close
Set rs = Nothing

sql = "INSERT INTO ClickSource (FeaturedEventsID, ClickPage, WhenClick, IPAddress) VALUES (" & lFeaturedEventsID & ", '" & sClickPage & "', '"
sql = sql & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
Set rs = conn.Execute(sql)
set rs = Nothing

sWebURL = Replace(sWebURL, "http://", "")
sWebURL = "http://" & sWebURL

conn.Close
Set conn = Nothing

Response.Redirect sWebURL
%>
<!DOCTYPE html>
<html>
<head>
<title>GSE&copy; Featured Clicks </title>
<meta name="description" content="Gopher State Events featured events clicks utility.">
</head>

<body>
&nbsp;
</body>
</html>
