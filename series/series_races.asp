<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lSeriesID
Dim sSeriesName
Dim iYear
Dim SeriesEvents()

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = "0"
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"
If CLng(lSeriesID) < 0 Then Response.Redirect "http://www.google.com"
	
iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If Not IsNumeric(iYear) Then Response.Redirect "http://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
sSeriesName = Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing

i = 0
ReDim SeriesEvents(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT se.EventID, se.EventName, e.Logo FROM SeriesEvents se INNER JOIN Events e ON se.EventID = e.EventID "
sql = sql & "WHERE se.SeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    SeriesEvents(0, i) = rs(0).Value
	SeriesEvents(1, i) = Replace(rs(1).Value, "''", "'")
    SeriesEvents(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve SeriesEvents(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE (Gopher State Events) Series Races</title>
<meta name="description" content="Gopher State Events (GSE) series races.">
<!--#include file = "../includes/js.asp" -->

<style type="text/css">
    td{
        padding: 2px 0 2px 5px;
    }
</style>
</head>

<body>
<div class="container">
    <img class="img-responsive" src="/graphics/html_header.png" alt="Gopher State Events" style="margin-top: 15px;">
    <h4 class="h4">Series Races: <%=sSeriesName%></h4>

	<table class="table">
        <tr>
		    <%For i = 0 To UBound(SeriesEvents, 2) - 1%>
				<td valign="top">
                    <a href="javascript:pop('/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>',800,600)" rel="nofollow">
                    <img class="img-responsive" src="/events/logos/<%=SeriesEvents(2, i)%>" alt="<%=SeriesEvents(1, i)%>">
                    </a>
                </td>
		    <%Next%>
        </tr>
	</table>
 </div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>