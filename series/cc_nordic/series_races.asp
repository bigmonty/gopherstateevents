<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lSeriesID
Dim sSeriesName
Dim iYear
Dim SeriesMeets()

lSeriesID = Request.QueryString("series_id")
		
iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If Not IsNumeric(iYear) Then Response.Redirect "http://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesName FROM CCSeries WHERE CCSeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
sSeriesName = Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing

i = 0
ReDim SeriesMeets(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT se.MeetsID, se.MeetName FROM CCSeriesMeets se INNER JOIN Meets e ON se.MeetsID = e.MeetsID "
sql = sql & "WHERE se.CCSeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    SeriesMeets(0, i) = rs(0).Value
	SeriesMeets(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve SeriesMeets(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE (Gopher State Events) CC/Nordic Series Races</title>
<meta name="description" content="Gopher State Events (GSE) Cross-Country/Nordic Ski Series RAces.">
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="row">
        <h1 class="h1">CC/NNordic Series Races: <%=sSeriesName%></h1>

	    <table class="table">
            <tr>
		        <%For i = 0 To UBound(SeriesMeets, 2) - 1%>
				    <td valign="top">
                        <a href="javascript:pop('/events/raceware_events.asp?event_id=<%=SeriesMeets(0, i)%>',800,600)" rel="nofollow">
                        <img src="/events/logos/<%=SeriesMeets(2, i)%>" alt="<%=SeriesMeets(1, i)%>" style="width: <%=770/UBound(SeriesMeets, 2) - 1%>px;">
                        </a>
                    </td>
		        <%Next%>
            </tr>
	    </table>
    </div>
 </div>
 <!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>