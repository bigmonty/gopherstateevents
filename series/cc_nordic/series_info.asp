<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, conn2, rs, sql, sql2, rs2
Dim i, j
Dim lSeriesID
Dim sSeriesName

Dim Series(), SeriesMeets()

lSeriesID = Request.QueryString("series_id")
		
Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.form.Item("submit_series") = "submit_series" Then lSeriesID = Request.Form.Item("series")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CCSeriesID, SeriesName, SeriesYear FROM CCSeries ORDER BY SeriesYear DESC, SeriesName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Series(0, i) = rs(0).Value
	Series(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve Series(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

ReDim SeriesMeets(3, 0)
If Not CLng(lSeriesID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName FROM CCSeries WHERE CCSeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing

    j = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT se.MeetsID, se.MeetName, se.MeetDate, se.Location FROM CCSeriesMeets se INNER JOIN Meets e ON se.MeetsID = e.MeetsID "
    sql = sql & "WHERE se.CCSeriesID = " & lSeriesID & " ORDER BY se.MeetDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesMeets(0, j) = rs(0).Value
	    SeriesMeets(1, j) = Replace(rs(1).Value, "''", "'")
        SeriesMeets(2, j) = rs(2).Value
        SeriesMeets(3, j) = GetRaceName(rs(0).Value)
	    j = j + 1
	    ReDim Preserve SeriesMeets(3, j)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function GetRaceName(lMeetsID)
    GetRaceName = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT RaceName FROM CCSeriesRaces sr INNER JOIN CCSeriesMeets se ON sr.CCSeriesMeetsID = se.CCSeriesMeetsID WHERE se.MeetsID = " & lMeetsID
    sql2 = sql2 & " AND se.CCSeriesID = " & lSeriesID
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        GetRaceName = GetRaceName & Replace(rs2(0).Value, "_", " ") & ", "
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If Not GetRaceName = vbNullString Then 
        GetRaceName = Trim(GetRaceName)
        GetRaceName = Left(GetRaceName, Len(GetRaceName) - 1)
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE (Gopher State Events) CC/Nordic Series Manager</title>
<meta name="description" content="Gopher State Events (GSE) Cross-Country/Nordic Ski Series Information page.">
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <h3 class="h3">Gopher State Cross-Country/Nordic Ski Race Series</h3>

    <ul class="nav">
        <li class="nav-item"><a class="nav-link" href="/series/cc_nordic/series_results.asp" onclick="openThis(this.href,1024,768);return false;" 
            style="font-weight: bold;color: #f00;">Series Standings</a></li>
        <li class="nav-item"><a class="nav-link" href="javascript:pop('how_it_works.asp',600,650)">How It Works</a></li>
    </ul>

    <%If UBound(Series, 2) > 0 Then%>
        <form role="form" class="form-inline" name="select_series" method="Post" action="series_info.asp">
        <label for="series">Select Series:</label>
        <select class="form-control" name="series" id="series" onchange="this.form.submit1.click();">
            <option value="">&nbsp;</option>
            <%For i = 0 To UBound(Series, 2) - 1%>
                <%If CLng(lSeriesID) = CLng(Series(0, i)) Then%>
                    <option value="<%=Series(0, i)%>" selected><%=Series(1, i)%></option>
                <%Else%>
                    <option value="<%=Series(0, i)%>"><%=Series(1, i)%></option>
                <%End If%>
            <%Next%>
        </select>
        <input type="hidden" name="submit_series" id="submit_series" value="submit_series">
        <input class="form-control" type="submit" name="submit1" id="submit1" value="Select Series To View">
        </form>
    <%End If%>

    <%If Not CLng(lSeriesID) = 0 Then%>
        <hr>
        <h4 class="h4">Series Meets:</h4>
        <table class="table table-striped">
            <tr>
                <th>No.</th>
                <th>Meet</th>
                <th>Date</th>
                <th>Race(s)</th>
            </tr>
            <%For i = 0 To UBound(SeriesMeets, 2) - 1%>
                <tr>
                    <td style="text-align:right;"><%=i + 1%>)</td>
                    <td><a href="javascript:pop('/events/raceware_events.asp?event_id=<%=SeriesMeets(0, i)%>',800,600)" 
                        rel="nofollow"><%=SeriesMeets(1, i)%></a></td>
                    <td><%=SeriesMeets(2, i)%></td>
                    <td><%=SeriesMeets(3, i)%></td>
                </tr>
            <%Next%>
        </table>
    <%End If%>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>