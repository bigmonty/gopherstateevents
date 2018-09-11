<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lSeriesID
Dim sSeriesName
Dim Series(), SeriesParts, SeriesRaces()

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

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

If Request.Form.item("submit_series") = "submit_series" Then
    lSeriesID = Request.Form.Item("series")
    If CLng(lSeriesID) = vbNullString Then lSeriesID = 0
End If

If Not CLng(lSeriesID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName FROM CCSeries WHERE CCSeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
 
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sp.CCSeriesPartsID, sp.PartName, sp.School, sr.Score FROM CCSeriesParts sp INNER JOIN CCSeriesResults sr "
    sql = sql & "ON sp.CCSeriesPartsID = sr.CCSeriesPartsID WHERE sp.CCSeriesID = " & lSeriesID & " ORDER BY sp.PartName"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        SeriesParts = rs.GetRows()
    Else
        ReDim SeriesParts(3, 0)
    End If
    rs.Close
    Set rs = Nothing

    'take out lowest score
'    For i = 0 To UBound(SeriesParts, 2)
'        Set rs = Server.CreateOBject("ADODB.Recordset")
'        sql = "SELECT ss.Points FROM CCSeriesParts sp INNER JOIN CCSeriesStdgs ss ON sp.CCSeriesPartsID = ss.CCSeriesPartsID WHERE sp.CCSeriesID = " 
'        sql = sql & lSeriesID & " AND ss.CCSeriesPartsID = " & SeriesParts(0, i) & " ORDER BY ss.Points"
'        rs.Open sql, conn, 1, 2
'        If rs.RecordCount > 0 Then SeriesParts(3, i) = CSng(SeriesParts(3, i)) - CSng(rs(0).Value)
'        rs.Close
'        Set rs = Nothing
'    Next

    j = 0
    ReDim SeriesRaces(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sr.RacesID, sm.MeetName, sm.MeetDate FROM CCSeriesRaces sr INNER JOIN CCSeriesMeets sm ON sr.CCSeriesMeetsID = sm.CCSeriesMeetsID "
    sql = sql & "WHERE sm.CCSeriesID = " & lSeriesID & " ORDER BY sm.MeetDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesRaces(0, j) = rs(0).Value
        SeriesRaces(1, j) = Replace(rs(1).Value, "''", "'") & "<br>" & Replace(rs(2).Value, "''", "'")
        j = j + 1
        ReDim Preserve SeriesRaces(1, j)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function GetMyPts(lPartID, lRaceID)
    GetMyPts = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Points FROM CCSeriesStdgs WHERE CCSeriesPartsID = " & lPartID & " AND RacesID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetMyPts = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE (Gopher State Events) CC/Nordic Series Results</title>
<meta name="description" content="Gopher State Events (GSE) Cross-Country/Nordic Ski Series Race-by-Race summary.">
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="row">
	    <h1 class="h1">Gopher State Events CC/Nordic Series Race-by-Race Results</h1>

        <ul class="nav">
            <li class="nav-item"><a class="nav-link" href="series_results.asp" style="font-weight: bold;color: #f00;">Standings</a></li>
            <li class="nav-item"><a class="nav-link" href="javascript:pop('how_it_works.asp',600,650)">How It Works</a></li>
            <li class="nav-item"><a class="nav-link" href="javascript:window.print();">Print Page</a></li>
        </ul>

        <form role="form" class="form-inline" name="get_series" method="post" action="race-by-race.asp">
        <div class="form-group"></div>
            <label for="series">Series:</label>
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
        </div>
        </form>

        <%If Not CLng(lSeriesID) = 0 Then%>
            <table class="table table-striped">
                <tr>
                    <th>No.</th>
                    <th>Name</th>
                    <th>School</th>
                    <%For i = 0 To UBound(SeriesRaces, 2) - 1%>
                        <th style="text-align: center;"><%=SeriesRaces(1, i)%></th>
                    <%Next%>
                    <th>Total</th>
                </tr>
                <%If UBound(SeriesParts, 2) > 0 Then%>
                    <%For j = 0 To UBound(SeriesParts, 2)%>
                        <tr>
                            <td><%=j + 1%>)</td>
                            <td><%=SeriesParts(1, j)%></td>
                            <td><%=SeriesParts(2, j)%></td>
                            <%For i = 0 To UBound(SeriesRaces, 2) - 1%>
                                <td style="text-align: center;"><%=GetMyPts(SeriesParts(0, j), SeriesRaces(0, i))%></td>
                            <%Next%>
                            <td style="text-align: right;"><%=SeriesParts(3, j)%></td>
                        </tr>
                    <%Next%>
                <%End If%>
            </table>
        <%End If%>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
