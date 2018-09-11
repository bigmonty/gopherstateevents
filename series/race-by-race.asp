<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lSeriesID
Dim sSeriesName, sSeriesRaces, sGender
Dim iYear, iAgeTo, iAgeFrom
Dim Series(), SeriesParts, SeriesRaces(), Categories(1, 13)

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"
If CLng(lSeriesID) < 0 Then Response.Redirect "http://www.google.com"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")
If CLng(iYear) < 0 Then Response.Redirect "http://www.google.com"

Categories(0,0) = "0"
Categories(1,0) = "Open"
Categories(0,1) = "14"
Categories(1,1) = "14 & Under"
Categories(0,2) = "19"
Categories(1,2) = "15 - 19"
Categories(0,3) = "24"
Categories(1,3) = "20 -24"
Categories(0,4) = "29"
Categories(1,4) = "25 - 29"
Categories(0,5) = "34"
Categories(1,5) = "30 - 34"
Categories(0,6) = "39"
Categories(1,6) = "35 - 39"
Categories(0,7) = "44"
Categories(1,7) = "40 - 44"
Categories(0,8) = "49"
Categories(1,8) = "45 - 49"
Categories(0,9) = "54"
Categories(1,9) = "50 - 54"
Categories(0,10) = "59"
Categories(1,10) = "55 - 59"
Categories(0,11) = "64"
Categories(1,11) = "60 -64"
Categories(0,12) = "69"
Categories(1,12) = "65 - 69"
Categories(0,13) = "99"
Categories(1,13) = "70 & Over"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID, SeriesName, SeriesYear FROM Series ORDER BY SeriesYear DESC, SeriesName"
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

    sGender = Request.Form.Item("gender")
    iAgeTo = Request.Form.Item("age_to")
End If

If sGender = vbNullString Then sGender = "M"
If CStr(iAgeTo) = vbNullString Then iAgeTo = "0"

If CInt(iAgeTo) > 14 Then
    iAgeFrom = CInt(iAgeTo) - 4
Else
    iAgeFrom = "0"
End If

If CLng(lSeriesID) > 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
 
    Set rs = Server.CreateObject("ADODB.Recordset")
    If iAgeTo = "0" Then
        sql = "SELECT sp.SeriesPartsID, sp.PartName, sp.Gender, sp.Age, sr.GndrScore FROM SeriesParts sp INNER JOIN SeriesResults sr "
        sql = sql & "ON sp.SeriesPartsID = sr.SeriesPartsID WHERE sp.SeriesID = " & lSeriesID & " AND sp.Gender = '" & sGender & "' ORDER BY sp.PartName"
    Else
        sql = "SELECT sp.SeriesPartsID, sp.PartName, sp.Gender, sp.Age, sr.AgeScore FROM SeriesParts sp INNER JOIN SeriesResults sr "
        sql = sql & "ON sp.SeriesPartsID = sr.SeriesPartsID WHERE sp.SeriesID = " & lSeriesID & " AND sp.Gender = '" & sGender & "' AND Age >= " 
        sql = sql & iAgeFrom & " AND Age <= " & iAgeTo & " ORDER BY sp.PartName"
    End If
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        SeriesParts = rs.GetRows()
    Else
        ReDim SeriesParts(4, 0)
    End If
    rs.Close
    Set rs = Nothing

    j = 0
    ReDim SeriesRaces(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sr.RaceID, sm.EventName, sr.RaceName, sm.EventDate FROM SeriesRaces sr INNER JOIN SeriesEvents sm ON sr.SeriesEventsID = sm.SeriesEventsID "
    sql = sql & "WHERE sm.SeriesID = " & lSeriesID & " ORDER BY sm.EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesRaces(0, j) = rs(0).Value
        SeriesRaces(1, j) = Replace(rs(1).Value, "''", "'") & "<br>" & Replace(rs(2).Value, "''", "'") & "<br>" & rs(3).Value
        j = j + 1
        ReDim Preserve SeriesRaces(1, j)

        sSeriesRaces = sSeriesRaces & rs(0).Value & ","

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    sSeriesRaces = Trim(Left(sSeriesRaces, Len(sSeriesRaces) - 1))
End If

Private Function GetMyPts(lPartID, lRaceID)
    GetMyPts = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    If CInt(iAgeTo) = 0 Then
        sql = "SELECT GndrPts FROM SeriesStdgs WHERE SeriesPartsID = " & lPartID & " AND RaceID = " & lRaceID
    Else
        sql = "SELECT AgePts FROM SeriesStdgs WHERE SeriesPartsID = " & lPartID & " AND RaceID = " & lRaceID
    End If
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetMyPts = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function

Private Function GetNumRaces(lPartID)
    GetNumRaces = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesStdgsID FROM SeriesStdgs WHERE SeriesPartsID = " & lPartID & " AND RaceID IN (" & sSeriesRaces & ")"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetNumRaces = rs.RecordCount
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE (Gopher State Events) Series Results</title>
<meta name="description" content="Gopher State Events (GSE) race series results: race-by-race.">
<!--#include file = "../includes/js.asp" -->

<!--
<link href="//cdn.datatables.net/1.10.2/css/jquery.dataTables.css" rel="stylesheet" type="text/css">
    
<script type="text/javascript" src="//code.jquery.com/jquery-1.11.1.min.js"></script>
<script type="text/javascript" src="//cdn.datatables.net/1.10.2/js/jquery.dataTables.min.js"></script>
-->
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-xs-5">
            <img src="/graphics/html_header.png" alt="Series Header" class="img-responsive">
        </div>
        <div class="col-xs-7">
            <h2 class="h2">GSE Series Standings: Race-by-Race Results</h2>
        </div>
    </div>

    <!--#include file = "series_nav.asp" -->

     <div class="row bg-warning">
        <form role="form" class="form-inline" name="get_series" method="post" action="race-by-race.asp?year=<%=iYear%>">
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
        <label for="age_to">Category:</label>
        <select class="form-control" name="age_to" id="age_to" onchange="this.form.submit1.click();">
            <%For i = 0 To UBound(Categories, 2)%>
                <%If CStr(iAgeTo) = CStr(Categories(0, i)) Then%>
                    <option value="<%=Categories(0, i)%>" selected><%=Categories(1, i)%></option>
                <%Else%>
                    <option value="<%=Categories(0, i)%>"><%=Categories(1, i)%></option>
                <%End If%>
            <%Next%>
        </select>
        <label for="gender">Gender:</label>
        <select class="form-control" name="gender" id="gender" onchange="this.form.submit1.click();">
            <%If sGender = "M" Then%>
                <option value="M">Male</option>
                <option value="F">Female</option>
            <%Else%>
                <option value="M">Male</option>
                <option value="F" selected>Female</option>
            <%End If%>
        </select>
        <input type="hidden" class="form-control" name="submit_series" id="submit_series" value="submit_series">
        <input type="submit" class="form-control" name="submit1" id="submit1" value="Select">
        </form>
    </div>

    <br><br>

    <%If Not CLng(lSeriesID) = 0 Then%>
        <table class="table table-striped">
            <tr>
                <th>No.</th>
                <th>Name</th>
                <th>M/F</th>
                <th>Age</th>
                <%For i = 0 To UBound(SeriesRaces, 2) - 1%>
                    <th style="text-align: center;"><%=SeriesRaces(1, i)%></th>
                <%Next%>
                <th style="text-align: right;">Total</th>
                <th style="text-align: right;">Races</th>
            </tr>
            <%If UBound(SeriesParts, 2) > 0 Then%>
                <%For j = 0 To UBound(SeriesParts, 2)%>
                    <tr>
                        <td><%=j + 1%>)</td>
                        <td style="white-space: nowrap;"><%=SeriesParts(1, j)%></td>
                        <td style="white-space: nowrap;text-align: center;"><%=SeriesParts(2, j)%></td>
                        <td style="text-align: center;"><%=SeriesParts(3, j)%></td>
                        <%For i = 0 To UBound(SeriesRaces, 2) - 1%>
                            <td style="text-align: center;"><%=GetMyPts(SeriesParts(0, j), SeriesRaces(0, i))%></td>
                        <%Next%>
                        <td style="text-align: right;"><%=SeriesParts(4, j)%></td>
                        <td style="text-align: right;"><%=GetNumRaces(SeriesParts(0, j))%></td>
                    </tr>
                <%Next%>
            <%End If%>
        </table>
    <%End If%>
</div>
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
