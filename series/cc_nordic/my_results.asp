<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lSeriesID, lMyID
Dim sSeriesName, sTime, sMyName, sGender
Dim iAgeTo, iAgeFrom, iMySchl, iNumGndr, iGndrPl
Dim sngGndrPts, sngMyGndrTtl
Dim SeriesMeets(), SeriesRaces()

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then Response.Redirect "http://www.google.com"

lMyID = Request.QueryString("my_id")
If CStr(lMyID) = vbNullString Then Response.Redirect "http://www.google.com"

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

j = 0
ReDim SeriesMeets(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate, Location FROM SeriesMeets WHERE MeetDate BETWEEN '1/1/2013' AND '" & Date & "' AND CCSeriesID = " & lSeriesID
sql = sql & " ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    SeriesMeets(0, j) = rs(0).Value
	SeriesMeets(1, j) = Replace(rs(1).Value, "''", "'")
    SeriesMeets(2, j) = rs(2).Value
    SeriesMeets(3, j) = Replace(rs(3).Value, "''", "'")
	j = j + 1
	ReDim Preserve SeriesMeets(3, j)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get series races
j = 0
ReDim SeriesRaces(2, 0)
For i = 0 To UBound(SeriesMeets, 2) - 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sr.RacesID, sr.RaceName, seMeetDate FROM CCSeriesRaces sr INNER JOIN CCSeriesMeets se ON sr.CCSeriesMeetsID = se.CCSeriesMeetsID "
    sql = sql & "WHERE se.MeetsID = " & SeriesMeets(0, i)
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesRaces(0, j) = rs(0).Value
        SeriesRaces(1, j) = SeriesMeets(1, i) & " " & Replace(rs(1).Value, "''", "'")
        SeriesRaces(2, j) = rs(2).Value
        j = j + 1
        ReDim Preserve SeriesRaces(2, j)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
Next

'get participant name
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PartName, Gender, School FROM CCSeriesParts WHERE RosterID = " & lMyID
rs.Open sql, conn, 1, 2
sMyName = Replace(rs(0).Value, "''", "'")
sGender = UCase(rs(1).Value)
iMySchl = rs(2).Value
rs.Close
Set rs = Nothing

sngMyGndrTtl = 0

Private Sub GetMyRslts(lThisRaceID)
    Dim bInRace

    bInRace = False

    iNumGndr = 0
    iGndrPl = 0
    sngGndrPts = 0

    sTime = vbNullString

    'see if they are in the race
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID FROM IndRslts WHERE RosterID = " & lMyID & " AND RacesID = " & lThisRaceID  & " AND RaceTime IS NOT NULL AND RaceTime > '00:00:00.000'" 
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then bInRace = True
    rs.Close
    Set rs = Nothing

    If bInRace = True Then
        'get gender total
        j = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ir.RosterID, ir.RaceTime FROM IndRslts ir INNER JOIN Roster p ON ir.RosterID = p.RosterID WHERE ir.RacesID = " 
        sql = sql & lThisRaceID & " AND ir.RaceTime IS NOT NULL AND ir.RaceTime > '00:00:00.000' AND p.Gender = '" & sGender & "' ORDER BY ir.Place" 
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iNumGndr = rs.RecordCount
        Do While Not rs.EOF
            j = j + 1
            If CLng(rs(0).Value) = CLng(lMyID) Then
                sTime = rs(1).Value
                iGndrPl = j
                Exit Do
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If Not (CInt(iNumGndr) = 0 Then sngGndrPts = Round(((CInt(iNumGndr) - CInt(iGndrPl) + 1)/CInt(iNumGndr))*100, 2)
        sngMyGndrTtl = CSng(sngMyGndrTtl) + CSng(sngGndrPts)
    End If
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE (Gopher State Meets) CC/Nordic Series Results</title>
<meta name="description" content="Gopher State Events (GSE) Cross-Country/Nordic Ski Series Individual Results page.">
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td,th{
        padding: 2px 0 2px 5px;
    }
    
    table{
        margin: 0;
        font-size: 0.88em;
    }
</style>
</head>


<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="row">
	    <h1 class="h1">My GSE CC/Nordic Series Results: <%=sSeriesName%></h1>

        <h4 class="h4"><%=sMyName%></h4>
        <h5 class="h5">Gender: <%=sGender%></h5>
        <h5 class="h5">School: <%=iMySchl%></h5>

        <table class="table-striped">
            <tr>
                <th>Meet-Race</th>
                <th>Date</th>
                <th>Time</th>
                <th>Pl</th>
                <th>Pts</th>
            </tr>
            <%For i = 0 To UBound(SeriesRaces, 2) - 1%>
                <%Call GetMyRslts(SeriesRaces(0, i))%>

                <tr>
                    <td style="text-align: left;"><%=SeriesRaces(1, i)%></td>
                    <td><%=SeriesRaces(2, i)%></td>
                    <td><%=sTime%></td>
                    <td><%=iGndrPl%></td>
                    <td><%=sngGndrPts%></td>
                </tr>
            <%Next%>
            <tr>
                <th style="text-align: right;padding-right: 10px;">Totals:</th>
                <th style="text-align: center;" colspan="4"><%=sngMyGndrTtl%></th>
            </tr>
        </table>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>