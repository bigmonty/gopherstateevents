<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lSeriesID, lMyID
Dim sSeriesName, sTime, sMyName, sGender
Dim iAgeTo, iAgeFrom, iMyAge, iNumGndr, iGndrPl, iCategPl, iNumCateg
Dim sngGndrPts, sngCategPts, sngMyCategTtl, sngMyGndrTtl
Dim SeriesEvents(), SeriesRaces(), Categories(1, 13)

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"
If CLng(lSeriesID) < 0 Then Response.Redirect "http://www.google.com"

lMyID = Request.QueryString("my_id")
If CStr(lMyID) = vbNullString Then lMyID = 0
If Not IsNumeric(lMyID) Then Response.Redirect "http://www.google.com"
If CLng(lMyID) < 0 Then Response.Redirect "http://www.google.com"

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

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
sSeriesName = Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing

j = 0
ReDim SeriesEvents(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate, Location FROM SeriesEvents WHERE EventDate BETWEEN '1/1/2013' AND '" & Date & "' AND SeriesID = " & lSeriesID
sql = sql & " ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    SeriesEvents(0, j) = rs(0).Value
	SeriesEvents(1, j) = Replace(rs(1).Value, "''", "'")
    SeriesEvents(2, j) = rs(2).Value
    SeriesEvents(3, j) = Replace(rs(3).Value, "''", "'")
	j = j + 1
	ReDim Preserve SeriesEvents(3, j)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get series races
j = 0
ReDim SeriesRaces(2, 0)
For i = 0 To UBound(SeriesEvents, 2) - 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sr.RaceID, sr.RaceName, se.EventDate FROM SeriesRaces sr INNER JOIN SeriesEvents se ON sr.SeriesEventsID = se.SeriesEventsID "
    sql = sql & "WHERE se.EventID = " & SeriesEvents(0, i)
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesRaces(0, j) = rs(0).Value
        SeriesRaces(1, j) = SeriesEvents(1, i) & " " & Replace(rs(1).Value, "''", "'")
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
sql = "SELECT PartName, Gender, Age FROM SeriesParts WHERE ParticipantID = " & lMyID
rs.Open sql, conn, 1, 2
sMyName = Replace(rs(0).Value, "''", "'")
sGender = UCase(rs(1).Value)
iMyAge = rs(2).Value
rs.Close
Set rs = Nothing

'get age group
For i = 0 To UBound(Categories, 2)
    If CInt(Categories(0, i)) >= CInt(iMyAge) Then
        iAgeTo = Categories(0, i)
        iAgeFrom = CInt(Categories(0, i)) - 4
        Exit For
    End If
Next

sngMyCategTtl = 0
sngMyGndrTtl = 0

Private Sub GetMyRslts(lThisRaceID)
    Dim bInRace

    bInRace = False

    iNumGndr = 0
    iNumCateg = 0

    iGndrPl = 0
    iCategPl = 0
    sngGndrPts = 0
    sngCategPts = 0

    sTime = vbNullString

    'see if they are in the race
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM IndResults WHERE ParticipantID = " & lMyID & " AND RaceID = " & lThisRaceID 
    sql = sql & " AND FnlTime IS NOT NULL AND FnlTime <> '00:00:00.000'" 
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then bInRace = True
    rs.Close
    Set rs = Nothing

    If bInRace = True Then
        'get gender total
        j = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ir.ParticipantID, ir.FnlTime FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " 
        sql = sql & lThisRaceID & " AND ir.FnlTime IS NOT NULL AND ir.FnlTime <> '00:00:00.000' AND p.Gender = '" & sGender & "' ORDER BY ir.EventPl" 
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

        'get category total
        j = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ir.ParticipantID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID "
        sql = sql & "INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID WHERE ir.RaceID = " & lThisRaceID & " AND ir.FnlTime IS NOT NULL "
        sql = sql & "AND ir.FnlTime <> '00:00:00.000' AND p.Gender = '" & sGender & "' AND pr.Age BETWEEN " & iAgeFrom & " AND " & iAgeTo  
        sql = sql & " ORDER BY EventPl"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iNumCateg = rs.RecordCount
        Do While Not rs.EOF
            j = j + 1
            If CLng(rs(0).Value) = CLng(lMyID) Then
                iCategPl = j
                Exit Do
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If Not (CInt(iNumGndr) = 0 Or CInt(iNumCateg) = 0) Then
            sngGndrPts = Round(((CInt(iNumGndr) - CInt(iGndrPl) + 1)/CInt(iNumGndr))*100, 2)
            sngCategPts = Round(((CInt(iNumCateg) - CInt(iCategPl) + 1)/CInt(iNumCateg))*100, 2)
        End If

        sngMyGndrTtl = CSng(sngMyGndrTtl) + CSng(sngGndrPts)
        sngMyCategTtl = CSng(sngMyCategTtl) + CSng(sngCategPts)
    End If
End Sub
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE (Gopher State Events) Series Results</title>
<meta name="description" content="GSE race series for road races, nordic ski, showshoe events, mountain bike, duathlon, and cross-country meet management (timing).">
<!--#include file = "../includes/js.asp" -->

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
<div style="width:470px;background-color:#fff;padding: 10px;margin: 10px;font-size: 0.9em;">
	<h3 style="margin:10px 0 0 0;padding:5px;">My GSE Series Results: <%=sSeriesName%></h3>

    <h4 style="background: none;text-align: left;color: #000;"><%=sMyName%></h4>
    <h5 style="background: none;text-align: left;color: #000;">Gender: <%=sGender%></h5>
    <h5 style="background: none;text-align: left;color: #000;">Age (on date of first race): <%=iMyAge%></h5>

    <table>
        <tr>
            <th rowspan="2" valign="bottom">Event-Race</th>
            <th rowspan="2" valign="bottom">Date</th>
            <th rowspan="2" valign="bottom">Time</th>
            <th style="border-bottom: 1px solid #ccc;" colspan="2">Gender</th>
            <th style="border-bottom: 1px solid #ccc;" colspan="2">Age Group</th>
        </tr>
        <tr>
            <th>Pl</th>
            <th>Points</th>
            <th>Pl</th>
            <th>Points</th>
        </tr>
        <%For i = 0 To UBound(SeriesRaces, 2) - 1%>
            <%Call GetMyRslts(SeriesRaces(0, i))%>

            <tr>
                <%If i mod 2 = 0 Then%>
                    <td style="text-align: left;" class="alt"><%=SeriesRaces(1, i)%></td>
                    <td class="alt"><%=SeriesRaces(2, i)%></td>
                    <td class="alt"><%=sTime%></td>
                    <td class="alt"><%=iGndrPl%></td>
                    <td class="alt"><%=sngGndrPts%></td>
                    <td class="alt"><%=iCategPl%></td>
                    <td class="alt"><%=sngCategPts%></td>
                <%Else%>
                    <td style="text-align: left;"><%=SeriesRaces(1, i)%></td>
                    <td><%=SeriesRaces(2, i)%></td>
                    <td><%=sTime%></td>
                    <td><%=iGndrPl%></td>
                    <td><%=sngGndrPts%></td>
                    <td><%=iCategPl%></td>
                    <td><%=sngCategPts%></td>
                <%End If%>
            </tr>
        <%Next%>
        <tr>
            <th style="text-align: right;padding-right: 10px;" colspan="4">Totals:</th>
            <th style="text-align: center;"><%=sngMyGndrTtl%></th>
            <th>&nbsp;</th>
            <th style="text-align: center;"><%=sngMyCategTtl%></th>
        </tr>
    </table>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>