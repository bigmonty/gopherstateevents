<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim iEvntReg, iEvntRegM, iEvntRegF, iEvntFin, iEvntFinM, iEvntFinF, iRaceReg, iRaceFin, iRaceRegM, iRaceRegF
Dim iRaceFinM, iRaceFinF
Dim iNumRace, iBegAge
Dim sEventName, sLogo, sEventRaces, sGender, sRaceName
Dim sngEvntPct, sngEvntPctM, sngEvntPctF, sngRacePct, sngRacePctM, sngRacePctF
Dim AgeGrps(), Races()
Dim dEventDate

'Response.Redirect "/misc/taking_break.htm"

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event name & event group
sql = "SELECT EventName, EventDate, Logo FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sLogo = rs(2).Value
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sEventRaces = sEventRaces & rs(0).Value & ", "
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

'****************************************************************************************
'get overall event reg
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PartRegID FROM PartReg WHERE RaceID IN (" & sEventRaces & ")"
rs.Open sql, conn, 1, 2
If rs.RecordCount  > 0 Then iEvntReg = rs.RecordCount
rs.Close
Set rs = Nothing

'get overall event fin
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ") AND FnlScnds > 0"
rs.Open sql, conn, 1, 2
If rs.RecordCount  > 0 Then iEvntFin = rs.RecordCount
rs.Close
Set rs = Nothing

sngEvntPct = "na"
If CInt(iEvntReg) > 0 Then sngEvntPct = Round((CInt(iEvntFin)/CInt(iEvntReg))*100, 2)

'get male event reg
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT pr.PartRegID FROM PartReg pr INNER JOIN Participant p ON pr.ParticipantID = p.ParticipantID "
sql = sql & "WHERE pr.RaceID IN (" & sEventRaces & ") AND p.Gender = 'M'"
rs.Open sql, conn, 1, 2
If rs.RecordCount  > 0 Then iEvntRegM = rs.RecordCount
rs.Close
Set rs = Nothing

'get male event fin
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID "
sql = sql & "WHERE ir.RaceID IN (" & sEventRaces & ") AND p.GEnder = 'M' AND ir.FnlScnds > 0"
rs.Open sql, conn, 1, 2
If rs.RecordCount  > 0 Then iEvntFinM = rs.RecordCount
rs.Close
Set rs = Nothing

sngEvntPctM = "na"
If CInt(iEvntRegM) > 0 Then sngEvntPctM = Round((CInt(iEvntFinM)/CInt(iEvntRegM))*100, 2)

'get female event reg
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT pr.PartRegID FROM PartReg pr INNER JOIN Participant p ON pr.ParticipantID = p.ParticipantID "
sql = sql & "WHERE pr.RaceID IN (" & sEventRaces & ") AND p.Gender = 'F'"
rs.Open sql, conn, 1, 2
If rs.RecordCount  > 0 Then iEvntRegF = rs.RecordCount
rs.Close
Set rs = Nothing

'get female event fin
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID "
sql = sql & "WHERE ir.RaceID IN (" & sEventRaces & ") AND p.GEnder = 'F' AND ir.FnlScnds > 0"
rs.Open sql, conn, 1, 2
If rs.RecordCount  > 0 Then iEvntFinF = rs.RecordCount
rs.Close
Set rs = Nothing

sngEvntPctF = "na"
If CInt(iEvntRegF) > 0 Then sngEvntPctF = Round((CInt(iEvntFinF)/CInt(iEvntRegF))*100, 2)

'***************************************************************************************

'***************************************************************************************
'get event races
i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Races(0, i) = rs(0).Value
    Races(1, i) = Replace(rs(1).Value, "''", "'")
    i = i + 1
    ReDim Preserve Races(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetRaceStats(lThisRace)
    'get overall event reg
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PartRegID FROM PartReg WHERE RaceID = " & lThisRace
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iRaceReg = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get overall event fin
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM IndResults WHERE RaceID = " & lThisRace & " AND FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iRaceFin = rs.RecordCount
    rs.Close
    Set rs = Nothing

    sngRacePct = "na"
    If CInt(iRaceReg) > 0 Then sngRacePct = Round((CInt(iRaceFin)/CInt(iRaceReg))*100, 2)

    'get male event reg
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT pr.PartRegID FROM PartReg pr INNER JOIN Participant p ON pr.ParticipantID = p.ParticipantID "
    sql = sql & "WHERE pr.RaceID = " & lThisRace & " AND p.Gender = 'M'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iRaceRegM = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get male event fin
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID "
    sql = sql & "WHERE ir.RaceID = " & lThisRace & " AND p.GEnder = 'M' AND ir.FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iRaceFinM = rs.RecordCount
    rs.Close
    Set rs = Nothing

    sngRacePctM = "na"
    If CInt(iRaceRegM) > 0 Then sngRacePctM = Round((CInt(iRaceFinM)/CInt(iRaceRegM))*100, 2)

    'get female event reg
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT pr.PartRegID FROM PartReg pr INNER JOIN Participant p ON pr.ParticipantID = p.ParticipantID "
    sql = sql & "WHERE pr.RaceID  = " & lThisRace & " AND p.Gender = 'F'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iRaceRegF = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get female event fin
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID "
    sql = sql & "WHERE ir.RaceID  = " & lThisRace & " AND p.GEnder = 'F' AND ir.FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iRaceFinF = rs.RecordCount
    rs.Close
    Set rs = Nothing

    sngRacePctF = "na"
    If CInt(iRaceRegF) > 0 Then sngRacePctF = Round((CInt(iRaceFinF)/CInt(iRaceRegF))*100, 2)
End Sub


i = 0
iBegAge = 0
ReDim AgeGrps(1, 0)
sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & 1173
sql = sql & " ORDER BY EndAge"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    AgeGrps(0, i) = iBegAge
    AgeGrps(1, i) = rs(0).Value
    iBegAge = rs(0).Value + 1
    i = i + 1
    ReDim Preserve AgeGrps(1, i)
    rs.MoveNext
Loop
Set rs = Nothing
%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/pace_per_mile.asp" -->
<!--#include file = "../../includes/pace_per_km.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<meta name="description" content="GSE Fitness Events Records for road races, nordic ski, showshoe events, mountain bike, duathlon, and triathlon timing.">
<title>GSE Individual Results Page</title>
<!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <div class="row">
        <a href="javascript:window.print();">Print</a>
        &nbsp;|&nbsp;
        <a href="javascript:void(0)" nof="LS_E" OnMouseOver="window.close();self.close();close();top.window.close()">Close</a>
    </div>
    <div class="row">
        <div class="col-sm-6">
            <img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
        </div>
       <div class="col-sm-6">
            <img src="/events/logos/<%=sLogo%>" alt="<%=sEventName%>" class="img-responsive">
        </div>
    </div>

    <h3 class="h3">Event & Race Statistics: for <%=sEventName%> on <%=dEventDate%></h3>

    <p>(Note:  The difference between registered participants and finishers includes no shows, dnfs, and dqs.)</p>

    <h4 class="h4 bg-danger">Event Stats</h4>
    <table class="table table-striped">
        <tr><th>&nbsp;</th><th>Registered</th><th>Finished</th><th>% Finished</th></tr>
        <tr><tr><td>Male</td><td><%=iEvntRegM%></td><td><%=iEvntFinM%></td><td><%=sngEvntPctM%>%</td></tr>
        <tr><td>Female</td><td><%=iEvntRegF%></td><td><%=iEvntFinF%></td><td><%=sngEvntPctF%>%</td></tr>
        <tr><td>Total</td><td><%=iEvntReg%></td><td><%=iEvntFin%></td><td><%=sngEvntPct%>%</td></tr>
    </table>

    <h4 class="h4 bg-danger">Race Stats</h4>

    <%For i = 0 To UBound(Races, 2) - 1%>
        <h4 class="h4"><%=Races(1, i)%></h4>
        <%Call GetRaceStats(Races(0, i))%>
        <table class="table table-striped">
            <tr><th>&nbsp;</th><th>Registered</th><th>Finished</th><th>% Finished</th></tr>
            <tr><tr><td>Male</td><td><%=iRaceRegM%></td><td><%=iRaceFinM%></td><td><%=sngRacePctM%>%</td></tr>
            <tr><td>Female</td><td><%=iRaceRegF%></td><td><%=iRaceFinF%></td><td><%=sngRacePctF%>%</td></tr>
            <tr><td>Total</td><td><%=iRaceReg%></td><td><%=iRaceFin%></td><td><%=sngRacePct%>%</td></tr>
        </table>

        <h5 class="h5">Age Group Stats</h5>

        <hr>
   <%Next%>
    
    <p>This page is currently under construction.  We appreciate your patience!</p>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
