<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, k
Dim lEventID, lRaceID, lParticipantID
Dim sEventName, sRaceName, sPartName, sGender, sGunTime, sNetTime, sChipStart, sDist, sAgeGrp, sLogo, sMyPix, sShowAge, sLocation, sPacePerMi, sPacePerKm
Dim iAge, iRacePl, iAgeGrpPl, iGenderPl, iBib, iBegAge, iEndAge
Dim AgeGrps()
Dim dEventDate

'Response.Redirect "/misc/taking_break.htm"

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect "http://www.google.com"
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

iBib = Request.QueryString("bib")
If CStr(iBib) = vbNullString Then iBib = 0
If Not IsNumeric(iBib) Then Response.Redirect "http://www.google.com"
If CLng(iBib) < 0 Then Response.Redirect("http://www.google.com")

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

sql = "SELECT RaceName, Dist, ShowAge FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'") 
sDist = rs(1).Value
sShowAge = rs(2).Value
Set rs = Nothing

sql = "SELECT ParticipantID, Age, AgeGrp FROM PartRace WHERE Bib = " & iBib & " AND RaceID = " & lRaceID
Set rs = conn.Execute(sql)
lParticipantID = rs(0).Value
iAge = rs(1).Value
sAgeGrp = rs(2).Value
Set rs = Nothing

sql = "SELECT LastName, FirstName, Gender, City, St FROM Participant WHERE ParticipantID = " & lParticipantID
Set rs = conn.Execute(sql)
sPartName = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") 
sGender = rs(2).Value
If Not rs(3).Value & "" = "" Then sLocation = Replace(rs(3).Value, "''", "'")
If Not rs(4).Value & "" = "" Then
    If sLocation = vbNullString Then 
        sLocation = Replace(rs(4).Value, "''", "'")
    Else
        sLocation = sLocation & ", " & Replace(rs(4).Value, "''", "'")
    End If
End If
Set rs = Nothing

'get age groups for this race
i = 0
iBegAge = 0
ReDim AgeGrps(1, 0)
sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & lRaceID
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

k = 1
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ParticipantID FROM IndResults WHERE RaceID = " & lRaceID & " AND FnlScnds > 0 AND EventPl > 0 ORDER BY FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If CLng(rs(0).Value) = CLng(lParticipantID) Then
        iRacePl = k
        Exit Do
    Else
        k = k + 1
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

k = 1
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ir.ParticipantID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " & lRaceID 
sql = sql & " AND p.Gender = '" & sGender & "' AND ir.FnlScnds > 0 AND ir.EventPl > 0 ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If CLng(rs(0).Value) = CLng(lParticipantID) Then
        iGenderPl = k
        Exit Do
    Else
        k = k + 1
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get end age
sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & lRaceID & " AND EndAge > " & iAge & " ORDER BY EndAge"
Set rs = conn.Execute(sql)
iEndAge = rs(0).Value
Set rs = Nothing

'get beg age
sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & lRaceID & " AND EndAge < " & iAge & " ORDER BY EndAge DESC"
Set rs = conn.Execute(sql)
If rs.EOF = rs.EOF Then
    iBegAge = 0
Else
    iBegAge = CInt(rs(0).Value) + 1
End If
Set rs = Nothing

iAgeGrpPl = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT pr.Bib FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID "
sql = sql & "INNER JOIN Participant p ON p.ParticipantID = ir.ParticipantID WHERE (ir.RaceID = " & lRaceID & " AND pr.RaceID = " & lRaceID 
sql = sql & ") AND p.Gender = '" & sGender & "' AND pr.AgeGrp = '" & sAgeGrp & "' AND ir.FnlTime IS NOT NULL "
sql = sql & "AND ir.FnlTime > '00:00:00.000' ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    iAgeGrpPl = CInt(iAgeGrpPl) + 1
    If CInt(rs(0).Value) = CInt(iBib) Then Exit Do
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sql = "SELECT ChipTime, FnlTime, ChipStart FROM IndResults WHERE RaceID = " & lRaceID & " AND ParticipantID = " & lParticipantID
Set rs = conn.Execute(sql)
sNetTime = rs(0).Value
sGunTime = rs(1).Value
sChipStart = rs(2).Value
Set rs = Nothing

Private Function AgeGrdTime(sThisGndr, iThisAge, sThisTime)
    Dim lngAgeGrDistID
    Dim sngThisTime

    sngThisTime = ConvertToSeconds(sThisTime)
    AgeGrdTime = "na"

    lngAgeGrDistID = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT AgeGrDistID FROM AgeGrDist WHERE Distance = '" & sDist & "'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then lngAgeGrDistID = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing

    If CLng(lngAgeGrDistID) > 0 Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT Factor FROM AgeGrFactors WHERE MF = '" & LCase(sThisGndr) & "' AND Age = " & iThisAge & " AND AgeGrDistID = " & lngAgeGrDistID
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then AgeGrdTime = ConvertToMinutes(CSng(sngThisTime)*CSng(rs2(0).Value))
        rs2.Close
        Set rs2 = Nothing
    End If
End Function
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
<div class="container" style="margin:5px;">
    <div class="row">
        <a href="javascript:window.print();">Print</a>
        &nbsp;|&nbsp;
        <a href="javascript:void(0)" nof="LS_E" OnMouseOver="window.close();self.close();close();top.window.close()">Close</a>
    </div>
    <img src="/events/logos/<%=sLogo%>" alt="<%=sEventName%>" class="img-responsive" width="75" style="float: right;">
    <h4 class="h4 bg-success">My Results: <br><%=sEventName%></h4>
    <h5 class="h5"><%=sRaceName%> on <%=dEventDate%></h5>
    <h4 class="h4"><%=sPartName%> (<%=iBib%>)</h4>

    <ul class="list-group">
        <li class="list-group-item list-group-item-warning" style="padding: 2px 15px;">Gender: <%=sGender%></li>
        <%If sShowAge = "y" Then%>
            <li class="list-group-item list-group-item-warning" style="padding: 2px 15px;">Age: <%=iAge%></li>
        <%End If%>
        <%IF UBound(AgeGrps, 2) > 1 Then%>
            <li class="list-group-item list-group-item-warning" style="padding: 2px 15px;">Age Group: <%=sAgeGrp%></li>
        <%End If%>
        <li class="list-group-item list-group-item-warning" style="padding: 2px 15px;">Location: <%=sLocation%></li>
    </ul>

    <ul class="list-group">
        <li class="list-group-item list-group-item-info" style="padding: 2px 15px;">Race Place: <%=iRacePl%></li>
        <li class="list-group-item list-group-item-info" style="padding: 2px 15px;">Gender Place: <%=iGenderPl%></li>
        <%IF UBound(AgeGrps, 2) > 1 Then%>
            <li class="list-group-item list-group-item-info" style="padding: 2px 15px;">Age Group Place: <%=iAgeGrpPl%></li>
        <%End If%>
    </ul>

    <ul class="list-group">
        <li class="list-group-item list-group-item-danger" style="padding: 2px 15px;">Gun Time: <%=sGunTime%></li>
        <li class="list-group-item list-group-item-danger" style="padding: 2px 15px;">Chip Time: <%=sNetTime%></li>
        <li class="list-group-item list-group-item-danger" style="padding: 2px 15px;">Start Delay: <%=sChipStart%></li>
        <li class="list-group-item list-group-item-danger" style="padding: 2px 15px;">Pace Per Mile: <%=PacePerMile(ConvertToSeconds(sNetTime), sDist)%></li>
        <li class="list-group-item list-group-item-danger" style="padding: 2px 15px;">Pace Per Km: <%=PacePerKm(ConvertToSeconds(sNetTime), sDist)%></li>
        <li class="list-group-item list-group-item-danger" style="padding: 2px 15px;">Age Graded Time: <%=AgeGrdTime(sGender, iAge, sNetTime)%></li>
    </ul>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
