<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, k
Dim lEventID, lRaceID, lParticipantID
Dim sEventName, sRaceName, sPartName, sGender, sGunTime, sNetTime, sChipStart, sDist, sAgeGrp, sLogo, sMyPix, sMantra, sShowAge
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

sql = "SELECT LastName, FirstName, Gender FROM Participant WHERE ParticipantID = " & lParticipantID
Set rs = conn.Execute(sql)
sPartName = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") 
sGender = rs(2).Value
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
sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & lRaceID & " AND EndAge < " & iAge & " ORDER BY EndAge"
Set rs = conn.Execute(sql)
If rs.EOF = rs.EOF Then
    iBegAge = 0
Else
    iBegAge = CInt(rs(0).Value) + 1
End If
Set rs = Nothing

k = 1
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ir.ParticipantID FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID WHERE ir.RaceID = " & lRaceID 
sql = sql & " AND pr.Age >= " & iBegAge & " AND pr.Age <= " & iEndAge & " AND ir.FnlScnds > 0 AND ir.EventPl > 0 ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If CLng(rs(0).Value) = CLng(lParticipantID) Then
        iAgeGrpPl = k
        Exit Do
    Else
        k = k + 1
    End If
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

If sMantra = vbNullString Then sMantra = "n/a"
%>
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
    <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
    <h4 class="h4">Individual Results For <%=sPartName%><br><small><%=sEventName%> - <%=sRaceName%></small></h4>

    <div class="row">
        <div class="col-sm-4">
            <%If Not sLogo & "" = "" Then%>
                <img src="/events/logos/<%=sLogo%>" alt="<%=sEventName%>"  class="img-responsive">
            <%End If%>
            <br><br>
            <%If sMyPix & "" = "" Then%>
                <img src="/graphics/photo_na.gif" alt="Photo Unavailable"  class="img-responsive">
            <%Else%>
            
            <%End If%>

            <p style="margin-top:10px;"><span style="font-weight: bold;">My Mantra:</span> <%=sMantra%></p>
        </div>
        <div class="col-sm-8">
            <ul>
                <li>Gender: <%=sGender%></li>
                <%If sShowAge = "y" Then%>
                    <li>Age: <%=iAge%></li>
                <%End If%>
                <%IF UBound(AgeGrps, 2) > 1 Then%>
                    <li>Age Group: <%=sAgeGrp%></li>
                <%End If%>
                <li>--------------------</li>
                <li>Race Place: <%=iRacePl%></li>
                <li>Gender Place: <%=iGenderPl%></li>
                <%IF UBound(AgeGrps, 2) > 1 Then%>
                    <li>Age Group Place: <%=iAgeGrpPl%></li>
                <%End If%>
                <li>--------------------</li>
                <li>Gun Time: <%=sGunTime%></li>
                <li>Net Time: <%=sNetTime%></li>
                <li>Start Delay: <%=sChipStart%></li>
                <li>--------------------</li>
                <li>Splits: n/a</li>
                <li>PR: n/a</li>
            </ul>
        </div>
    </div>
    <p>(Note:  You may add a mantra and a photograph by creating a "My History" 
        account <a href="http://www.gopherstateevents.com/perf_center/my_hist.asp" onclick="openThis(this.href,1024,768);return false;">here</a>.)</p>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
