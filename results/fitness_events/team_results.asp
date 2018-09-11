<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lRaceID, lEventID
Dim i, j
Dim iMaxScore, iMinScore, iBegAge, iMyBib
Dim sTeamTime, sTeamAverage, sScoreMeth, sScoreMethod, sShowDetail, sEventName, sRaceName
Dim sngTeamAverage, sngMyTime
Dim TeamRslts, IndRslts(), AgeGroups(), Genders(2)
Dim dEventDate
 
lRaceID = Request.QueryString("race_id")

sShowDetail = Request.QueryString("show_detail")
If sShowDetail = vbNullString Then sShowDetail = "n"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT RaceName, EventID FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
lEventID = rs(1).Value
Set rs = Nothing

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = rs(0).Value
dEventDate = rs(1).Value
Set rs = Nothing

'get genders
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Male, Female, Combined FROM TeamGenders WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    Genders(0) = rs(0).Value
    Genders(1) = rs(1).Value
    Genders(2) = rs(2).Value
End If
rs.Close
Set rs = Nothing
    
'get age groups
i = 0
iBegAge = 0
ReDim AgeGroups(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AgeGrp, EndAge FROM TeamAgeGrps WHERE RaceID = " & lRaceID & " ORDER BY EndAge"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    AgeGroups(0, i) = rs(0).Value
    AgeGroups(1, i) = iBegAge
    AgeGroups(2, i) = rs(1).Value
    iBegAge = rs(1).Value + 1
    i = i + 1
    ReDim Preserve AgeGroups(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get scoring data
sql = "SELECT MaxScore, MinScore, ScoreMethod FROM TeamScoring WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
iMaxScore = rs(0).Value
iMinScore = rs(1).Value
sScoreMethod = rs(2).Value
Set rs = Nothing
   
'get score method
Select Case sScoreMethod
    Case "time"
        sScoreMeth = "Combined Time"
    Case "average"
        sScoreMeth = "Average Time"
    Case "min score"
        sScoreMeth = "Fewest Points (based on overall race place)"
    Case "max score"
        sScoreMeth = "Most Points (based on number of participants)"
End Select

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamPl, TeamName, CombTime, AvgTime, CCPts, NordicPts, TotalTime, TotalAvg, NumFin FROM Teams WHERE RaceID = "
sql = sql & lRaceID & " AND TeamPl >= 0 AND CCPts > 0 ORDER BY TeamPl"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    TeamRslts = rs.GetRows()
Else
    ReDim TeamRslts(9, 0)
End If
rs.Close
Set rs = Nothing

For i = 0 To UBound(TeamRslts, 2)
    If TeamRslts(3, i) = "8:20:00.0" Then 
        TeamRslts(1, i) = "DNF"
        TeamRslts(3, i) = "---"
        TeamRslts(4, i) = "---"
        TeamRslts(5, i) = "--"
        TeamRslts(6, i) = "--"
        If TeamRslts(7, i) = "8:20:00.0" Then
            TeamRslts(7, i) = "---"
            TeamRslts(8, i) = "---"
        End If
    End If
Next

Private Sub IndResults(lTeamID)
    Dim x

    x = 0
    ReDim IndRslts(6, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, MmbrName, Gender, Age, TmPlace FROM TeamMmbrs WHERE TeamsID = " & lTeamID
    sql = sql & " ORDER BY TmPlace"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Call GetMyRslts(rs(0).Value)
            
        IndRslts(0, x) = rs(0).Value
        IndRslts(1, x) = iMyBib
        IndRslts(2, x) = Replace(rs(1).Value, "''", "'")
        IndRslts(3, x) = rs(2).Value
        IndRslts(4, x) = rs(3).Value
        If sngMyTime = "30000" Then
            IndRslts(5, x) = "---"
        Else
            IndRslts(5, x) = rs(4).Value
        End If
        IndRslts(6, x) = sngMyTime
            
        x = x + 1
        ReDim Preserve IndRslts(6, x)
            
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetMyRslts(lThisPart)
    Dim rs2, sql2
    
    iMyBib = 0
    sngMyTime = "30000"
    
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Bib FROM PartRace WHERE ParticipantID = " & lThisPart & " AND RaceID = " & lRaceID & " ORDER BY PartRaceID DESC"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then iMyBib = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FnlScnds FROM IndResults WHERE ParticipantID = " & lThisPart & " AND RaceID = " & lRaceID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then
        If rs2(0).Value = "0" Then
            sngMyTime = "30000"
        Else
            sngMyTime = rs2(0).Value
        End If
    End If
    rs2.Close
    Set rs2 = Nothing
End Sub
%>
<!--#include file = "../../includes/convert_to_minutes.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events (GSE) Team Results</title>
<meta name="description" content="Gopher State Events (GSE) Team Results.">
<!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img class="img-responsive" src="/graphics/html_header.png" alt="Team Results">
    <h1 class="h1"><%=sEventName%> Team Results</h1>
    <h2 class="h2"><%=sRaceName%>  on <%=dEventDate%></h2>

    <div>
        <%If sShowDetail = "y" Then%>
            <a href="team_results.asp?race_id=<%=lRaceID%>&amp;show_detail=n">Simple View</a>
        <%Else%>
            <a href="team_results.asp?race_id=<%=lRaceID%>&amp;show_detail=y">Detailed View</a>
        <%End If%>
        &nbsp;|&nbsp;
        <a href="javascript:window.print()">Print This</a>
    </div>
    <br>
    <div class="col-sm-3 bg-warning">
         <h5 class="h5">Scoring Parameters:</h5>
        <ul>
            <li>Minimum Participants to Score: <%=iMinScore%></li>
            <li>Maximum Participants to Score: <%=iMaxScore%></li>
            <li>Scoring Method: <%=sScoreMeth%></li>
        </ul>
    </div>
    <div class="col-sm-1">&nbsp;</div>
    <div class="col-sm-8 bg-danger">
        <h5 class="h5">Scoring Legend:</h5>
        <ul>
           <li>"CUMULATIVE": combined time for scoring members.  Not useful when there is not a fixed
                number of scoring members.</li>
            <li>"AVERAGE": average time for scoring members.  Useful when there is not a fixed number of scoring members.</li>
            <li>"SCORE": sum of points for scoring members (similar to cross country running).  Lowest score wins.</li>
            <li>"POINTS": assigns maximum value to first team finisher decreases by one from there.  Highest score wins.</li>
        </ul>
    </div>

        <%If sShowDetail = "y" Then%>
            <h4 class="h4">Detailed View</h4>

            <table class="table">
                <tr><th>PL</th><th>TEAM</th><th>CUMULATIVE</th><th>AVERAGE</th><th>SCORE</th><th>POINTS</th></tr>
                <%For i = 0 To UBound(TeamRslts, 2)%>
                    <tr>
                        <%If TeamRslts(3, i) = "---" Then%>
                            <td class="bg-success">---</td>
                        <%Else%>
                            <td class="bg-success"><%=i + 1%></td>
                        <%End If%>
                        <td class="bg-success"><%=TeamRslts(2, i)%></td>
                        <td class="bg-success"><%=TeamRslts(3, i)%></td>
                        <td class="bg-success"><%=TeamRslts(4, i)%></td>
                        <td class="bg-success"><%=TeamRslts(5, i)%></td>
                        <td class="bg-success"><%=TeamRslts(6, i)%></td>
                    </tr>

                    <tr>
                        <td style="padding-left: 50px;" colspan="6">
                            <%Call IndResults(TeamRslts(0, i))%>
                            <table class="table table-condensed">
                                <tr><th>BIB</th><th style="text-align: left;">NAME</th><th>MF</th><th>AGE</th><th>PL</th><th>TIME</th></tr>   
                                <%For j = 0 To UBound(IndRslts, 2) - 1%>    
                                    <tr>
                                        <td><%=IndRslts(1, j)%></td>
                                        <td><%=IndRslts(2, j)%></td>
                                        <td><%=IndRslts(3, j)%></td>
                                        <td><%=IndRslts(4, j)%></td>
                                        <td><%=IndRslts(5, j)%></td>
                                        <td>
                                            <%If CSng(IndRslts(6, j)) = 30000 Then%>
                                                DNF
                                            <%Else%>
                                                <%=ConvertToMinutes(CSng(IndRslts(6, j)))%>
                                            <%End If%>
                                        </td>
                                    </tr>
                                <%Next%>         
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: right;" colspan="6">
                            Team Total Time: <%=TeamRslts(7, i)%><br>
                            Team Average Time: <%=TeamRslts(8, i)%>
                        </td>
                    </tr>
                <%Next%>
            </table>
        <%Else%>
            <h4 class="h4">Simple View</h4>
            <table class="table table-striped">
                <tr><th>PL</th><th>TEAM</th><th>CUMULATIVE</th><th>AVERAGE</th><th>SCORE</th><th>POINTS</th></tr>
                <%For i = 0 To UBound(TeamRslts, 2)%>
                    <tr>
                        <%If TeamRslts(3, i) = "---" Then%>
                            <td>---</td>
                        <%Else%>
                            <td><%=i + 1%></td>
                        <%End If%>
                        <td><%=TeamRslts(2, i)%></td>
                        <td><%=TeamRslts(3, i)%></td>
                        <td><%=TeamRslts(4, i)%></td>
                        <td><%=TeamRslts(5, i)%></td>
                        <td><%=TeamRslts(6, i)%></td>
                    </tr>
                <%Next%>
            </table>
        <%End If%>
    </div>
</div>
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
