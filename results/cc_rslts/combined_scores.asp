<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisRace, lThisTeam, lMeetID
Dim i, j, k
Dim sIncrDecr, sThisTeam, sMeetName, sRacesToCombine, sRaceNames, sSport
Dim iNumRaces
Dim sngTotalPts
Dim RacesToCombine(), CombinedScores(), MeetTeams(), Races
Dim dMeetDate
   
lMeetID = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate, Sport FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
sSport = rs(2).Value
Set rs = Nothing
	
sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
Races = rs.GetRows()
Set rs = Nothing

ReDim RacesToCombine(2, 0)

If Request.Form.Item("submit_races") = "submit_races" Then
    sIncrDecr = Request.Form.Item("incr_decr")

	sRacesToCombine = Request.Form.Item("races")

    j = 0
    For i = 1 To Len(sRacesToCombine)
        If Mid(sRacesToCombine, i, 1) = "," Then
            RacesToCombine(0, j) = lThisRace
            j = j + 1
            ReDim Preserve RacesToCombine(2, j)

            lThisRace = vbNullString
        Else
            lThisRace = lThisRace & Mid(sRacesToCombine, i, 1)
            If i = Len(sRacesToCombine) Then 
                RacesToCombine(0, j) = lThisRace
                j = j + 1
                ReDim Preserve RacesToCombine(2, j)
            End If
        End If
    Next
End If

If sIncrDecr = vbNullString Then
    If sSport = "Nordic Ski" Then
        sIncrDecr = "Decreasing"
    Else 
        sIncrDecr = "Increasing"
    End If
End If

If UBound(RacesToCombine, 2) > 0 Then
    'get race names
    For i = 0 To UBound(RacesToCombine)
        sql = "SELECT RaceName, RaceDesc FROM Races WHERE RacesID = " & RacesToCombine(0, i)
        Set rs = conn.Execute(sql)
        RacesToCombine(1, i) = Replace(rs(0).Value, "''", "'")
        RacesToCombine(2, i) = Replace(rs(1).Value, "''", "'")
        Set rs = Nothing
    Next

    iNumRaces = UBound(RacesToCombine, 2)
End If

'get meet teams
i = 0
ReDim MeetTeams(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT mt.TeamsID, t.TeamName FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID WHERE mt.MeetsID = " & lMeetID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    MeetTeams(0, i) = rs(0).Value
    MeetTeams(1, i) = Replace(rs(1).Value, "''", "'")
    i = i + 1
    ReDim Preserve MeetTeams(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
    
k = 0
ReDim CombinedScores(iNumRaces + 1, 0)
For j = 0 To UBound(MeetTeams, 2) - 1
    lThisTeam = MeetTeams(0, j)
    sngTotalPts = 0
        
    CombinedScores(0, k) = MeetTeams(1, j) 'team name
    'get score for each race for this team
    For i = 0 To iNumRaces - 1
        lThisRace = RacesToCombine(0, i)
        CombinedScores(i + 1, k) = GetTeamScore(lThisRace, lThisTeam)
        sngTotalPts = CSng(sngTotalPts) + CSng(CombinedScores(i + 1, k))
    Next
        
    'enter total points
    CombinedScores(iNumRaces + 1, k) = sngTotalPts
        
    k = k + 1
    ReDim Preserve CombinedScores(iNumRaces + 1, k)
Next
    
Dim TempArr()
ReDim TempArr(iNumRaces + 1)

'sort by total score
For i = 0 To UBound(CombinedScores, 2) - 1
    For j = 0 To UBound(CombinedScores, 2)
        If sIncrDecr = "Decreasing" Then
            If CSng(CombinedScores(iNumRaces + 1, i)) > CSng(CombinedScores(iNumRaces + 1, j)) Then
                For k = 0 To iNumRaces + 1
                    TempArr(k) = CombinedScores(k, i)
                    CombinedScores(k, i) = CombinedScores(k, j)
                    CombinedScores(k, j) = TempArr(k)
                Next
            End If
        Else
            If CombinedScores(iNumRaces + 1, i) < CombinedScores(iNumRaces + 1, j) Then
                For k = 0 To iNumRaces + 1
                    TempArr(k) = CombinedScores(k, i)
                    CombinedScores(k, i) = CombinedScores(k, j)
                    CombinedScores(k, j) = TempArr(k)
                Next
            End If
        End If
    Next
Next

Private Function GetTeamScore(lThisRaceID, lThisTeamID)
    Dim rs2, sql2
    
    GetTeamScore = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Score FROM TmRslts WHERE RacesID = " & lThisRaceID & " AND TeamsID = " & lThisTeamID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetTeamScore = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CC/Nordic Results Manager: Combined Team Scores</title>
<<meta name="description" content="Cross-Country & Nordic Ski Results by Gopher State Events, a conventional timing service offererd by H51 Software, LLC in Minnetonka, MN.">
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
	<div><a href="javascript:window.print();">Print</a></div>

	<h4 class="h4">Combined Team Scores for <%=sMeetName%> on <%=dMeetDate%></h4>

    <form class="form-inline" name="get_races" method="post" action="combined_scores.asp?meet_id=<%=lMeetID%>">
    <label for="races">Races:</label>
    <select class="form-control" name="races" id="races" multiple size="<%=UBound(Races, 2)%>">
        <%For i = 0 to UBound(Races, 2)%>
            <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
        <%Next%>
    </select>
    <label for="gender">Sort By:</label>
    <select class="form-control" name="incr_decr" id="incr_decr" multiple size="2">
        <%If sIncrDecr = "Decreasing" Then%>
            <option value="Decreasing" selected>Decreasing</option>
            <option value="Increasing">Increasing</option>
        <%Else%>
            <option value="Decreasing">Decreasing</option>
            <option value="Increasing" selected>Increasing</option>
        <%End If%>
    </select>
    <input type="hidden" class="form-control" name="submit_races" id="submit_races" value="submit_races">
    <input type="submit" class="form-control" name="get_races" id="get_races" value="Combine These">
    </form>
	<br>
    <div class="bg-danger">
        Please note:  You can not combine scores across genders since a boys team and a girls team are separate teams.
        For instance, it can be used for combining scores in Classical and Freestyle
        techniques in Nordic Skiing and across multiple races in the same event.
    </div>

    <%If Not sRacesToCombine = vbNullString Then%>				
        <h4 class="h4">Scores</h4>  
	    <table class="table table-striped">
            <tr>
                <th>PL</th>
                <th>SCHOOL</th>
                <%For i = 0 To iNumRaces - 1%>
                    <th><%=UCase(RacesToCombine(2, i))%></th>
                <%Next%>
                <th>SCORE</th>
            </tr>
            <%k = 1%>
            <%For i = 0 To UBound(CombinedScores, 2)%>
                <%If CombinedScores(iNumRaces + 1, i) > 0 Then%>
                    <tr>
                        <td><%=k%></td>
                        <td><%=CombinedScores(0, i)%></td>
                        <%For j = 0 To iNumRaces - 1%>
                            <td><%=CombinedScores(j + 1, i)%></td>
                        <%Next%>
                        <td><%=CombinedScores(iNumRaces + 1, i)%></td>
                    </tr>

                    <%k = k + 1%>
                <%End If%>
            <%Next%>
        </table>
    <%End If%>
</div>
<!--#include file = "../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
