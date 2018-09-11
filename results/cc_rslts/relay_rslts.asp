<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k, m
Dim lMeetID, lRaceID
Dim sMeetName, sRaceName, sRaceGender, sDisplayBy, sPartName, sTeamName, sRelayName, sSport
Dim iNumParts
Dim RelayTeams(), TeamParts(4), MeetTeams(), Races(), Finishers()
Dim dMeetDate
Dim bFound

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

sDisplayBy = Request.QueryString("display_by")
If sDisplayBy = vbNullString Then sDisplayBy = "school"
'also by school or time

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("get_race") = "get_race" Then
	lRaceID = Request.Form.Item("races")
End If

sql = "SELECT r.RaceDesc, m.MeetsID, m.MeetName, m.MeetDate, r.Gender, m.Sport FROM Races r INNER JOIN Meets m ON r.MeetsID = m.MeetsID WHERE r.RacesID = " 
sql = sql & lRaceID
Set rs = conn.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'")
lMeetID = rs(1).Value
sMeetName = rs(2).Value 
dMeetDate = rs(3).Value
sRaceGender = rs(4).Value 
sSport = rs(5).Value
Set rs = Nothing

sql = "SELECT NumParts FROM Relays WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
If rs.BOF and rs.EOF Then
    bFound = False
Else
    iNumParts = rs(0).Value
    bFound = True
End If
Set rs = Nothing

If bFound = False Then
    iNumParts = 4

    sql = "INSERT INTO Relays(RacesID) VALUES (" & lRaceID & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If

'get teams for this race
i = 0
ReDim MeetTeams(1, 0)
sql = "SELECT mt.TeamsID, t.TeamName FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lMeetID & " AND t.Gender = '" & Left(sRaceGender, 1) & "' ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve MeetTeams(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim Races(1, 0)
sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID & " AND IndivRelay = 'Relay'"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub DisplayBySchool(lTeamID)
    'get race parts (the teams themselves)
    i = 0
    ReDim RelayTeams(4, 0)
    sql = "SELECT r.RosterID, r.FirstName, r.LastName, ir.Bib, ir.Place, ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " 
    sql = sql & lRaceID & " AND r.TeamsID = " & lTeamID & " AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        RelayTeams(0, i) = rs(0).Value
        RelayTeams(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
        RelayTeams(2, i) = rs(3).Value
        RelayTeams(3, i) = rs(4).Value
        RelayTeams(4, i) = rs(5).Value
        i = i + 1
        ReDim Preserve RelayTeams(4, i)
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub DisplayByPlace()
    'get race parts (the teams themselves)
    i = 0
    ReDim RelayTeams(4, 0)
    sql = "SELECT r.RosterID, r.FirstName, r.LastName, ir.Bib, ir.Place, ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " 
    sql = sql & lRaceID & " AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        RelayTeams(0, i) = rs(0).Value
        RelayTeams(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
        RelayTeams(2, i) = rs(3).Value
        RelayTeams(3, i) = rs(4).Value
        RelayTeams(4, i) = rs(5).Value
        i = i + 1
        ReDim Preserve RelayTeams(4, i)
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub DisplayBySplit()
    Dim x, y, z
    Dim SortArr(4)

    x = 0
    ReDim Finishers(4, 0)
    sql = "SELECT RosterID, Bib, SplitTime, RelayTeamID FROM RelayMmbrs WHERE RacesID = " & lRaceID & " AND SplitTime > '00:00'"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        Call GetRelayData(rs(0).Value, rs(3).Value)

        Finishers(0, x) = sPartName
        Finishers(1, x) = sTeamName
        Finishers(2, x) = rs(1).Value
        Finishers(3, x) = Round(ConvertToSeconds(rs(2).Value), 2)
        Finishers(4, x) = sRelayName

        x = x + 1
        ReDim Preserve Finishers(4, x)

        rs.MoveNext
    Loop 
    Set rs = Nothing

    'sort by time
    For x = 0 To UBound(Finishers, 2) - 2
        For y = x + 1 To UBound(Finishers, 2) - 1
            If CSng(Finishers(3, x)) > CSng(Finishers(3, y)) Then
                For z = 0 To 4
                    SortArr(z) = Finishers(z, x)    
                    Finishers(z, x) = Finishers(z, y)
                    Finishers(z, y) = SortArr(z)
                Next
            End If
        Next
    Next

    For x = 0 To UBound(Finishers, 2) - 1
        Finishers(3, x) = ConvertToMinutes(Finishers(3, x))
    Next
End Sub

'get relay team participants
Private Sub GetTeamParts(lRelayTeamID, iLeg)
    TeamParts(0) = 0
    TeamParts(1) = vbNullString
    Teamparts(2) = "00:00:00.000"
    Teamparts(3) = "00:00:00.000"
    Teamparts(4) = "00:00:00.000"

    sql = "SELECT RosterID, Bib, StartTime, FinishTime, SplitTime FROM RelayMmbrs WHERE RelayTeamID = " 
    sql = sql & lRelayTeamID & " AND RacesID = " & lRaceID & " AND Leg = " & iLeg
    Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
        '---
    Else
        TeamParts(0) = GetMyName(rs(0).Value)
        TeamParts(1) = rs(1).Value
        Teamparts(2) = rs(2).Value
        TeamParts(3) = rs(3).Value
        Teamparts(4) = rs(4).Value
    End If 
    Set rs = Nothing
End Sub

Private Function GetMyName(lRosterID)
    sql2 = "SELECT FirstName, LastName FROM Roster WHERE RosterID = " & lRosterID 
    Set rs2 = conn.Execute(sql2)
    If rs.BOF and rs.EOF Then
        GetMyName = "Undetermined"
    Else
        GetMyName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    End If    
    Set rs2 = Nothing
End Function

Private Sub GetRelayData(lRosterID, lRelayID)
    sql2 = "SELECT FirstName, LastName FROM Roster WHERE RosterID = " & lRosterID 
    Set rs2 = conn.Execute(sql2)
    If rs.BOF and rs.EOF Then
        sPartName = "Undetermined"
    Else
        sPartName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    End If    
    Set rs2 = Nothing

    sql2 = "SELECT FirstName, LastName FROM Roster WHERE RosterID = " & lRelayID 
    Set rs2 = conn.Execute(sql2)
    If rs.BOF and rs.EOF Then
        sRelayName = "Undetermined"
    Else
        sRelayName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    End If    
    Set rs2 = Nothing

    sql2 = "SELECT t.TeamName FROM Roster r INNER JOIN Teams t ON r.TeamsID = t.TeamsID WHERE r.RosterID = " & lRelayID 
    Set rs2 = conn.Execute(sql2)
    If rs.BOF and rs.EOF Then
        sTeamName = "Undetermined"
    Else
        sTeamName = Replace(rs2(0).Value, "''", "'")
    End If    
    Set rs2 = Nothing
End Sub

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->

<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Relay Results</title>
<meta name="description" content="GSE cross-country/nordic ski relay results.">
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">

	<div class="bg-warning">
		<a href="javascript:window.print();">Print</a>
	</div>

	<h4 class="h4">Relay Results for <%=sMeetName%> on <%=dMeetDate%></h4>

	<form role="form" class="form-inline" name="get_team" method="post" action="relay_rslts.asp?race_id=<%=lRaceID%>&amp;display_by=<%=sDisplayBy%>">
	<label for="races">Select Race:</label>
	<select class="form-control" name="races" id="races" onchange="this.form.submit1.click();">
		<%For i = 0 to UBound(Races, 2) - 1%>
			<%If CLng(Races(0, i)) = CLng(lRaceID) Then%>
				<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
			<%Else%>
				<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
			<%End If%>
		<%Next%>
	</select>
	<input type="hidden" name="get_race" id="get_race" value="get_race">
	<input type="submit" class="form-control" name="submit1" id="submit1" value="Get This Race">
	</form>

    <ul class="list-inline">
        <li class="list-group-item">Display By:</li>
        <li class="list-group-item"><a href="relay_rslts.asp?race_id=<%=lRaceID%>&amp;display_by=school">School</a></li>
        <li class="list-group-item"><a href="relay_rslts.asp?race_id=<%=lRaceID%>&amp;display_by=place">Place</a></li>
        <li class="list-group-item"><a href="relay_rslts.asp?race_id=<%=lRaceID%>&amp;display_by=split_time">Split Time</a></li>
    </ul>
    
    <%Select Case sDisplayBy%>
        <%Case "school"%>
            <h4 class="h4">Displayed by School</h4>
            <%For m = 0 To UBound(MeetTeams, 2) - 1%>  
                <%Call DisplayBySchool(MeetTeams(0, m))%>

                <%If UBound(RelayTeams, 2) > 0 Then%>
                    <h5 class="h5"><%=MeetTeams(1, m)%></h5>   

                    <%For i = 0 To UBound(RelayTeams, 2) - 1%>
                        <h5 class="h5"><%=RelayTeams(1, i)%> (Bib: <%=RelayTeams(2, i)%>)&nbsp;&nbsp;<%=RelayTeams(4, i)%></h5>

                        <table class="table table-striped">
                            <tr><th>Leg</th><th>Participant</th><th>Bib</th><th>Start Time</th><th>Finish Time</th><th>Split Time</th></tr>
                            <%For j = 1 To iNumParts%>
                                <%Call GetTeamParts(RelayTeams(0, i), j)%>
                                <tr>
                                    <th><%=j%></th>
                                    <td><%=TeamParts(0)%></td>
                                    <td><%=TeamParts(1)%></td>
                                    <td><%=TeamParts(2)%></td>
                                    <td><%=TeamParts(3)%></td>
                                    <td><%=TeamParts(4)%></td>
                                </tr>
                            <%Next%>
                        </table>
                    <%Next%>
                <%End If%>
            <%Next%>
        <%Case "place"%>
            <h4 class="h4">Displayed by Place</h4>
        
            <%Call DisplayByPlace()%>

            <%For i = 0 To UBound(RelayTeams, 2) - 1%>
                <h5 class="h5"><%=RelayTeams(1, i)%> (Bib: <%=RelayTeams(2, i)%>)&nbsp;&nbsp;<%=RelayTeams(4, i)%></h5>
                <table class="table table-striped">
                    <tr><th>Leg</th><th>Participant</th><th>Bib</th><th>Start Time</th><th>Finish Time</th><th>Split Time</th></tr>
                    <%For j = 1 To iNumParts%>
                        <%Call GetTeamParts(RelayTeams(0, i), j)%>
                        <tr>
                            <th><%=j%></th>
                            <td><%=TeamParts(0)%></td>
                            <td><%=TeamParts(1)%></td>
                            <td><%=TeamParts(2)%></td>
                            <td><%=TeamParts(3)%></td>
                            <td><%=TeamParts(4)%></td>
                        </tr>
                    <%Next%>
                </table>
            <%Next%>
        <%Case "split_time"%>
            <h4 class="h4">Displayed by Split Time</h4>
        
            <%Call DisplayBySplit()%>

            <%If sSport = "Nordic Ski" Then%>
                <p class="bg-danger">IMPORTANT NOTE:  If different techniques were used by different participants, that will affect the validity of
                these rankings.</p>
            <%End If%>

            <table class="table table-striped">
                <tr>
                    <th>No.</th>
                    <th>Participant</th>
                    <th>School</th>
                    <th>Bib</th>
                    <th>Split Time</th>
                    <th>Relay Team</th>
                </tr>
                <%For i = 0 To UBound(Finishers, 2) - 1%>
                    <tr>
                        <th><%=i%></th>
                        <td><%=Finishers(0, i)%></td>
                        <td><%=Finishers(1, i)%></td>
                        <td><%=Finishers(2, i)%></td>
                        <td><%=Finishers(3, i)%></td>
                        <td><%=Finishers(4, i)%></td>
                    </tr>
                <%Next%>
            </table>
    <%End Select%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
