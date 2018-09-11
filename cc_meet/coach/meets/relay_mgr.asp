<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k
Dim lTeamID, lMeetID, lRaceID
Dim iNumParts
Dim sMeetRaces, sGender, sMeetName, sTeamName, sRaceName, sShowNotes
Dim Roster(), Races(), RelayTeams(), TeamParts(4)
Dim dMeetDate, dShutdown
Dim bFound

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lTeamID = Request.QueryString("team_id")
lMeetID = Request.QueryString("meet_id")
lRaceID = Request.QueryString("race_id")

sShowNotes = Request.QueryString("show_notes")
If sShowNotes = vbNullString Then sShowNotes = "n"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name, date
sql = "SELECT MeetName, MeetDate, WhenShutdown FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value 
dMeetDate = rs(1).Value 
dShutdown = rs(2).Value
Set rs = Nothing

'get team name, gender
sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = rs(0).Value & " (" & rs(1).Value & ")"
Set rs = Nothing

sql = "SELECT Gender FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sGender = rs(0).Value
Set rs = Nothing
	
'convert gender to full word
Select Case sGender
	Case "M"
		sGender = "Male"
	Case "F"
		sGender = "Female"
End Select

i = 0
ReDim Roster(1, 0)
sql = "SELECT RosterID, FirstName, LastName, Gender FROM Roster WHERE TeamsID = " & lTeamID & " AND Archive = 'n' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Roster(0, i) = rs(0).Value
	Roster(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & GetGrade(rs(0).Value) & ")"
  	i = i + 1
	ReDim Preserve Roster(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID & " AND (Gender = '" & sGender & "' OR Gender = 'Open') AND IndivRelay = 'Relay'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If sMeetRaces = vbNullString Then 
        sMeetRaces = rs(0).Value & ", "
    Else
        sMeetRaces = sMeetRaces & rs(0).Value & ", "
    End If

	Races(0, i) = rs(0).Value
	Races(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Len(sMeetRaces) > 0 Then sMeetRaces = Left(sMeetRaces, Len(sMeetRaces) - 1)

If Request.Form.Item("get_race") = "get_race" Then
	lRaceID = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_legs") = "submit_legs" Then
    Call GetRelayTeams()

    iNumParts = GetNumParts()

    For i = 0 To UBound(RelayTeams, 2) - 1
       For j = 1 To iNumParts
           'get existing legs in the db for this team
            bFound = False
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT Leg FROM RelayMmbrs WHERE RelayTeamID = " & RelayTeams(0, i) & " AND RacesID = " & lRaceID & " AND Leg = " & j
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then bFound = True
            rs.Close
            Set rs = Nothing

            If Not Request.Form.Item("bib_" & RelayTeams(0, i) & "_" & j) = vbNullString Then
                If IsNumeric(Request.Form.Item("bib_" & RelayTeams(0, i) & "_" & j)) Then
                    If bFound = True Then
                        Set rs = Server.CreateObject("ADODB.Recordset")
                        sql = "SELECT RosterID, Bib FROM RelayMmbrs WHERE RelayTeamID = " & RelayTeams(0, i) & " AND RacesID = " & lRaceID & " AND Leg = " 
                        sql = sql & j
                        rs.Open sql, conn, 1, 2
                        rs(0).Value = Request.Form.Item("roster_id_" & RelayTeams(0, i) & "_" & j)
                        rs(1).Value = Request.Form.Item("bib_" & RelayTeams(0, i) & "_" & j)
                        rs.Update
                        rs.Close
                        Set rs = Nothing
                    Else
                        sql = "INSERT INTO RelayMmbrs(RelayTeamID, RacesID, RosterID, Bib, Leg) VALUES(" & RelayTeams(0, i) & ", " & lRaceID & ", "
                        sql = sql & Request.Form.Item("roster_id_" & RelayTeams(0, i) & "_" & j) & ", " 
                        sql = sql & Request.Form.Item("bib_" & RelayTeams(0, i) & "_" & j) & ", " & j & ")"
                        Set rs = conn.Execute(sql)
                        Set rs = Nothing
                    End If
                End If
            End if
        Next
    Next
End If

If CStr(lRaceID) = vbNullString Then lRaceID = 0

If Not CLng(lRaceID) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RaceDesc FROM Races WHERE RacesID = " & lRaceID
	rs.Open sql, conn, 1, 2
	sRaceName = Replace(rs(0).Value, "''", "'")
	rs.Close
	Set rs = Nothing

    iNumParts = GetNumParts()

    Call GetRelayTeams()
End If

Private Function GetNumParts()
    bFound = False
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT NumParts FROM RElays WHERE RacesID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0  Then
        GetNumParts = rs(0).Value
        bFound = True
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        GetNumParts = 4

        sql = "INSERT INTO Relays(RacesID) VALUES (" & lRaceID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
End Function

Private Sub GetRelayTeams()
    'get race parts (the teams themselves)
    i = 0
    ReDim RelayTeams(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT r.RosterID, r.FirstName, r.LastName, ir.Bib FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " 
    sql = sql & lRaceID & " AND r.TeamsID = " & lTeamID & " ORDER BY ir.Bib"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        RelayTeams(0, i) = rs(0).Value
        RelayTeams(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
        RelayTeams(2, i) = rs(3).Value
        i = i + 1
        ReDim Preserve RelayTeams(2, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
	
Private Function GetGrade(lMyID)
    GetGrade = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	If Month(Date) < 8 Then
        sql2 = "SELECT Grade" & Right(CStr(Year(Date) - 1), 2) & " FROM Grades WHERE RosterID = " & lMyID
    Else
        sql2 = "SELECT Grade" & Right(CStr(Year(Date)), 2) & " FROM Grades WHERE RosterID = " & lMyID
    End If
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetGrade = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function

'get relay team participants
Private Sub GetTeamParts(lRelayTeamID, iLeg)
    TeamParts(0) = 0
    TeamParts(1) = vbNullString
    Teamparts(2) = "00:00:00.000"
    Teamparts(3) = "00:00:00.000"
    Teamparts(4) = "00:00:00.000"

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID, Bib, StartTime, FinishTime, SplitTime FROM RelayMmbrs WHERE RelayTeamID = " & lRelayTeamID & " AND RacesID = " & lRaceID 
    sql = sql & " AND Leg = " & iLeg
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        TeamParts(0) = rs(0).Value
        TeamParts(1) = rs(1).Value
        Teamparts(2) = rs(2).Value
        TeamParts(3) = rs(3).Value
        Teamparts(4) = rs(4).Value
    End If 
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Cross-Country Coach Home</title>
<!--#include file = "../../../includes/js.asp" --> 
</head>

<body>
<div class="container">
	<h4 class="h4"><%=sTeamName%> Relay Team Manager for <%=sMeetName%> on <%=dMeetDate%></h4>

    <%If UBound(Races, 2) + 0 Then%>
        <%If CDate(dShutdown) + 2 > Now Or Session("role") = "admin" Then%>
		    <form role="form" class="form-inline" name="get_team" method="post" action="relay_mgr.asp?meet_id=<%=lMeetID%>&amp;team_id=<%=lTeamID%>&amp;show_notes=<%=sShowNotes%>">
		    <label for="races">Select Race:</label>
		    <select class="form-control" name="races" id="races" onchange="this.form.submit1.click();">
			    <option value="0">&nbsp;</option>
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
         
            <h5 class="h5">Important Notes About GSE Relays </h5>   

            <%If sShowNotes = "y" Then%>
                <div class="bg-info">
                    <a href="relay_mgr.asp?meet_id=<%=lMeetID%>&amp;team_id=<%=lTeamID%>&amp;race_id=<%=lRaceID%>&amp;show_notes=n">Hide Notes</a>
                </div>
                <ul class="list-group">
                    <li class="list-group-item">
                        The relay team MAY NOT be an actual individual member of the team but must be added to the roster!  Using a convention such as 
                        "Varsity A Mtka" (with "Mtka" being the last name) would work.
                    </li>
                    <li class="list-group-item">
                        The last relay participant (the anchor leg) MUST wear the bib assigned to the relay team itself and all four legs must be
                        represented by bibs in consecutive ascending order!
                    </li>
                    <li class="list-group-item">
                        Every relay team leg MUST have a bib assigned to it but does NOT have to have a pre-determined participant.  In fact, the actual participant can
                        be added or changed later once it is determined who actually participated.
                    </li>
                    <li class="list-group-item">
                        One individual can ski or run multiple legs on the same or different teams BUT MUST WEAR DIFFERENT BIBS EACH TIME.
                    </li>
                    <li class="list-group-item">
                        There can be NO DUPLICATE BIBS worn by different participants and no bib can be worn for more than one leg.
                    </li>
                    <li class="list-group-item">
                        A participants's actual time will be determined by the difference between their finish time (clock time) and the time of the person 
                        that competed before them.  It is up to the coaches and the meet management to ensure that the participants do not leave early.
                    </li>
                </ul>
            <%Else%>
                <div class="bg-info">
                    <a href="relay_mgr.asp?meet_id=<%=lMeetID%>&amp;team_id=<%=lTeamID%>&amp;race_id=<%=lRaceID%>&amp;show_notes=y">Show Notes</a>
                </div>
            <%End If%>

            <hr>

            <%If Not CLng(lRaceID) = 0 Then%>
                <h5 class="h5"><%=sRaceName%></h5>
            
                <form role="form" class="form" name="team_members" method="post" 
                    action="relay_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;show_notes=<%=sShowNotes%>">
                <%For i = 0 To UBound(RelayTeams, 2) - 1%>
                    <h5 class="h5"><%=RelayTeams(1, i)%> (Bib: <%=RelayTeams(2, i)%>)</h5>
                    <table class="table table-striped">
                        <tr><th>Leg</th><th style="text-align:left;">Participant</th><th>Bib</th><th>Start Time</th><th>Finish Time</th><th>Split Time</th></tr>
                        <%For j = 1 To iNumParts%>
                            <%Call GetTeamParts(RelayTeams(0, i), j)%>
                            <tr>
                                <th>
                                    <input type="text" class="form-control" name="leg_<%=RelayTeams(0, i)%>_<%=j%>" id="leg_<%=RelayTeams(0, i)%>_<%=j%>"
                                               value="<%=j%>" disabled>
                                </th>
                                <td>
                                    <select class="form-control" name="roster_id_<%=RelayTeams(0, i)%>_<%=j%>" id="roster_id_<%=RelayTeams(0, i)%>_<%=j%>">
                                        <option value="0">Undetermined</option>
                                        <%For k = 0 To UBound(Roster, 2) - 1%>
                                            <%If CLng(Roster(0, k)) = CLng(TeamParts(0)) Then%>
                                                <option value="<%=Roster(0, k)%>" selected><%=Roster(1, k)%></option>
                                            <%Else%>
                                                <option value="<%=Roster(0, k)%>"><%=Roster(1, k)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                                <td>
                                    <%If CInt(j) = CInt(iNumParts) Then%>
                                        <input type="text" class="form-control" name="display_bib_<%=RelayTeams(0, i)%>_<%=j%>" id="display_bib_<%=RelayTeams(0, i)%>_<%=j%>"
                                               value="<%=RelayTeams(2, i)%>" disabled>
                                        <input type="hidden" name="bib_<%=RelayTeams(0, i)%>_<%=j%>" id="bib_<%=RelayTeams(0, i)%>_<%=j%>"
                                               value="<%=RelayTeams(2, i)%>">
                                    <%Else%>
                                        <%If CDate(dShutdown) < Now Then%>
                                            <input type="text" class="form-control" name="locked_bib_<%=RelayTeams(0, i)%>_<%=j%>" id="locked_bib_<%=RelayTeams(0, i)%>_<%=j%>"
                                                value="<%=TeamParts(1)%>" disabled>
                                            <input type="hidden" name="bib_<%=RelayTeams(0, i)%>_<%=j%>" id="bib_<%=RelayTeams(0, i)%>_<%=j%>"
                                                value="<%=TeamParts(1)%>">
                                        <%Else%>
                                            <input type="text" class="form-control" name="bib_<%=RelayTeams(0, i)%>_<%=j%>" id="bib_<%=RelayTeams(0, i)%>_<%=j%>"
                                                value="<%=TeamParts(1)%>">
                                        <%End If%>
                                    <%End If%>
                                </td>
                                <td><%=TeamParts(2)%></td>
                                <td><%=TeamParts(3)%></td>
                                <td><%=TeamParts(4)%></td>
                            </tr>
                        <%Next%>
                    </table>
                <%Next%>
                <div class="form-group">
		            <input type="hidden" name="submit_legs" id="submit_legs" value="submit_legs">
		            <input type="submit" class="form-control" name="submit2" id="submit2" value="Submit Participants/Bibs">
                </div>
                </form>
            <%End If%>
        <%Else%>
            <p class="bg-warning">
                The time for entering relay teams in this meet has expired,  Please contact
                <a href="mailto:bob.schneider@gopherstateevents.com">Gopher State Events</a> with concerns or to inquire about entries.
            </p>
        <%End If%>
    <%Else%>
        <p class="bg-warning">
            None of the races in this meet are relays.  Please contact <a href="mailto:bob.schneider@gopherstateevents.com">Gopher State Events</a> 
            if you believe you have reached this message in error.
        </p>
    <%End If%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
