<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k, m
Dim lMeetID, lRaceID
Dim sMeetName, sRaceName, sRaceGender
Dim iNumParts
Dim RelayTeams(), TeamParts(4), MeetTeams(), Races()
Dim dMeetDate
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("get_race") = "get_race" Then
	lRaceID = Request.Form.Item("races")
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.RaceDesc, m.MeetsID, m.MeetName, m.MeetDate, r.Gender FROM Races r INNER JOIN Meets m ON r.MeetsID = m.MeetsID WHERE r.RacesID = " & lRaceID
rs.Open sql, conn, 1, 2
sRaceName = Replace(rs(0).Value, "''", "'")
lMeetID = rs(1).Value
sMeetName = rs(2).Value 
dMeetDate = rs(3).Value
sRaceGender = rs(4).Value 
rs.Close
Set rs = Nothing

bFound = False
Set rs = SErver.CreateObject("ADODB.Recordset")
sql = "SELECT NumParts FROM Relays WHERE RacesID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0  Then
    iNumParts = rs(0).Value
    bFound = True
End If
rs.Close
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
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID & " AND IndivRelay = 'Relay'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetRelayTeams(lTeamID)
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
        TeamParts(0) = GetMyName(rs(0).Value)
        TeamParts(1) = rs(1).Value
        Teamparts(2) = rs(2).Value
        TeamParts(3) = rs(3).Value
        Teamparts(4) = rs(4).Value
    End If 
    rs.Close
    Set rs = Nothing
End Sub

Private Function GetMyName(lRosterID)
    GetMyName = "Undetermined"
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName FROM Roster WHERE RosterID = " & lRosterID 
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetMyName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE Cross-Country Coach Home</title>
<!--#include file = "../../includes/meta2.asp" -->




<style type="text/css">
    td,th{padding-right: 5px;}
    li{
        padding-top: 5px;
    }
</style>
</head>

<body style="background-color: #ececd8;">
<div style="margin: 10px;padding: 10px;background-color: #fff;width: 550px;">
	<h4 class="h4">Relay Teams for <%=sMeetName%> on <%=dMeetDate%></h4>

    <div style="text-align: left;margin: 0;padding: 0;font-size: 0.8em;">
		<form name="get_team" method="post" action="relay_teams.asp?race_id=<%=lRaceID%>&amp;meet_id=<%=lMeetID%>">
		<span style="font-weight:bold;">Select Race:</span>&nbsp;
		<select name="races" id="races" onchange="this.form.submit1.click();">
			<%For i = 0 to UBound(Races, 2) - 1%>
				<%If CLng(Races(0, i)) = CLng(lRaceID) Then%>
					<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
				<%Else%>
					<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
				<%End If%>
			<%Next%>
		</select>
		<input type="hidden" name="get_race" id="get_race" value="get_race">
		<input type="submit" name="submit1" id="submit1" value="Get This Race">
		</form>
    </div>
            
    <h5 style="background: none;color: #039;text-align: left;margin-top: 10px;"><%=sRaceName%></h5>
    
    <%For m = 0 To UBound(MeetTeams, 2) - 1%>  
        <h4 style="text-align: left;padding-left: 5px;"><%=MeetTeams(1, m)%></h4>   
        
        <%Call GetRelayTeams(MeetTeams(0, m))%>

        <%For i = 0 To UBound(RelayTeams, 2) - 1%>
            <h5 style="text-align: left;margin-top: 10px;font-size: 0.8em;background-color: #ececd8;"><%=RelayTeams(1, i)%> (Bib: <%=RelayTeams(2, i)%>)</h5>
            <table style="margin-bottom: 10px;font-size: 0.8em;">
                <tr><th>Leg</th><th style="text-align:left;">Participant</th><th>Bib</th><th>Start Time</th><th>Finish Time</th><th>Split Time</th></tr>
                <%For j = 1 To iNumParts%>
                    <%Call GetTeamParts(RelayTeams(0, i), j)%>
                    <tr>
                        <th><%=j%></th>
                        <td style="text-align:left;"><%=TeamParts(0)%></td>
                        <td><%=TeamParts(1)%></td>
                        <td><%=TeamParts(2)%></td>
                        <td><%=TeamParts(3)%></td>
                        <td><%=TeamParts(4)%></td>
                    </tr>
                <%Next%>
            </table>
        <%Next%>
    <%Next%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
