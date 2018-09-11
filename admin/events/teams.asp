<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lRaceID, lEventID, lTeamID, lParticipantID, lAgeGrpID
Dim sEventName, sNewGender, sNewAgeGrp, sNewTeamName, sMale, sFemale, sCombined, sGender, sMmbrName
Dim sTeamName, sTeamGender, sTeamAgeGroup, sScoreMethod, sNewAgeGrp2, sAgeGrp
Dim iAge, iMinScore, iMaxScore, iNewEndAge, iEndAge
Dim dEventDate
Dim RaceArray(), Genders(2, 1), AgeGroups(), Teams(), TeamMembers(), AvailParts(), TeamGenders(2)
Dim ScoreMeth(3)
Dim bFound

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
lRaceID = Request.QueryString("race_id")
lTeamID = Request.QueryString("team_id")

lAgeGrpID = Request.QueryString("age_grp_id")
If CStr(lAgeGrpID) = vbNullString Then lAgeGrpID = 0

Genders(0, 0) = "Combined"
Genders(0, 1) = "y"
Genders(1, 0) = "Male"
Genders(1, 1) = "n"
Genders(2, 0) = "Female"
Genders(2, 1) = "n"

ScoreMeth(0) = "min score"
ScoreMeth(1) = "max score"
ScoreMeth(2) = "time"
ScoreMeth(3) = "average"

TeamGenders(0) = "Combined"
TeamGenders(1) = "Male"
TeamGenders(2) = "Female"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

i = 0
ReDim RaceArray(1, 0)	
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RaceArray(0, i) = rs(0).value
	RaceArray(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve RaceArray(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If UBound(RaceArray, 2) = 1 Then 
    lRaceID = RaceArray(0, 0)
Else
    If CStr(lRaceID) = vbNullString Then lRaceID = RaceArray(0, 0)
End If

Dim Events
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_edit_age_grp") = "submit_edit_age_grp" Then
    sAgeGrp = Request.Form.Item("age_grp_name")
    iEndAge = Request.Form.Item("end_age")

    If Request.Form.Item("delete_age_grp") = "y" Then
        sql = "DELETE FROM TeamAgeGrps WHERE TeamAgeGrpsID = " & lAgeGrpID
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AgeGrp, EndAge FROM TeamAgeGrps WHERE TeamAgeGrpsID = " & lAgeGrpID
        rs.Open sql, conn, 1, 2
        rs(0).Value = sAgeGrp
        rs(1).Value = iEndAge
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_age_grp") = "submit_age_grp" Then
    lAgeGrpID = Request.Form.Item("age_grp")
ElseIf Request.Form.Item("submit_new_age_grp") = "submit_new_age_grp" Then
    sNewAgeGrp2 = Request.Form.Item("new_age_grp2")
    iNewEndAge = Request.Form.Item("new_end_age")

    sql = "INSERT INTO TeamAgeGrps (RaceID, AgeGrp, EndAge) VALUES (" & lRaceID & ", '" & sNewAgeGrp2 & "', " 
    sql = sql & iNewEndAge & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_scoring") = "submit_scoring" Then
    iMaxScore = Request.Form.Item("max_score")
    iMinScore = Request.Form.Item("min_score")
    sScoreMethod = Request.Form.Item("score_method")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MaxScore, MinScore, ScoreMethod FROM TeamScoring WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    rs(0).Value = iMaxScore
    rs(1).Value = iMinScore
    rs(2).Value = sScoreMethod
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("edit_team") = "edit_team" Then
    If Request.Form.Item("delete_team") = "on" Then
        sql = "DELETE FROM Teams WHERE RaceID = " & lRaceID & " AND TeamName = '" & lTeamID & "'"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT TeamName, Gender, AgeGroup FROM Teams WHERE TeamsID = " & lTeamID
        rs.Open sql, conn, 1, 2
        rs(0).Value = Replace(Request.Form.Item("team_name"), "''", "")
        rs(1).Value = Request.Form.Item("team_gender")
        rs(2).Value = Request.Form.Item("team_age_group")
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("delete_part") = "delete_part" Then
    lParticipantID = Request.Form.Item("team_parts")
    
    sql = "DELETE FROM TeamMmbrs WHERE ParticipantID = " & lParticipantID & " AND TeamsID = " & lTeamID
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("select_part") = "select_part" Then
    lParticipantID = Request.Form.Item("participants")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT LastName, FirstName, Gender FROM Participant WHERE ParticipantID = " & lParticipantID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        sMmbrName = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        sGender = rs(2).Value
    Else    
        sMmbrName = "Unknown"
        sGender = "Unknown"
    End If
    rs.Close
    Set rs = Nothing
    
    'get participant age
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Age FROM PartRace WHERE ParticipantID = " & lParticipantID & " AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    iAge = rs(0).Value
    rs.Close
    Set rs = Nothing
    
    sql = "INSERT INTO TeamMmbrs (ParticipantID, TeamsID, MmbrName, Gender, Age) VALUES (" & lParticipantID & ", "
    sql = sql & lTeamID & ", '" & Replace(sMmbrName, "'", "''") & "', '" & sGender & "', " & iAge & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("select_team") = "select_team" Then
	lTeamID = Request.Form.Item("teams")
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("create_team") = "create_team" Then
    sNewAgeGrp = Request.Form.Item("new_age_grp")
    sNewGender = Request.Form.Item("new_gender")
    sNewTeamName = Request.Form.Item("new_team_name")
    
    sql = "INSERT INTO Teams (RaceID, TeamName, Gender, AgeGroup) VALUES (" & lRaceID & ", '"  & Replace(sNewTeamName, "'", "") & "', '" 
    sql = sql & sNewGender & "', '" & sNewAgeGrp & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_team") = "submit_team" Then
    lTeamID = Request.Form.Item("teams")
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_genders") = "submit_genders" Then
    sMale = "n"
    sFemale = "n"
    sCombined = "n"

    If Request.Form.Item("male") = "on" Then sMale = "y"
    If Request.Form.Item("female") = "on" Then sFemale = "y"
    If Request.Form.Item("combined") = "on" Then sCombined = "y"

    bFound = False
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Male, Female, Combined FROM TeamGenders WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs(0).Value = sMale
        rs(1).Value = sFemale
        rs(2).Value = sCombined
        rs.Update
        bFound = True
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO TeamGenders (RaceID, Male, Female, Combined) VALUES (" & lRaceID & ", '" & sMale & "', '"
        sql = sql & sFemale & "', '" & sCombined & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
End If

If CStr(lTeamID) = vbNullString Then lTeamID = 0

i = 0
ReDim AgeGroups(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamAgeGrpsID, AgeGrp FROM TeamAgeGrps WHERE RaceID = " & lRaceID & " ORDER BY EndAge"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    AgeGroups(0, i) = rs(0).Value
    AgeGroups(1, i) = rs(1).Value
    i = i + 1
    ReDim Preserve AgeGroups(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Combined, Male, Female FROM TeamGenders WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    Genders(0, 1) = rs(0).Value
    Genders(1, 1) = rs(1).Value
    Genders(2, 1) = rs(2).Value
End If
rs.Close
Set rs = Nothing

i = 0
ReDim Teams(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamName FROM Teams WHERE RaceID = " & lRaceID & " ORDER BY TeamName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Teams(0, i) = rs(0).Value
    Teams(1, i) = Replace(rs(1).Value, "''", "'")
    i = i + 1
    ReDim Preserve Teams(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get age group settings
If CLng(lAgeGrpID) > 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AgeGrp, EndAge FROM TeamAgeGrps WHERE TeamAgeGrpsID = " & lAgeGrpID
    rs.Open sql, conn, 1, 2
    sAGeGrp = rs(0).Value
    iEndAge = rs(1).Value
    rs.Close
    Set rs = Nothing
End If

'default values if no record exists
iMinScore = 2
iMaxScore = 2
sScoreMethod = "average"

bFound = False
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MaxScore, MinScore, ScoreMethod FROM TeamScoring WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    iMaxScore = rs(0).Value
    iMinScore = rs(1).Value
    sScoreMethod = rs(2).Value
    rs.Update
    
    bFound = True
End If
rs.Close
Set rs = Nothing

If bFound = False Then
    sql = "INSERT INTO TeamScoring (RaceID, MaxScore, MinScore, ScoreMethod) VALUES (" & lRaceID & ", " & iMaxScore
    sql = sql & ", " & iMinScore & ", '" & sScoreMethod & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If

If CLng(lTeamID) > 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamName, Gender, AgeGroup FROM Teams WHERE RaceID = " & lRaceID & " AND TeamsID = " & lTeamID
    rs.Open sql, conn, 1, 2
    sTeamName = Replace(rs(0).Value, "''", "'")
    sTeamGender = rs(1).Value
    sTeamAgeGroup = rs(2).Value
    rs.Close
    Set rs = Nothing

    i = 0
    ReDim RaceParts(2, 0)
    Set rs = Server.CreateObject("ADODB.REcordset")
    sql = "SELECT p.ParticipantID, p.LastName, p.FirstName, pr.Bib FROM PartRace pr INNER JOIN Participant p "
    sql = sql & "ON pr.ParticipantID = p.ParticipantID WHERE pr.RaceID = " & lRaceID & " ORDER BY p.LastName, p.FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF 
        RaceParts(0, i) = rs(0).Value
        RaceParts(1, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
        RaceParts(2, i) = rs(3).Value
        i = i + 1
        ReDim Preserve RaceParts(2, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    i = 0
    ReDim TeamMembers(1, 0)
    Set rs = Server.CreateObject("ADODB.REcordset")
    sql = "SELECT ParticipantID, MmbrName FROM TeamMmbrs WHERE TeamsID = " & lTeamID & " ORDER BY MmbrName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        TeamMembers(0, i) = rs(0).Value 
        TeamMembers(1, i) = Replace(rs(1).Value, "''", "'") & " (" & GetMyBib(rs(0).Value, lRaceID) & ")"
        i = i + 1
        ReDim Preserve TeamMembers(1, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    j = 0
    ReDim AvailParts(1, 0)
    
    For i = 0 To UBound(RaceParts, 2) - 1
        bFound = False
        Set rs = Server.CreateObject("ADODB.REcordset")
        sql = "SELECT tm.ParticipantID FROM TeamMmbrs tm INNER JOIN Teams t ON tm.TeamsID = t.TeamsID WHERE t.RaceID = " & lRaceID
        sql = sql & " AND tm.ParticipantID = " & RaceParts(0, i)
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then bFound = True
        rs.Close
        Set rs = Nothing
        
        If bFound = False Then 
            AvailParts(0, j) = RaceParts(0, i) 
            AvailParts(1, j) = Replace(RaceParts(1, i), "''", "") & " (" & RaceParts(2, i) & ")"
            j = j + 1
            ReDim Preserve AvailParts(1, j)
        End If
    Next
End If

Private Function GetMyBib(lThisPart, lMyRaceID)
    GetMyBib = 0
    
    Set rs2 = Server.CreateObject("ADODB.REcordset")
    sql2 = "SELECT Bib FROM PartRace WHERE ParticipantID = " & lThisPart & " AND RaceID = " & lMyRaceID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetMyBib = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Team Data</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h3 class="h3">Team Info:&nbsp;<%=sEventName%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="teams.asp?event_id=<%=lEventID%>">
			<label for="events">Events:</label>
			<select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Events, 2)%>
					<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%Else%>
						<option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			<input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
			</form>
			<br>

			<%If Not Clng(lEventID) = 0 Then%>
			    <!--#include file = "../../includes/event_nav.asp" -->
				
			    <%If UBound(RaceArray, 2) > 1 Then%>
				    <form class="form-inline" name="get_races" method="post" action="teams.asp?event_id=<%=lEventID%>">
                    <label for="races">Select Race:</label>
                    <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
                        <option value="">&nbsp;</option>
                        <%For i = 0 to UBound(RaceArray, 2) - 1%>
                            <%If CLng(lRaceID) = CLng(RaceArray(0, i)) Then%>
                                <option value="<%=RaceArray(0, i)%>" selected><%=RaceArray(1, i)%></option>
                            <%Else%>
                                <option value="<%=RaceArray(0, i)%>"><%=RaceArray(1, i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                    <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
                    <input class="form-control" type="submit" name="get_race" id="get_race" value="Get Race Info">
				    </form>
			    <%End If%>
			
			    <%If Not CLng(lRaceID) = 0 Then%>
                    <br>
                    <div class="row">
                        <div class="col-sm-3">
                            <h5 class="h5">Scoring</h5>

                            <form class="form" name="scoring" method="Post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
                            <div class="form-group">
                                <label for="score_method">Scoring Method</label>
                                <select class="form-control" name="score_method" id="score_method">
                                    <%For i = 0 To UBound(ScoreMeth)%>
                                        <%If CStr(sScoreMethod) = CStr(ScoreMeth(i)) Then%>
                                            <option value="<%=ScoreMeth(i)%>" selected><%=ScoreMeth(i)%></option>
                                        <%Else%>
                                            <option value="<%=ScoreMeth(i)%>"><%=ScoreMeth(i)%></option>
                                        <%End If%>
                                    <%Next%>
                                </select>
                            </div>
                            <div class="form-group">
                                <label for="min_score">Min Num Score</label>
                                <select class="form-control" name="min_score" id="min_score">
                                    <%For i = 0 To 100%>
                                        <%If CInt(iMinScore) = CInt(i) Then%>
                                            <option value="<%=i%>" selected><%=i%></option>
                                        <%Else%>
                                            <option value="<%=i%>"><%=i%></option>
                                        <%End If%>
                                    <%Next%>
                                </select>
                            </div>
                            <div class="form-group">
                                <label for="max_score">Max Num Score</label>
                                <select class="form-control" name="max_score" id="max_score">
                                    <%For i = 0 To 100%>
                                        <%If CInt(iMaxScore) = CInt(i) Then%>
                                            <option value="<%=i%>" selected><%=i%></option>
                                        <%Else%>
                                            <option value="<%=i%>"><%=i%></option>
                                        <%End If%>
                                    <%Next%>
                                </select>
                            </div>
                            <input type="hidden" name="submit_scoring" id="submit_scoring" value="submit_scoring">
                            <input type="submit" class="form-control" name="save_scoring" id="save_scoring" value="Save Changes">
                            </form>
                        </div>
                        <div class="col-sm-3">
                            <h5 class="h5">Add Age Group</h5>

                            <form class="form" name="new_age_group" method="Post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
                            <div class="form-group">
                                <label for="new_age_grp2">Age Group Name:</label>
                                <input class="form-control" type="text" name="new_age_grp2" id="new_age_grp2">
                            </div>
                            <div class="form-group">
                                <label for="new_end_age">End Age</label>
                                <select class="form-control" name="new_end_age" id="new_end_age">
                                    <%For i = 5 To 110%>
                                        <option value="<%=i%>"><%=i%></option>
                                    <%Next%>
                                </select>
                            </div>
                            <input type="hidden" name="submit_new_age_grp" id="submit_new_age_grp" value="submit_new_age_grp">
                            <input type="submit" class="form-control" name="save_age_grp" id="save_age_grp" value="Save Age Group">
                            </form>
                            <br>
                           <h5 class="h5">Genders</h5>
                            <form class="form" name="new_team" method="post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
                            <div class="form-group">
                                <%If Genders(0, 1) = "n" Then%>
                                    &nbsp;&nbsp;<input type="checkbox" name="combined" id="combined">&nbsp;Combined <br>
                                <%Else%>
                                    &nbsp;&nbsp;<input type="checkbox" name="combined" id="combined" checked>&nbsp;Combined <br>
                                <%End If%>
                                <%If Genders(1, 1) = "n" Then%>
                                    &nbsp;&nbsp;<input type="checkbox" name="male" id="male">&nbsp;Male <br>
                                <%Else%>
                                    &nbsp;&nbsp;<input type="checkbox" name="male" id="male" checked>&nbsp;Male <br>
                                <%End If%>
                                <%If Genders(2, 1) = "n" Then%>
                                    &nbsp;&nbsp;<input type="checkbox" name="female" id="female">&nbsp;Female <br>
                                <%Else%>
                                    &nbsp;&nbsp;<input type="checkbox" name="female" id="female" checked>&nbsp;Female <br>
                                <%End If%>
                            </div>
                            <div class="form-group">
                                <input type ="hidden" name="submit_genders" id="submit_genders" value="submit_genders">
                                <input class="form-control" type="submit" name="this_gender" id="this_gender" value="Edit Genders">
                            </div>
                            </form>
                        </div>
                        <div class="col-sm-3">
                            <h5 class="h5">View/Edit Age Groups</h5>

                            <form class="form" name="get_age_grp" method="post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
                            <div class="form-group">
                                <label for="age_grp">Select Age Group:</label>
                                <select class="form-control" name="age_grp" id="age_grp" onchange="this.form.get_age_grp.click()">
                                    <option value="">&nbsp;</option>
                                    <%For i = 0 to UBound(AgeGroups, 2) - 1%>
                                        <%If CLng(lAgeGrpID) = CLng(AgeGroups(0, i)) Then%>
                                            <option value="<%=AgeGroups(0, i)%>" selected><%=AgeGroups(1, i)%></option>
                                        <%Else%>
                                            <option value="<%=AgeGroups(0, i)%>"><%=AgeGroups(1, i)%></option>
                                        <%End If%>
                                    <%Next%>
                                </select>
                            </div>
                            <input type="hidden" name="submit_age_grp" id="submit_age_grp" value="submit_age_grp">
                            <input class="form-control" type="submit" name="get_age_grp" id="get_age_grp" value="Get This">
                            </form>
                            <br>
                            <%If CLng(lAgeGrpID) > 0 Then%>
                                <form class="form" name="edit_age_group" method="Post" 
                                      action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;age_grp_id=<%=lAgeGrpID%>">
                                <div class="form-group">
                                    <label for="age_grp_name">Age Group Name:</label>
                                    <input class="form-control" type="text" name="age_grp_name" id="age_grp_name"
                                    value="<%=sAgeGrp%>">
                                </div>
                                <div class="form-group">
                                    <label for="end_age">End Age</label>
                                    <select class="form-control" name="end_age" id="end_age">
                                        <%For i = 5 To 110%>
                                            <%If CInt(iEndAge) = CInt(i) Then%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label for="delete_age_grp">Delte Age Group</label>
                                    <select class="form-control" name="delete_age_grp" id="delete_age_grp">
                                        <option value="n">n</option>
                                        <option value="y">y</option>
                                    </select>
                                </div>
                                <input type="hidden" name="submit_edit_age_grp" id="submit_edit_age_grp" value="submit_edit_age_grp">
                                <input type="submit" class="form-control" name="save_age_grp" id="save_age_grp" value="Save Changes">
                                </form>
                            <%End If%>
                        </div>
                        <div class="col-sm-3 bg-success">
                            <h5 class="h5">Create New Team</h5>

                            <form class="form" name="new_team" method="post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
                            <div class="form-group">
                                <label for="new_team_name">Team Name:</label>
                                <input class="form-control" type="text" name="new_team_name" id="new_team_name">
                            </div>
                            <div class="form-group">
                                <label>Gender:</label><br>
                                &nbsp;&nbsp;<input type="checkbox" name="combined" id="combined">&nbsp;Combined <br>
                                &nbsp;&nbsp;<input type="checkbox" name="male" id="male">&nbsp;Male <br>
                                &nbsp;&nbsp;<input type="checkbox" name="female" id="female">&nbsp;Female <br>
                            </div>
                            <div class="form-group">
                                <label for="new_age_grp">Age Grp:</label>
                                <select class="form-control" name="new_age_grp" id="new_age_grp">
                                    <%For i = 0 To UBound(AgeGroups, 2) - 1%>
                                        <option value="<%=AgeGroups(0, i)%>"><%=AgeGroups(1, i)%></option>
                                    <%Next%>
                                </select>
                            </div>
                            <div class="form-group">
                                <input type ="hidden" name="create_team" id="create_team" value="create_team">
                                <input class="form-control" type="submit" name="this_team" id="this_team" value="Create Team">
                            </div>
                            </form>
                        </div>
                    </div>
                    <br>
                    <div class="row">
                        <div class="col-sm-12 bg-danger">
                            <h5 class="h5">Existing Teams</h5>
                            <form class="form-inline" name="existing_teams" method="Post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
                            <label for="teams">Select Team</label>
                            <select class="form-control" name="teams" onchange="this.form.get_team.click()">
                                <option value=""></option>
                                <%For i = 0 To UBound(Teams, 2) - 1%>
                                    <%If  CLng(lTeamID) = CLng(Teams(0, i)) Then%>
                                        <option value="<%=Teams(0, i)%>" selected><%=Teams(1, i)%></option>
                                    <%Else%>
                                        <option value="<%=Teams(0, i)%>"><%=Teams(1, i)%></option>
                                    <%End If%>
                                <%Next%>
                            </select>
                            <input type ="hidden" name="select_team" id="select_team" value="select_team">
                            <input class="form-control" type="submit" name="get_team" id="get_team" value="Get This">
                            </form>
                            <br>
                        </div>
                    </div>
                    <%If CLng(lTeamID) > 0 Then%>
                        <br>
                        <div class="row">
                            <div class="col-sm-4">
                                <h5 class="h5">Edit Team</h5>

                                <form class="form" name="edit_this_team" method="Post" action ="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;team_id=<%=lTeamID%>">
                                <div class="form-group">
                                    <label for="team_name">Team Name</label>
                                    <input class="form-control" type="text" name="team_name" id="team_name" value="<%=sTeamName%>">
                                </div>
                                <div class="form-group">
                                    <label for="team_gender">Team Gender</label>
                                    <select class="form-control" name="team_gender" id="team_gender">
                                        <%For i = 0 To UBound(TeamGenders)%>
                                            <%If CStr(TeamGenders(i)) = sTeamGender Then%>
                                                <option value="<%=TeamGenders(i)%>" selected><%=TeamGenders(i)%></option>
                                            <%Else%>
                                                <option value="<%=TeamGenders(i)%>"><%=TeamGenders(i)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label for="team_age_group">Team Age Group</label>
                                    <select class="form-control" name="team_age_group" id="team_age_group">
                                        <%For i = 0 To UBound(AgeGroups, 2) - 1%>
                                            <%If CStr(AgeGroups(1, i)) = sTeamAgeGroup Then%>
                                                <option value="<%=AgeGroups(0, i)%>" selected><%=AgeGroups(1, i)%></option>
                                            <%Else%>
                                                <option value="<%=AgeGroups(0, i)%>"><%=AgeGroups(1, i)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label for="delete_team" style="color:red;">Delete Team (No Undo!)</label>
                                    <input type="checkbox" name="delete_team" id="delete_team">
                                </div>
                                <div class="form-group">
                                    <input type ="hidden" name="edit_team" id="edit_team" value="edit_team">
                                    <input class="form-control" type="submit" name="edit_this_team" id="edit_this_team" value="Save Changes">
                                </div>
                                </form>
                            </div>
                            <div class="col-sm-8">
                                <h5 class="h5">Team Members</h5>
                                <form class="form-inline" name="team_parts" method="Post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;team_id=<%=lTeamID%>">
                                <label for="team_parts">Select Participant</label>
                                <select class="form-control" name="team_parts">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(TeamMembers, 2) - 1%>
                                        <option value="<%=TeamMembers(0, i)%>"><%=TeamMembers(1, i)%></option>
                                    <%Next%>
                                </select>
                                <input type ="hidden" name="delete_part" id="delete_part" value="delete_part">
                                <input class="form-control" type="submit" name="remove_part" id="remove_part" value="Remove From Team">
                                </form>

                                <br>

                                <h5 class="h5">Available Participants</h5>
                                <form class="form-inline" name="available_parts" method="Post" action="teams.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;team_id=<%=lTeamID%>">
                                <label for="participants">Select Participant</label>
                                <select class="form-control" name="participants">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(AvailParts, 2) - 1%>
                                        <option value="<%=AvailParts(0, i)%>"><%=AvailParts(1, i)%></option>
                                    <%Next%>
                                </select>
                                <input type ="hidden" name="select_part" id="select_part" value="select_part">
                                <input class="form-control" type="submit" name="get_part" id="get_part" value="Add to Team">
                                </form>
                            </div>
                        </div>
                    <%End If%>
			    <%End If%>
            <%End If%>
        </div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>