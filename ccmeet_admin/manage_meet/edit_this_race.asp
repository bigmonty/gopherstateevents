<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lRaceID
Dim RaceDetails(22)
Dim sDelete
Dim bFound

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	sDelete = Request.Form.Item("delete")
	
	If sDelete = "y" Then
		sql = "DELETE FROM Races WHERE RacesID = " & lRaceID
		Set rs = conn.Execute(sql)
		Set rs = Nothing

    	Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
		Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
	Else
		RaceDetails(0) = Replace(Request.Form.Item("race_desc"), "'", "''")
		RaceDetails(1) = Request.Form.Item("race_time")
		RaceDetails(2) = Request.Form.Item("race_dist")
		RaceDetails(3) = Request.Form.Item("race_units")
		RaceDetails(4) = Request.Form.Item("gender")
		RaceDetails(5) = Request.Form.Item("score_method")
		RaceDetails(6) = Request.Form.Item("num_allow")
		RaceDetails(7) = Request.Form.Item("num_score")
		RaceDetails(8) = Replace(Request.Form.Item("comments"), "'", "''")
		RaceDetails(9) = Replace(Request.Form.Item("tm_awds"), "'", "''")
		RaceDetails(10) = Replace(Request.Form.Item("ind_awds"), "'", "''")
		RaceDetails(11) = Request.Form.Item("start_type")
	    RaceDetails(12) = Request.Form.Item("race_name")
        RaceDetails(13) = Request.Form.Item("indiv_relay")
        RaceDetails(14) = Request.Form.Item("team_scores")
        If Not Request.Form.Item("results_notes") & "" = "" Then RaceDetails(15) = Replace(Request.Form.Item("results_notes"), "'", "''")
        RaceDetails(16) = Request.Form.Item("num_splits")
        RaceDetails(17) = Request.Form.Item("view_order")
        RaceDetails(18) = Request.Form.Item("technique")
        RaceDetails(19) = Request.Form.Item("num_laps")
        RaceDetails(20) = Request.Form.Item("show_results")
        RaceDetails(21) = Request.Form.Item("stage_race")
        RaceDetails(22) = Request.Form.Item("order_by")

		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, ScoreMethod, NumAllow, NumScore, Comments, TmAwds, IndAwds, StartType, "
        sql = sql & "RaceName, IndivRelay, TeamScores, ResultsNotes, NumSplits, ViewOrder, Technique, NumLaps, ShowResults, StageRace, OrderBy "
        sql = sql & "FROM Races WHERE RacesID = " & lRaceID
		rs.Open sql, conn, 1, 2
		For i = 0 to 22
			rs(i).Value = RaceDetails(i)
		Next
		rs.Update
		rs.Close
		Set rs = Nothing

        If RaceDetails(11) = "Pursuit" Then
             bFound = False
             Set rs = Server.CreateObject("ADODB.Recordset")
             sql = "SELECT PursuitID FROM Pursuit WHERE RacesID = " & lRaceID
             rs.Open sql, conn, 1, 2
             If rs.RecordCount > 0 Then bFound = True
             rs.Close
             Set rs = Nothing
             
             If bFound = False Then
                 sql = "INSERT INTO Pursuit (RacesID) VALUES (" & lRaceID & ")"
                 Set rs = conn.Execute(sql)
                 Set rs = Nothing
             End If
        Else
            sql = "DELETE FROM Pursuit WHERE RacesID = " & lRaceID
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
	End If

    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
End If

sql = "SELECT RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, ScoreMethod, NumAllow, NumScore, Comments, TmAwds, IndAwds, StartType, RaceName, "
sql = sql & "IndivRelay, TeamScores, ResultsNotes, NumSplits, ViewOrder, Technique, NumLaps, ShowResults, StageRace, OrderBy "
sql = sql & "FROM Races WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
For i = 0 to 22
	If Not rs(i).Value & "" = "" Then RaceDetails(i) = rs(i).Value
Next
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Edit CCMeet Race</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<h4 class="h4">Edit CCMeet Race</h4>
	
	<form role="form" class="form" name="edit_race" method="post" action="edit_this_race.asp?race_id=<%=lRaceID%>">
    <table class="table table-striped">
		<tr>
			<th>Name:</th>
			<td>
				<input class="form-control" type="text" name="race_name" id="race_name" value="<%=RaceDetails(12)%>">
			</td>
			<th>Description:</th>
			<td>
				<input class="form-control" type="text" name="race_desc" id="race_desc" value="<%=RaceDetails(0)%>">
			</td>
		</tr>
		<tr>
			<th>Time:</th>
			<td><input class="form-control" type="text" name="race_time" id="race_time" value="<%=RaceDetails(1)%>"></td>
			<th>Distance:</th>
			<td style="white-space:nowrap;">
				<input class="form-control" type="text" name="race_dist" id="race_dist" value="<%=RaceDetails(2)%>"> 
				<select class="form-control" name="race_units" id="race_units">
					<%Select Case RaceDetails(3)%>
						<%Case "miles"%>
							<option value="miles" selected>miles</option>
							<option value="kms">kms</option>
							<option value="yards">yds</option>
							<option value="meters">meters</option>
						<%Case "kms"%>
							<option value="miles">miles</option>
							<option value="kms" selected>kms</option>
							<option value="yards">yds</option>
							<option value="meters">meters</option>
						<%Case "yds"%>
							<option value="miles">miles</option>
							<option value="kms">kms</option>
							<option value="yards" selected>yds</option>
							<option value="meters">meters</option>
						<%Case "meters"%>
							<option value="miles">miles</option>
							<option value="kms">kms</option>
							<option value="yards">yds</option>
							<option value="meters" selected>meters</option>
					<%End Select%>
				</select>
			</td>
		</tr>
		<tr>
			<th>Gender:</th>
			<td>
				<select class="form-control" name="gender" id="gender">
					<%Select Case RaceDetails(4)%>
						<%Case "Male"%>
							<option value="Male" selected>Male</option>
							<option value="Female">Female</option>
							<option value="Open">Open</option>
						<%Case "Female"%>
							<option value="Male">Male</option>
							<option value="Female" selected>Female</option>
							<option value="Open">Open</option>
						<%Case "Open"%>
							<option value="Male">Male</option>
							<option value="Female">Female</option>
							<option value="Open" selected>Open</option>
					<%End Select%>
				</select>
			</td>
			<th>Scoring Method:</th>
			<td>
				<select class="form-control" name="score_method" id="score_method">
					<%Select Case RaceDetails(5)%>
						<%Case "Place"%>
							<option value="Place" selected>Place</option>
							<option value="Time">Time</option>
							<option value="Points">Points</option>
							<option value="Points">Pursuit</option>
							<option value="None">None</option>
						<%Case "Time"%>
							<option value="Place">Place</option>
							<option value="Time" selected>Time</option>
							<option value="Points">Points</option>
							<option value="Points">Pursuit</option>
							<option value="None">None</option>
						<%Case "Points"%>
							<option value="Place">Place</option>
							<option value="Time">Time</option>
							<option value="Points" selected>Points</option>
							<option value="Pursuit">Pursuit</option>
							<option value="None">None</option>
						<%Case "Pursuit"%>
							<option value="Place">Place</option>
							<option value="Time">Time</option>
							<option value="Points">Points</option>
							<option value="Pursuit" selected>Pursuit</option>
							<option value="None">None</option>
						<%Case "None"%>
							<option value="Place">Place</option>
							<option value="Time">Time</option>
							<option value="Points">Points</option>
							<option value="Pursuit">Pursuit</option>
							<option value="None" selected>None</option>
					<%End Select%>
				</select>
			</td>
		</tr>
		<tr>
			<th>Per Team:</th>
			<td>
				<select class="form-control" name="num_allow" id="num_allow">
					<%For i = 0 to 100%>
						<%If CInt(RaceDetails(6)) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
			</td>
			<th>Num Score:</th>
			<td>
				<select class="form-control" name="num_score" id="num_score">
					<%For i = 0 to 100%>
						<%If CInt(RaceDetails(7)) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
			</td>
		</tr>
		<tr>
			<th>Start Type:</th>
			<td>
				<select class="form-control" name="start_type" id="start_type">
					<%Select Case RaceDetails(11)%>
						<%Case "Mass"%>
							<option value="Mass" selected>Mass</option>
							<option value="Interval">Interval</option>
							<option value="Wave">Wave</option>
							<option value="Pursuit">Pursuit</option>
						<%Case "Interval"%>
							<option value="Mass">Mass</option>
							<option value="Interval" selected>Interval</option>
							<option value="Wave">Wave</option>
							<option value="Pursuit">Pursuit</option>
						<%Case "Wave"%>
							<option value="Mass">Mass</option>
							<option value="Interval">Interval</option>
							<option value="Wave" selected>Wave</option>
							<option value="Pursuit">Pursuit</option>
						<%Case "Pursuit"%>
							<option value="Mass">Mass</option>
							<option value="Interval">Interval</option>
							<option value="Wave">Wave</option>
							<option value="Pursuit" selected>Pursuit</option>
						<%Case Else%>
                            <option value="">&nbsp;</option>
							<option value="Mass">Mass</option>
							<option value="Interval">Interval</option>
							<option value="Wave">Wave</option>
							<option value="Pursuit">Pursuit</option>
					<%End Select%>
				</select>
			</td>
			<th>Team Awards:</th>
			<td><input class="form-control" type="text" name="tm_awds" id="tm_awds" value="<%=RaceDetails(9)%>"></td>
		</tr>
		<tr>
			<th>Ind Awards:</th>
			<td><input class="form-control" type="text" name="ind_awds" id="ind_awds" value="<%=RaceDetails(10)%>"></td>
			<th>Race Type:</th>
			<td>
				<select class="form-control" name="indiv_relay" id="indiv_relay">
					<%Select Case RaceDetails(13)%>
						<%Case "Relay"%>
							<option value="Indiv">Indiv</option>
							<option value="Relay" selected>Relay</option>
						<%Case Else%>
							<option value="Indiv" selected>Indiv</option>
							<option value="Relay">Relay</option>
					<%End Select%>
				</select>
			</td>
		</tr>
		<tr>
			<th>Team Scores:</th>
			<td>
				<select class="form-control" name="team_scores" id="team_scores">
					<%If RaceDetails(14) = "y" Then%>
                        <option value="y" selected>Yes</option>
						<option value="n">No</option>
					<%Else%>
                        <option value="y">Yes</option>
						<option value="n" selected>No</option>
					<%End If%>
				</select>
			</td>
 			<th>Num Splits:</th>
			<td>
				<select class="form-control" name="num_splits" id="num_splits">
                    <%For i = 0 To 4%>
						<%If CInt(RaceDetails(16)) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
                    <%Next%>
				</select>
			</td>
		</tr>
		<tr>
 			<th>View Order:</th>
			<td>
				<select class="form-control" name="view_order" id="view_order">
                    <%For i = 0 To 25%>
						<%If CInt(RaceDetails(17)) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
                    <%Next%>
				</select>
			</td>
			<th>Technique:</th>
			<td>
				<select class="form-control" name="technique" id="technique">
                    <option value="">&nbsp;</option>
					<%If RaceDetails(18) = "Classical" Then%>
                        <option value="Classical" selected>Classical</option>
						<option value="Freestyle">Freestyle</option>
					<%ElseIf RaceDetails(18) = "Freestyle" Then%>
                        <option value="Classical">Classical</option>
						<option value="Freestyle" selected>Freestyle</option>
                    <%Else%>
                        <option value="Classical">Classical</option>
						<option value="Freestyle">Freestyle</option>
					<%End If%>
				</select>
                (nordic ski only)
			</td>
        </tr>
        <tr>
 			<th>Num Laps:</th>
			<td>
				<select class="form-control" name="num_laps" id="num_laps">
                    <%For i = 1 To 6%>
						<%If CInt(RaceDetails(19)) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
                    <%Next%>
				</select>
			</td>
			<th>Show Results:</th>
			<td>
				<select class="form-control" name="show_results" id="show_results">
					<%If RaceDetails(20) = "y" Then%>
                        <option value="y" selected>Yes</option>
						<option value="n">No</option>
					<%Else%>
                        <option value="y">Yes</option>
						<option value="n" selected>No</option>
					<%End If%>
				</select>
			</td>
		</tr>
        <tr>
			<th>Stage Race?</th>
			<td>
				<select class="form-control" name="stage_race" id="stage_race">
					<%If RaceDetails(21) = "y" Then%>
                        <option value="y" selected>Yes</option>
						<option value="n">No</option>
					<%Else%>
                        <option value="y">Yes</option>
						<option value="n" selected>No</option>
					<%End If%>
				</select>
			</td>
			<th>Order Rslts By:</th>
			<td>
				<select class="form-control" name="order_by" id="order_by">
					<%Select Case RaceDetails(22)%>
						<%Case "time"%>
							<option value="time" selected>Time</option>
							<option value="place">Place</option>
						<%Case Else%>
							<option value="time">Time</option>
							<option value="place" selected>Place</option>
					<%End Select%>
				</select>
			</td>

		</tr>
		<tr>
			<td style="font-weight:bold;text-align:right" valign="top">
				Comments:
			</td>
			<td colspan="3">
				<textarea class="form-control" name="comments" id="comments" rows="3"><%If Not RaceDetails(8) & "" = "" Then Response.Write(Replace(RaceDetails(8), "''", "'"))%></textarea>
			</td>
		</tr>
		<tr>
			<td style="font-weight:bold;text-align:right" valign="top">
				Results Notes:
			</td>
			<td colspan="3">
				<textarea class="form-control" name="results_notes" id="results_notes" rows="3"><%If Not RaceDetails(15) & "" = "" Then Response.Write(Replace(RaceDetails(15), "''", "'"))%></textarea>
			</td>
		</tr>
		<tr>
			<td class="bg-danger text-danger" colspan="4">
				Delete This Race? (NOTE: There is no undo for this action!)
				<select class="form-control" name="delete" id="delete">
					<option value="n">No</option>
					<option value="y">Yes</option>
				</select>
			</td>
		</tr>
		<tr>
			<td style="text-align:center" colspan="4">
				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input class="form-control" type="submit" name="submit" id="submit" tabindex="10" value="Save Changes">
			</td>
		</tr>
	</table>
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
