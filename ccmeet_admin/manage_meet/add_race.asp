<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisMeet, lRaceID
Dim sRaceName, sRaceDesc, sRaceTime, iRaceDist, sRaceUnits, sMeetName, sSport
Dim sGender, sScoreMethod, iNumAllow, iNumScore, sComments, sStartType, sRaceType
Dim sTmAwds, sIndAwds
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetName, Sport FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
sSport = rs(1).Value
Set rs = Nothing

If Request.Form.Item("submit_race") = "submit_race" Then
    sRaceName = Replace(Request.Form.Item("race_name"), "'", "''")
    sRaceDesc = Replace(Request.Form.Item("race_desc"), "'", "''")
    sRaceTime = Request.Form.Item("race_time")
	iRaceDist = Request.Form.Item("race_dist")
    sRaceUnits = Request.Form.Item("race_units")
    sGender = Request.Form.Item("gender")
    sScoreMethod = Request.Form.Item("score_meth")
	iNumAllow = Request.Form.Item("num_allow")
 	iNumScore = Request.Form.Item("num_score")
	sComments = Replace(Request.Form.Item("comments"), "'", "''")
    sTmAwds = Replace(Request.Form.Item("tm_awds"), "'", "''")
    sIndAwds = Replace(Request.Form.Item("ind_awds"), "'", "''")
    sStartType = Request.Form.Item("start_type")
    sRaceType = Request.Form.Item("indiv_relay")

    'add the race
    sql = "INSERT INTO Races(MeetsID, RaceName, RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, ScoreMethod, "
    sql = sql & "NumAllow, NumScore, Comments, TmAwds, IndAwds, StartType, IndivRelay) VALUES (" & lThisMeet & ", '" & sRaceName 
    sql = sql & "', '" & sRaceDesc & "', '" & sRaceTime & "', " & iRaceDist & ", '" & sRaceUnits & "', '" & sGender & "', '" 
    sql = sql & sScoreMethod & "', " & iNumAllow & ", " & iNumScore & ", '" & sComments & "', '" & sTmAwds & "', '" 
    sql = sql & sIndAwds & "', '" & sStartType & "', '" & sRaceType & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
    
    'get race id
    sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lThisMeet & " AND RaceName = '" & sRaceName 
    sql = sql & "' AND RaceDesc = '" & sRaceDesc & "' AND RaceTime = '" & sRaceTime & "' AND RaceDist = " 
    sql = sql & iRaceDist & " AND RaceUnits = '" & sRaceUnits & "' AND Gender = '" & sGender & "' AND ScoreMethod = '" 
    sql = sql & sScoreMethod & "' AND NumAllow = " & iNumAllow & " AND NumScore = " & iNumScore 
    sql = sql & " AND Comments = '" & sComments & "' AND TmAwds = '" & sTmAwds
    sql = sql & "' AND IndAwds = '" & sIndAwds & "' AND StartType = '" & sStartType & "' ORDER BY RacesID DESC"
    Set rs = conn.Execute(sql)
    lRaceID = rs(0).Value
    Set rs = Nothing
    
    'insert race delay
    sql = "INSERT INTO RaceDelay (RacesID) VALUES (" & lRaceID & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    If sSport = "Nordic Ski" Then
        'insert race into run order
        sql = "INSERT INTO RunOrder (RacesID) VALUES (" & lRaceID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        If sStartType = "Pursuit" Then
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
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE  Add CC Race</title>

<script>
function chkFields(){
	if (document.add_race.race_name.value==''){
		alert('You must supply a race name!');
		return false;
	}
	else
		if (document.add_race.race_desc.value==''){
			alert('You must supply a race description!');
			return false;
		}
	else
		if (document.add_race.race_time.value==''){
			alert('You must supply a race time!');
			return false;
		}
	else
		if (document.add_race.race_dist.value==''){
			alert('You must supply a race distance!');
			return false;
		}
	else
		if (document.add_race.race_units.value==''){
			alert('You must supply units for this race!');
			return false;
		}
	else
		if (document.add_race.score_meth.value==''){
			alert('You must supply a scoring method for this race!');
			return false;
		}
	else
		if (document.add_race.num_allow.value==''){
			alert('You must submit the number of runners allowed per team (select 0 for unlimited)!');
			return false;
		}
	else
		if (document.add_race.num_score.value==''){
			alert('You must submit the number of runners that score for each team (select 0 if the race is not scored)!');
			return false;
		}
	else
		return true;
}
</script>
</head>

<body onload="document.add_race.race_name.focus();">
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>
			
			<h4 class="h4">Add Race for <%=sMeetName%></h4>
			
			<div style="text-align:right;background-color:#ececd8;font-size:0.85em;margin:0 0 10px 0;">
				<a href="races.asp?meet_id=<%=lThisMeet%>">Back</a>
			</div>
	
			<form name="add_race" method="post" action="add_race.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkFields()">
			<table>
				<tr>
					<th>Race Name:</th>
					<td><input type="text" name="race_name" id="race_name" size="5" maxlength="5"></td>
					<th>Description:</th>
					<td><input type="text" name="race_desc" id="race_desc" size="20" maxlength="50"></td>
				</tr>
				<tr>
					<th>Race Time:</th>
					<td><input type="text" name="race_time" id="race_time" size="8" maxlength="12"></td>
					<th>Distance:</th>
					<td>
						<input type="text" name="race_dist" id="race_dist" size="5" maxlength="12">
						<select name="race_units" id="race_units">
							<option value="">&nbsp;</option>
							<option value="miles">miles</option>
							<option value="kms">km</option>
							<option value="yds">yds</option>
							<option value="meters">meters</option>
						</select>
					</td>
				</tr>
				<tr>
					<th>Gender:</th>
					<td>
						<select name="gender" id="gender">
							<option value="">&nbsp;</option>
							<option value="Male">Male</option>
							<option value="Female">Female</option>
							<option value="Open">Open</option>
						</select>
					</td>
					<th>Scoring Method:</th>
					<td>
						<select name="score_meth" id="score_meth">
							<option value="">&nbsp;</option>
							<option value="Place">Place</option>
							<option value="Time">Time</option>
							<option value="Points">Points</option>
							<option value="None">None</option>
						</select>
					</td>
				</tr>
				<tr>
					<th>Start Type:</th>
					<td>
						<select name="start_type" id="start_type">
							<option value="Mass">Mass</option>
							<option value="Interval">Interval</option>
							<option value="Wave">Wave</option>
                            <option value="Pursuit">Pursuit</option>
						</select>
					</td>
			        <th>Race Type:</th>
			        <td>
				        <select name="indiv_relay" id="indiv_relay">
							<option value="Indiv">Indiv</option>
							<option value="Relay">Relay</option>
				        </select>
			        </td>
				</tr>
				<tr>
					<th valign="top">Team Awards:</th>
					<td><textarea name="tm_awds" id="tm_awds" rows="2" cols="35" style="font-size:1.1em;"></textarea></td>
					<th valign="top">Ind Awards:</th>
					<td><textarea name="ind_awds" id="ind_awds" rows="2" cols="35" style="font-size:1.1em;"></textarea></td>
				</tr>
				<tr>
					<th>Num Allowed/Team:</th>
					<td>
						<select name="num_allow" id="num_allow">
							<option value="0">Unlimited</option>
							<%For i = 1 To 100%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
					</td>
					<th>Num Score/Team:</th>
					<td>
						<select name="num_score" id="num_score">
							<%For i = 0 To 100%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
					</td>
				</tr>
				<tr>
					<th>Comments:</th>
					<td colspan="3"><textarea name="comments" id="comments"  rows="2" cols="35" style="font-size:1.1em;"></textarea></td>
				</tr>
				<tr>
					<td style="text-align:center" colspan="4">
						<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
						<input type="submit" name="submit" id="submit" tabindex="13" value="Submit This Race">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
