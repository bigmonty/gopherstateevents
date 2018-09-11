<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lTeamID, lThisMeet
Dim sTeamName, sGender, sCoachName, sCoachPhone, sCoachEmail, sUserName, sPassword, sMeetClass
Dim MeetClasses()

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lTeamID = Request.QueryString("team_id")
lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

ReDim MeetClasses(1, 0)

If Request.Form.Item("submit_this") = "submit_this" Then
	sTeamName = Replace(Request.Form.Item("team_name"), "'", "''")
	sGender = Request.Form.Item("gender")
	
	If Request.Form.Item("remove") = "y" Then
		sql = "DELETE FROM MeetTeams WHERE TeamsID = " & lTeamID & " AND MeetsID = " & lThisMeet
		Set rs = conn.Execute(sql)
		Set rs = Nothing

        Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
		Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
	Else
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
		rs.Open sql, conn, 1, 2
		rs(0).Value = sTeamName
		rs(1).Value = sGender
		rs.Update
		rs.Close
		Set rs = Nothing

        Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
	End If
End If

sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = Replace(rs(0).Value, "''", "'")
sGender = rs(1).Value
Set rs = Nothing
	
sql = "SELECT c.FirstName, c.LastName, c.Phone, c.Email, c.UserID, c.Password FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID "
sql = sql & "WHERE t.TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sCoachName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
sCoachPhone = rs(2).Value
sCoachEmail = rs(3).Value
sUserName = rs(4).Value
sPassword = rs(5).Value
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>CCMeet Edit Team</title>
<!--#include file = "../../includes/js.asp" -->

<script>
function chkFields(){
	if (document.edit_team.team_name.value==''){
		alert('You must supply a race description!');
		return false;
	}
	else 
		if (document.edit_team.gender.value==''){
			alert('You must supply a race time!');
			return false;
		}
	else 
		return true;
}
</script>

</head>
<body>
<div class="container">
	<h4 class="h4">Edit This Team</h4>

    <div class="col-sm-4 bg-success">
        <br>
	    <ul class="list-group">
		    <li class="list-group-item">Name: <%=sCoachName%></li>
		    <li class="list-group-item">Phone: <%=sCoachPhone%></li>
		    <li class="list-group-item"><a href="mailto:<%=sCoachEmail%>"><%=sCoachEmail%></a></li>
            <li class="list-group-item">User Name:&nbsp;<%=sUserName%></li>
            <li class="list-group-item">Password:&nbsp;<%=sPassword%></li>
	    </ul>
    </div>
    <div class="col-sm-8">
	    <form role="form" class="form-horizontal" name="edit_team" method="post" action="edit_this_team.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>" onsubmit="return chkFields()">
		<div class="form-group">
			<label for="team_name" class="control-label col-xs-4">Team:</label>
			<div class="col-xs-8">
                <input type="text" class="form-control" name="team_name" id="team_name" maxlength="50" value="<%=sTeamName%>">
            </div>
		</div>
		<div class="form-group">
			<label for="gender" class="control-label col-xs-4">M/F:</label>
			<div class="col-xs-8">
				<select class="form-control" name="gender" id="gender">
					<%Select Case sGender%>
						<%Case "M"%>
							<option value="M" selected>Male</option>
							<option value="F">Female</option>
						<%Case "F"%>
							<option value="M">Male</option>
							<option value="F" selected>Female</option>
					<%End Select%>
				</select>
            </div>
		</div>
		<div class="form-group">
			<label for="remove" class="control-label col-xs-4">Remove:</label>
			<div class="col-xs-8">
				<select class="form-control" name="remove" id="remove">
					<option value="n" selected>No</option>
					<option value="y">Yes</option>
				</select>
            </div>
		</div>
        <%If Session("role") = "admin" Then%>
			<div class="form-group">
				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input type="submit" class="form-control" name="submit" id="submit" tabindex="4" value="Save Changes">
			</div>
        <%End If%>
	    </form>
    </div>
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
