<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lTeamID, lMyID
Dim sFirstName, sLastName, sGender, sNewFirst, sNewLast, sNewGender, sErrMsg, sGradeYear, sArchive, sTeamName, sTeamGender
Dim iNewGrade, iGrade, iMyGrade
Dim RosterArr(), DeleteArr()
Dim dShutdown, dMeetDate
Dim bInsertThis

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lTeamID = Request.QueryString("team_id")

sArchive = Request.QueryString("archive")
If sArchive = vbNullString Then sArchive = "n"

'get year for roster grades
If Month(Date) <=7 Then
	sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
Else
	sGradeYear = Right(CStr(Year(Date)), 2)	
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Gender, TeamName FROM Teams WHERE TeamsID = " & lTeamID
rs.Open sql, conn, 1, 2
sTeamGender = rs(0).Value
sTeamName = Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

If Request.Form.Item("edit_roster") = "edit_roster" Then
	i = 0
	ReDim DeleteArr(0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RosterID, FirstName, LastName, Gender, Archive FROM Roster WHERE TeamsID = " & lTeamID & " AND Archive = '" & sArchive & "'"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		If Request.Form.Item("update_" & rs(0).Value) = "y" Then
			sFirstName = Replace(Request.Form.Item("first_name_" & rs(0).value), "'", "''")
			sLastName = Replace(Request.Form.Item("last_name_" & rs(0).value), "'", "''")
			sGender = Request.Form.Item("gender_" & rs(0).value)
            iMyGrade = Request.Form.Item("grade_" & rs(0).value)
            sArchive = Request.Form.Item("archive_" & rs(0).value)

			If sFirstName = vbNullString Then
				rs(1).Value = rs(1).OriginalValue
			Else
				rs(1).Value = sFirstName
			End If
	
			If sLastName = vbNullString Then
				rs(2).Value = rs(2).OriginalValue
			Else
				rs(2).Value = sLastName
			End If
	
			rs(3).Value = sGender
			rs(4).Value = sArchive	
			rs.Update
				
			Call UpdateGrade(rs(0).Value, iMyGrade)
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
ElseIf Request.Form.Item("add_part") = "add_part" Then
	sNewFirst = Replace(Request.Form.Item("first_name"), "'", "''")
	sNewLast = Replace(Request.Form.Item("last_name"), "'", "''")
	sNewGender = Request.Form.Item("gender")
	iNewGrade = Request.Form.Item("grade")
    
    'see if they exist in the db
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FirstName, LastName, Gender FROM Roster WHERE FirstName = '" & sNewFirst & "' AND LastName = '" 
    sql = sql & sNewLast & "' AND Gender = '" & sNewGender & "' AND TeamsID = " & lTeamID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs.Close
        Set rs = Nothing
        
        sErrMsg = "An athlete with this information already exists in the database.  If you do not see them on your roster they may have been archived.  "
		sErrMsg = sErrMsg & "You can view your archived athletes by clicking the 'Show Archives' link above, and then change them to active if you wish.  "
		sErrmsg = sErrMsg & " Otherwise, please make a minor change in this athlete's name (add a middle initial for instance) and then re-enter them."
    Else
        rs.Close
        Set rs = Nothing
        
        'insert team member
        sql = "INSERT INTO Roster (TeamsID, FirstName, LastName, Gender) VALUES (" & lTeamID & ", '" & sNewFirst & "', '" & sNewLast & "', '" 
		sql = sql & sNewGender & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
        
        'get roster id
        sql = "SELECT RosterID FROM Roster WHERE TeamsID = " & lTeamID & " AND FirstName = '" & sNewFirst & "' AND LastName = '"
        sql = sql & sNewLast & "' AND Gender = '" & sNewGender & "' ORDER BY RosterID DESC"
        Set rs = conn.Execute(sql)
        lMyID = rs(0).Value
        Set rs = Nothing
 
		'get year for roster grades
		If Month(Date) <=7 Then
			sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
		Else
			sGradeYear = Right(CStr(Year(Date)), 2)	
		End If
       
        'insert grade
        sql = "INSERT INTO Grades (RosterID, Grade" & sGradeYear & ") VALUES (" & lMyID & ", " & iNewGrade & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
		
		sNewFirst = vbNullString
		sNewLast = vbNullString
		iNewGrade = 0
    End If
End If

'get the time to the next meet this team is participating in
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT m.MeetDate, m.WhenShutdown FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID "
sql = sql & "WHERE mt.TeamsID = " & lTeamID & " AND m.MeetDate >= '" & Date & "' ORDER BY m.MeetDate"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
	dMeetDate = rs(0).Value
	dShutdown = rs(1).Value
End If
rs.Close
Set rs = Nothing

i = 0
ReDim RosterArr(5, 0)
sql = "SELECT RosterID, FirstName, LastName, Gender, Archive FROM Roster WHERE TeamsID = " & lTeamID
sql = sql & " AND Archive = '" & sArchive & "' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RosterArr(0, i) = rs(0).Value
	RosterArr(1, i) = Replace(rs(1).Value, "''", "'")
	RosterArr(2, i) = Replace(rs(2).Value, "''", "'")
	RosterArr(3, i) = GetGrade(rs(0).Value)
	RosterArr(4, i) = rs(3).Value
    RosterArr(5, i) = rs(4).Value
	i = i + 1
	ReDim Preserve RosterArr(5, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Function IncrGrade(lThisPart)
	IncrGrade = 0
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & Right(CStr(Year(Date)), 2) & ", Grade" & Right(CStr(Year(Date)), 2) - 1 & " FROM Grades WHERE RosterID = " & lThisPart
	rs2.Open sql2, conn, 1, 2
	If Not rs2(1).Value & "" = "" Then 
		IncrGrade = CInt(rs2(1).Value) + 1
		rs2(0).Value = CInt(rs2(1).Value) + 1
		rs2.Update
	End If
	rs2.Close
	Set rs2 = Nothing
End Function
	
Private Function UpdateGrade(lMyID, iCurrGrade)
    bInsertThis = False

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then 
        rs2(0).Value = iCurrGrade
	    rs2.Update
    Else
        bInsertThis = True
    End If
	rs2.Close
	Set rs2 = Nothing

    If bInsertThis = True Then
        sql2 = "INSERT INTO Grades (RosterID,  Grade" & Right(CStr(Year(Date)), 2) & ") Values (" & lMyID & ", " & iCurrGrade & ")"
        Set rs2 = conn.Execute(sql2)
        Set rs2 = Nothing
    End If
End Function

Private Function GetGrade(lMyID)
    GetGrade = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetGrade = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function

If CStr(iNewGrade) = vbNullString Then iNewGrade = 0
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Edit Roster</title>
<!--#include file = "../../includes/js.asp" -->

<script>
function chkFields(){
	if (document.add_part.first_name.value==''||
	document.add_part.last_name.value==''||
	document.add_part.gender.value==''||
	document.add_part.grade.value==''){
		alert('All fields are required!');
		return false;
	}
	else
		return true;
}
</script>
</head>
<body style="background: none;background-color: #fff;" onload="document.add_part.first_name.focus();">
<div style="margin: 10px;padding: 10px;font-size: 0.9em;">
	<h3 class="h3">Edit GSE Cross-Country Roster</h3>
    <h4 class="h4"><%=sTeamName%> (<%=sTeamGender%>)</h4>
			
	<div style="text-align:right;font-size:0.9em;">
		<%If sArchive = "y" Then%>
            <a href="edit_roster.asp?team_id=<%=lTeamID%>&amp;archive=n">Active Roster</a>
        <%Else%>
            <a href="edit_roster.asp?team_id=<%=lTeamID%>&amp;archive=y">Archived Roster</a>
        <%End If%>
        &nbsp;|&nbsp;
	    <a href="javascript:pop('/ccmeet_admin/manage_team/roster_upload/batch_upload.asp?team_id=<%=lTeamID%>',800,600)">Upload Roster</a>
	</div>

	<div style="width:400px;float:left;">
		<form name="edit_roster" method="post" action="edit_roster.asp?team_id=<%=lTeamID%>&amp;archive=<%=sArchive%>">
		<%If sArchive = "y" Then%>
            <h4 style="background-color:#ececec;">Archived Roster</h4>
        <%Else%>
            <h4 style="background-color:#ececec;">Active Roster</h4>
        <%End If%>
		<table>
			<tr>
				<td style="text-align:center" colspan="6">
					<input type="hidden" name="edit_roster" id="edit_roster" value="edit_roster">
					<input type="submit" name="submit1" id="submit1" value="Save Changes">
				</td>
			</tr>
			<tr>
				<td style="font-weight:bold;text-align:right;border-bottom:1px solid #003399">No.</td>
				<td style="font-weight:bold;text-align:left;border-bottom:1px solid #003399;white-space:nowrap;">First Name</td>
				<td style="font-weight:bold;text-align:left;border-bottom:1px solid #003399;white-space:nowrap;">Last Name</td>
				<td style="font-weight:bold;text-align:center;border-bottom:1px solid #003399">Grade</td>
				<td style="font-weight:bold;text-align:center;border-bottom:1px solid #003399">Gender</td>
                <td style="font-weight:bold;text-align:center;border-bottom:1px solid #003399">Archive?</td>
			</tr>
			<%For i = 0 to UBound(RosterArr, 2) - 1%>
				<tr>
					<td style="text-align:right">
						<%=i +1%>)
					</td>
					<td style="text-align:left">
						<input type="text" name="first_name_<%=RosterArr(0, i)%>" id="first_name_<%=RosterArr(0, i)%>" 
									size="15" maxlength="15" value="<%=RosterArr(1, i)%>">
						<input type="hidden" name="update_<%=RosterArr(0, i)%>" id="update_<%=RosterArr(0, i)%>" value="y">
					</td>
					<td>
						<input type="text" name="last_name_<%=RosterArr(0, i)%>" id="last_name_<%=RosterArr(0, i)%>" 
									size="25" maxlength="25" value="<%=RosterArr(2, i)%>">
					</td>
					<td>
						<select name="grade_<%=RosterArr(0, i)%>" id="grade_<%=RosterArr(0, i)%>"> 
							<%For j = 0 to 16%>
								<%If CInt(RosterArr(3, i)) = CInt(j) Then%>
									<option value="<%=j%>" selected><%=j%></option>
								<%Else%>
									<option value="<%=j%>"><%=j%></option>
								<%End If%>
							<%Next%>
							</select>
					</td>
					<td>
						<select name="gender_<%=RosterArr(0, i)%>" id="gender_<%=RosterArr(0, i)%>"> 
							<%If RosterArr(4, i) ="M" Then%>
								<option value="M" selected>Male</option>
								<option value="F">Female</option>
							<%Else%>
								<option value="M">Male</option>
								<option value="F" selected>Female</option>
							<%End If%>
						</select>
					</td>
					<td>
						<select name="archive_<%=RosterArr(0, i)%>" id="archive_<%=RosterArr(0, i)%>"> 
							<%If RosterArr(5, i) ="y" Then%>
								<option value="y" selected>Yes</option>
								<option value="n">No</option>
							<%Else%>
								<option value="y">Yes</option>
								<option value="n" selected>No</option>
							<%End If%>
						</select>
					</td>
				</tr>
			<%Next%>
		</table>
		</form>
	</div>
				
	<div style="margin-left:525px;background-color:#ececd8;width:250px;">
		<%If Not sErrMsg = vbNullString Then%>
			<p><%=sErrMsg%></p>
		<%End If%>

		<form name="add_part" method="post" action="edit_roster.asp?team_id=<%=lTeamID%>&amp;archive=n" 
			onsubmit="return chkFields()">
		<h4 class="h4">Add Particpant</h4>
		<table style="margin-top:10px;">
			<tr>
				<td style="text-align:right;white-space:nowrap;">First Name:</td>
				<td style="text-align:left">
					<input type="text" name="first_name" id="first_name" size="15" maxlength="15" value="<%=sNewFirst%>">
				</td>
			</tr>
			<tr>
				<td style="text-align:right;white-space:nowrap;">Last Name:</td>
				<td style="text-align:left">
					<input type="text" name="last_name" id="last_name" size="25" maxlength="25" value="<%=sNewLast%>">
				</td>
			</tr>
			<tr>
				<td style="text-align:right">Grade:</td>
				<td style="text-align:left">
					<select name="grade" id="grade"> 
						<option value="">&nbsp;</option>
						<%For i = 3 To 16%>
							<%If CInt(iNewGrade) = CInt(i) Then%>
								<option value="<%=i%>" selected><%=i%></option>
							<%Else%>
								<option value="<%=i%>"><%=i%></option>
							<%End If%>
						<%Next%>
					</select>
				</td>
			</tr>
			<tr>
				<td style="text-align:right">Gender:</td>
				<td style="text-align:left">
					<select name="gender" id="gender"> 
						<option value="">&nbsp;</option>
						<%If sNewGender = "M" Then%>
							<option value="M" selected>Male</option>
							<option value="F">Female</option>
						<%ElseIf sGender = "F" Then%>
							<option value="M">Male</option>
							<option value="F" selected>Female</option>
						<%Else%>
							<option value="M">Male</option>
							<option value="F">Female</option>
						<%End If%>
				    </select>
				</td>
			</tr>
			<tr>
				<td style="text-align:center" colspan="2">
					<input type="hidden" name="add_part" id="add_part" value="add_part">
					<input type="submit" name="submit2" id="submit2" value="Add Participant">
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
