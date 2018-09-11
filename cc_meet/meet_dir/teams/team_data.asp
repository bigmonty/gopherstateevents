<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2, rs3, sql3
Dim i
Dim lThisMeet
Dim sMeetName, sCoachName, sCoachEmail, sTeamClass
Dim iNumEntrants, iMTotal, iFTotal, iTotalEntrants
Dim MTeams(), FTeams(), MeetArr()
Dim dMeetDate

Dim sMapLink, sMeetInfoSheet, sCourseMap
Dim dWhenShutdown

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_meet") = "submit_meet" Then 
    lThisMeet = Request.Form.Item("meets")
End If

i = 0
ReDim MeetArr(1, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDirID = " & Session("my_id") & " ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetArr(0, i) = rs(0).Value
	MeetArr(1, i) = rs(1).Value & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve MeetArr(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If UBound(MeetArr, 2) = 1 Then lThisMeet = MeetArr(0, 0)

If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

If Not CLng(lThisMeet) = 0 Then
    sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
    Set rs = conn.Execute(sql)
    sMeetName = Replace(rs(0).Value, "''", "'")
    dMeetDate = rs(1).Value
    Set rs = Nothing

	'get maplink
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMapLink = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT InfoSheet FROM MeetInfo WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMeetInfoSheet = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get course map
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Map FROM CourseMap WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sCourseMap = rs(0).Value
	rs.Close
	Set rs = Nothing

    'get participating teams info	
    ReDim MTeams(5, 0)
    i = 0
    sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'M' ORDER BY t.TeamName"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        Call GetTeamData(rs(0).Value)

	    MTeams(0, i) = rs(0).Value
	    MTeams(1, i) = Replace(rs(1).Value, "''", "'")
        MTeams(2, i) = sCoachName
        MTeams(3, i) = sCoachEmail
        MTeams(4, i) = sTeamClass
        MTeams(5, i) = iNumEntrants

        iMTotal = CInt(iMTotal) + iNumEntrants
        iTotalEntrants = CInt(iTotalEntrants) + iNumEntrants

	    i = i + 1
	    ReDim Preserve MTeams(5, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing

    ReDim FTeams(5, 0)
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'F' ORDER BY t.TeamName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Call GetTeamData(rs(0).Value)

	    FTeams(0, i) = rs(0).Value
	    FTeams(1, i) = Replace(rs(1).Value, "''", "'")
        FTeams(2, i) = sCoachName
        FTeams(3, i) = sCoachEmail
        FTeams(4, i) = sTeamClass
        FTeams(5, i) = iNumEntrants

        iFTotal = CInt(iFTotal) + iNumEntrants
        iTotalEntrants = CInt(iTotalEntrants) + iNumEntrants

	    i = i + 1
	    ReDim Preserve FTeams(5, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing
End If

Private Sub GetTeamData(lThisTeam)
    sCoachName = vbNullString
    sCoachEmail = vbNullString
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT c.FirstName, c.LastName, c.Email FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID WHERE t.TeamsID = " & lThisTeam
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then
        sCoachName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
        sCoachEmail = rs2(2).Value
    End If
    rs2.Close
    Set rs2 = Nothing

    sTeamClass = "n/a"
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet & " AND TeamsID = " & lThisTeam
    rs2.Open sql2, conn, 1, 2
    If Not rs2(0).Value & "" = "" Then sTeamClass = ClassName(rs2(0).Value)
    rs2.Close
    Set rs2 = Nothing

    iNumEntrants = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT ir.RosterID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.MeetsID = " & lThisMeet 
    sql2 = sql2 & " AND r.TeamsID = " & lThisTeam
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then iNumEntrants = rs2.RecordCount
    rs2.Close
    Set rs2 = Nothing
End Sub

Private Function ClassName(lThisClass)
    Set rs3 = Server.CreateObject("ADODB.Recordset")
    sql3 = "SELECT ClassName FROM MeetClasses WHERE MeetClassesID = " & lThisClass 
    rs3.Open sql3, conn, 1, 2
    ClassName = rs3(0).Value
    rs3.Close
    Set rs3 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>Cross-Country/Nordic Ski Team Data</title>
<!--#include file = "../../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	<!--#include file = "../../../includes/meet_dir_menu.asp" -->

	<h4 class="h4">CC/Nordic Meet Director: Team Data</h4>

	<form class="form-inline" name="get_meets" method="post" action="team_data.asp?meet_id=<%=lThisMeet%>">
	<div>
		<label for="meets">Select Meet:</label>
		<select class="form-control" name="meets" id="meets" onchange="this.form.submit1.click();">
			<option value="">&nbsp;</option>
			<%For i = 0 to UBound(MeetArr, 2) - 1%>
				<%If CLng(lThisMeet) = CLng(MeetArr(0, i)) Then%>
					<option value="<%=MeetArr(0, i)%>" selected><%=MeetArr(1, i)%></option>
				<%Else%>
					<option value="<%=MeetArr(0, i)%>"><%=MeetArr(1, i)%></option>
				<%End If%>
			<%Next%>
		</select>
		<input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
		<input type="submit" class="form-control" name="submit1" id="submit1" value="Get This">
	</div>
	</form>
			
    <%If Not CLng(lThisMeet) = 0 Then%>
		<!--#include file = "../meet_dir_nav.asp" -->
			
        <h5 class="h5">Total Entrants:&nbsp;<%=iTotalEntrants%><h5>

		<div class="col-xs-6">
			<h4 class="h4">Female Teams</h4>
			<table class="table table-striped">
                <tr><th>No.</th><th>Team</th><th>Coach</th><th>Class</th><th>Entrants</th></tr>
                <%For i = 0 To UBound(FTeams, 2) - 1%>
                    <tr>
                        <td><%=i + 1%>)</td>
                        <td><%=FTeams(1, i)%></td>
                        <td><a href="mailto:<%=FTeams(3, i)%>"><%=FTeams(2, i)%></a></td>
                        <td><%=FTeams(4, i)%></td>
                        <td><%=FTeams(5, i)%></td>
                    </tr>
                <%Next%>
                <tr><th style="text-align: right;" colspan="4">Total:</th><th style="text-align: right;"><%=iFTotal%></th></tr>
            </table>
		</div>
        <div class="col-xs-6">
			<h4 class="h4">Male Teams</h4>
			<table class="table table-striped">
                <tr><th>No.</th><th>Team</th><th>Coach</th><th>Class</th><th>Entrants</th></tr>
                <%For i = 0 To UBound(MTeams, 2) - 1%>
                    <tr>
                        <td><%=i + 1%>)</td>
                        <td><%=MTeams(1, i)%></td>
                        <td><a href="mailto:<%=MTeams(3, i)%>"><%=MTeams(2, i)%></a></td>
                        <td><%=MTeams(4, i)%></td>
                        <td><%=MTeams(5, i)%></td>
                    </tr>
                <%Next%>
                <tr><th style="text-align: right;" colspan="4">Total:</th><th style="text-align: right;"><%=iMTotal%></th></tr>
            </table>
		</div>
    <%End If%>
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
