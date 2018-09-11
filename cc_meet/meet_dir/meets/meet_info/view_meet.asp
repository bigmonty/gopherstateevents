<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lThisMeet
Dim sMeetName, dMeetDate, sMeetSite, sMeetHost, sWebSite, sComments, sEntryFee, sMeetInfoSheet, sCourseMap, sMapLink
Dim iTotalParts
Dim MTeams(), FTeams(), Races(), MeetArr()
Dim dWhenShutdown

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_meet") = "submit_meet" Then lThisMeet = Request.Form.Item("meets")

If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

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

If Not CLng(lThisMeet) = 0 Then
	'get meet info
	sql = "SELECT MeetName, MeetDate, MeetSite, MeetHost, WebSite, Comments, EntryFee, WhenShutdown FROM Meets WHERE MeetsID = " & lThisMeet
	Set rs = conn.Execute(sql)
	sMeetName = Replace(rs(0).Value, "''", "'")
	dMeetDate = rs(1).Value
	If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then sMeetHost = Replace(rs(3).Value, "''", "'")
	sWebSite = rs(4).Value
	If Not rs(5).Value & "" = "" Then sComments = Replace(rs(5).Value, "''", "'")
	sEntryFee = rs(6).Value
	dWhenShutdown = rs(7).Value
	Set rs = Nothing
	
	'get participating teams info	
	ReDim MTeams(0)
	i = 0
	sql = "SELECT t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
	sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'M' ORDER BY t.TeamName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		MTeams(i) = rs(0).Value
		i = i + 1
		ReDim Preserve MTeams(i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	ReDim FTeams(0)
	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
	sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'F' ORDER BY t.TeamName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		FTeams(i) = rs(0).Value
		i = i + 1
		ReDim Preserve FTeams(i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	'get race information
	i = 0
	ReDim Races(6, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RacesID, RaceName, RaceTime, RaceDist, RaceUnits, RaceDesc FROM Races WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		For j = 0 to 4
			Races(j, i) = rs(j).Value
		Next
		Races(5, i) = FieldSize(rs(0).Value)
        Races(6, i) = Replace(rs(5).Value, "''", "'")
		i = i + 1
		ReDim Preserve Races(6,i)
		rs.MoveNext
	Loop
	rs.Close
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

    iTotalParts = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID FROM IndRslts WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iTotalParts = rs.RecordCount
    rs.Close
    Set rs = Nothing
End If

Private Function FieldSize(lThisRaceID)
	FieldSize = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRaceID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then FieldSize = rs2.RecordCount
	rs2.Close
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>GSE Upload Meet Info</title>
<!--#include file = "../../../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
	<!--#include file = "../../../../includes/meet_dir_menu.asp" -->

		<h4 class="h4">CC/Nordic Meet Director: Meet Home Page</h4>

		<form class="form-inline" name="get_meets" method="post" action="view_meet.asp">
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
			<!--#include file = "../../meet_dir_nav.asp" -->
				
			<h4 class="h4">General Meet Information: </h4>

			<div class="bg-info">
				<span style="font-weight:bold;">Meet Host:</span>&nbsp;<%=sMeetHost%>&nbsp;&nbsp;&nbsp;
				<span style="font-weight:bold;">Web Site:</span><%=sWebSite%>&nbsp;&nbsp;&nbsp;
				<span style="font-weight:bold;">Entry Fee:</span>&nbsp;$<%=sEntryFee%>&nbsp;&nbsp;&nbsp;
				<span style="font-weight:bold;">Meet Site:</span>&nbsp;<%=sMeetSite%>&nbsp;&nbsp;&nbsp;
				<span style="font-weight:bold;">Comments:</span><%=sComments%>
			</div>

            <div class="col-xs-6">
				<h5 class="h5">Race Information: </h5>
					
				<table class="table table-striped">
					<tr>
						<th>Race</th>
                        <th>Alias</th>
						<th>Time</th>
						<th>Dist</th>
						<th>Entries</th>
					</tr>
					<%For i = 0 to UBound(Races, 2) - 1%>
                        <tr>
							<td>
								<a href="javascript:pop('/cc_meet/meet_dir/races/race_details.asp?race_id=<%=Races(0, i)%>',400,650)"><%=Races(1, i)%></a>
							</td>
                            <td><%=Races(6, i)%></td>
							<td><%=Races(2, i)%></td>
							<td><%=Races(3, i)%> <%=Races(4, i)%></td>
							<td><%=Races(5, i)%></td>
						</tr>
					<%Next%>
					<tr>
						<th colspan="4">
							Total Entries:&nbsp;<%=iTotalParts%>
						</th>
					</tr>
				</table>
            </div>
            <div class="col-xs-3">
				<h5 class="h5">Female Teams (<%=UBound(FTeams)%>)</h5>
				<ol class="list-group">
					<%For i = 0 to UBound(FTeams) - 1%>
						<li class="list-group-item"><%=FTeams(i)%></li>
					<%Next%>
				</ol>
            </div>
            <div class="col-xs-3">
				<h5 class="h5">Male Teams (<%=UBound(MTeams)%>)</h5>
				<ol class="list-group">
					<%For i = 0 to UBound(MTeams) - 1%>
						<li class="list-group-item"><%=MTeams(i)%></li>
					<%Next%>
				</ol>
            </div>
		<%End If%>
	</div>
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
