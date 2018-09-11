<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lMeetID, sMeetName, dMeetDate
Dim lRaceID, sRaceName, iDist, sUnits
Dim i, j, k, m, x
Dim RsltsArr(), RacesArr(), TempArr(4)
Dim lTeamID, sTeamName, sGender
Dim lRosterID
Dim sGradeYear
Dim sOrderResultsBy

If Not Session("role") = "coach" Then Response.Redirect "/default.asp?sign_out=y"

lMeetID = Request.QueryString("meet_id")
lTeamID = Request.QueryString("team_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
   
'get order by
sql = "SELECT OrderBy FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sOrderResultsBy = rs(0).Value
Set rs = Nothing

sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = rs(0).Value
sGender = rs(1).Value
Set rs = Nothing

sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value

If Month(rs(1).Value) <=7 Then
	sGradeYear = Right(CStr(Year(rs(1).Value) - 1), 2)
Else
	sGradeYear = Right(CStr(Year(rs(1).Value)), 2)	
End If

Set rs = Nothing

Select Case sGender
	Case "M"
		sGender = "Male"
	Case "F"
		sGender = "Female"
End Select
    
i = 0
ReDim RacesArr(0)
sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lMeetID & " AND Gender = '" & sGender & "'"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    RacesArr(i) = rs(0).Value
    i = i + 1
    ReDim Preserve RacesArr(i)
    rs.MoveNext
Loop
Set rs = Nothing

Private Sub GetRaceRslts(lRaceID)
	sql = "SELECT RaceName, RaceDist, RaceUnits FROM Races WHERE RacesID = " & lRaceID
	Set rs = conn.Execute(sql)
	sRaceName = rs(0).value
	iDist = rs(1).value
	sUnits = rs(2).Value
	Set rs = Nothing
		
	ReDim RsltsArr(4, 0)
	k = 0 
	sql = "SELECT r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.RaceTime, t.TeamsID FROM IndRslts ir "
	sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Grades g On g.RosterID = r.RosterID "
	sql = sql & "INNER JOIN Teams t ON r.TeamsID = t.TeamsID WHERE ir.RacesID = " & RacesArr(i) 
	sql = sql & " AND ir.Place > 0 ORDER BY ir.Place"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    RsltsArr(0, k) = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")	'name
		RsltsArr(1, k) = k + 1
	    RsltsArr(2, k) = rs(2).Value	'grade
	    RsltsArr(3, k) = rs(3).Value	'time
	    RsltsArr(4, k) = rs(4).Value	'team
		k = k + 1
	    ReDim Preserve RsltsArr(4, k)
	    rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

	If sOrderResultsBy = "Time" Then	    
		're-order results if order by time
		For x = 0 To UBound(RsltsArr, 2) - 2
		    For m = x + 1 To UBound(RsltsArr, 2) - 1
		        If ConvertToSeconds(RsltsArr(3, x)) > ConvertToSeconds(RsltsArr(3, m)) Then
		            'swap places if first time is slower than last
		            For k = 0 To 4
		                TempArr(k) = RsltsArr(k, x)
		                RsltsArr(k, x) = RsltsArr(k, m)
		                RsltsArr(k, m) = TempArr(k)
		            Next
		        End If
		    Next
		Next
		
		'get race place
		For x = 0 To UBound(RsltsArr, 2) - 1
			RsltsArr(1, x) = x + 1
		Next
	End If
End Sub

%>
<!--#include file = "../../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../../includes/per_mile_cc.asp" -->
<!--#include file = "../../../includes/per_km_cc.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Our Results</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<h4 class="h4">Our Results for <%=sMeetName%>&nbsp;on &nbsp;<%=dMeetDate%></h4>
					
			<ul class="nav">
				<li class="nav-item"><a class="nav-link" href="lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lMeetID%>">Back</a></li>
				<li class="nav-item"><a class="nav-link" href="javascript:window.print()">Print This</a></li>
				<li class="nav-item"><a class="nav-link" href="dwnld_our_rslts.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lMeetID%>" 
					onClick="openThis(this.href,1024,768);return false;">Download This</a></li>
			</ul>
					
			<%For i = 0 to UBound(RacesArr) - 1%>
				<%Call GetRaceRslts(RacesArr(i))%>
				<h5 class="h5">Race: <%=sRaceName%>&nbsp;&nbsp;Distance: <%=iDist%>&nbsp;<%=sUnits%></h5>
				<table class="table table-striped">
					<tr>
						<th>Rnr</th>
						<th>Name</th>
						<th>Pl</th>
						<th>Gr</th>
						<th>Time</th>
						<th>Per Mi</th>
						<th>Per Km</th>
					</tr>
					<%k = 0%>
					<%For j = 0 to UBound(RsltsArr, 2) - 1%>
						<%If CLng(RsltsArr(4, j)) = CLng(lTeamID) Then%>
							<%k = k + 1%>
							<tr>
								<td><%=k%></td>
								<td><%=RsltsArr(0, j)%></td>
								<td><%=RsltsArr(1, j)%></td>
								<td><%=RsltsArr(2, j)%></td>
								<td><%=RsltsArr(3, j)%></td>
								<td>
									<%If RsltsArr(3, j) <> "00:00" Then%>
										<%=PacePerMile(RsltsArr(3, j), iDist, sUnits)%>
									<%End If%>
								</td>
								<td>
									<%If RsltsArr(3, j) <> "00:00" Then%>
										<%=PacePerKm(RsltsArr(3, j), iDist, sUnits)%>
									<%End If%>
								</td>
							</tr>
						<%End If%>
					<%Next%>
				</table>
			<%Next%>
		</div>
	</div>
</div>
<!--#include file = "../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
