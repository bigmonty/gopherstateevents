<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet, lPrevMeet, lRaceID
Dim sSport, sStartType
Dim Meets(), Races(), MeetTeams()
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetDate, Sport FROM Meets WHERE MeetsID = " & lThisMeet 
Set rs = conn.Execute(sql)
dMeetDate = rs(0).Value
sSport = rs(1).Value
Set rs = Nothing

If Request.Form.Item("submit_this") = "submit_this" Then
	lPrevMeet = Request.Form.Item("meets")
	
	If Not CStr(lPrevMeet) & "" = "" Then
		'add races
		i = 0
		ReDim Races(28, 0)
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Races WHERE MeetsID = " & lPrevMeet
		rs.Open sql, conn, 1, 2
		Do While Not rs.EOF
			For j = 0 To 28
				Races(j, i) = rs(j).Value
			Next
			i = i + 1
			ReDim Preserve Races(28, i)
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing

		For i = 0 To UBound(Races, 2) - 1
			sql = "INSERT INTO Races (MeetsID, RaceName, RaceDesc, RaceTime, RaceDist, RaceUnits, Technique, Gender, ScoreMethod, NumAllow, NumScore, "
			sql = sql & "Comments, TmAwds, IndAwds, RemoveInc, StartType, RaceBreak, IndivRelay, TimeAcc, TieBreakersID, Advancement, TeamScores, "
            sql = sql & "ResultsNotes, NumSplits, MinTime, RaceClass, ViewOrder) VALUES (" 
			sql = sql & lThisMeet & ", '" & Races(3, i) & "', '" & Races(4, i) & "', '" & Races(5, i) & "','" & Races(6, i) & "', '" 
			sql = sql & Races(7, i) & "','" & Races(8, i) & "','" & Races(9, i) & "','" & Races(10, i) & "','" & Races(11, i) & "', '" & Races(12, i) 
			sql = sql & "', '" & Races(13, i) & "', '" & Races(14, i) & "', '" & Races(15, i) & "', '" & Races(16, i) & "', '" & Races(17, i) & "', '" 
			sql = sql & Races(18, i) & "', '" & Races(19, i) & "', '" & Races(20, i) & "', '" & Races(21, i) & "', '" & Races(22, i) & "', '" 
            sql = sql & Races(23, i) & "', '" & Races(24, i) & "', '" & Races(25, i) & "', '" & Races(26, i) & "', '" & Races(27, i) & "', '" & Races(28, i) & "')" 
    		Set rs = conn.Execute(sql)
			Set rs = Nothing

		    'get race id
			Set rs = Server.CreateObject("ADODB.REcordset")
		    sql = "SELECT RacesID, RaceSrvrID, StartType FROM Races WHERE MeetsID = " & lThisMeet & " AND RaceName = '" & Races(3, i) & "' ORDER BY RacesID DESC"
		    rs.Open sql, conn, 1, 2
		    lRaceID = rs(0).Value
			rs(1).Value = rs(0).Value
			sStartType = rs(2).Value
			rs.Update
			rs.Close
		    Set rs = Nothing
		    
		    'insert race delay
		    sql = "INSERT INTO RaceDelay (RacesID, RaceSrvrID) VALUES (" & lRaceID & ", " & lRaceID & ")"
		    Set rs = conn.Execute(sql)
		    Set rs = Nothing

			If sSport = "Nordic Ski" Then
				'insert race into run order
				sql = "INSERT INTO RunOrder (RacesID, RaceSrvrID) VALUES (" & lRaceID & ", " & lRaceID & ")"
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
						sql = "INSERT INTO Pursuit (RacesID, RaceSrvrID) VALUES (" & lRaceID & ", " & lRaceID & ")"
						Set rs = conn.Execute(sql)
						Set rs = Nothing
					End If
				End If
			End If
		Next
		
		'add teams
		i = 0
		ReDim MeetTeams(0)
		sql = "SELECT TeamsID FROM MeetTeams WHERE MeetsID = " & lPrevMeet
		Set rs = conn.Execute(sql)
		Do While Not rs.EOF
			MeetTeams(i) = rs(0).Value
			i = i + 1
			ReDim Preserve MeetTeams(i)
			rs.MoveNext
		Loop
		Set rs = Nothing
		
		For i = 0 To UBound(MeetTeams) - 1
			sql = "INSERT INTO MeetTeams (TeamsID, MeetsID) VALUES (" & MeetTeams(i) & ", " & lThisMeet & ")"
			Set rs = conn.Execute(sql)
			Set rs = Nothing
		Next
		
		Response.Redirect "manage_meet.asp?meet_id=" & lThisMeet
	End If
End If

i = 0
ReDim Meets(2, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetsID <> " & lThisMeet & " AND MeetDate <= '" & dMeetDate & "' AND Sport = '" & sSport
sql = sql & "' ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Meets(0, i) = rs(0).Value
	Meets(1, i) = rs(1).Value
	Meets(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve Meets(2, i)
	rs.MoveNext
Loop
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE  Admin Clone CC Meet</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>
			
			<h4 class="h4">Clone Previous Meet</h4>
				
			<p>The purpose of this page is to import the races, teams, and other details to a meet from a previous meet, usually the previous year's
			edition, to save time and effort preparing the event.  THERE IS NO UNDO FOR THIS ACTION!</p>
			
			<form role="form" class="form-inline" name="import_data" method="post" action="clone_prev.asp?meet_id=<%=lThisMeet%>">
			<label for="meets">Select Meet To Clone:</label>
			<select class="form-control" name="meets" id="meets">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Meets, 2) - 1%>
					<option value="<%=Meets(0, i)%>"><%=Meets(1, i)%> (<%=Meets(2, i)%>)</option>
				<%Next%>
			</select>
			<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
			<input class="form-control" type="submit" name="submit1" id="submit1" value="Import Data">
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
