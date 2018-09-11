<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lEventID, lRaceID, lParticipantID, lEventGrp
Dim sEventName, sRaceDist
Dim iBegAge
Dim Records(), Races()
Dim BestTime(3), sngThisTime
Dim dEventDate


'Response.Redirect "/misc/taking_break.htm"

lRaceID = Request.QueryString("race_id")

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event name & event group
ReDim Races(2, 0)
If Not CLng(lEventID) = 0 Then
	sql = "SELECT EventName, EventGrp, EventDirID FROM Events WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
	lEventGrp = rs(1).Value
	Set rs = Nothing
	
	i = 0
	sql = "SELECT RaceID, Dist, RaceName FROM RaceData WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		Races(0, i) = rs(0).Value
		Races(1, i) = rs(1).Value
        Races(2, i) = rs(2).Value
		i = i + 1
		ReDim Preserve Races(2, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	If UBound(Races, 2) = 1 Then lRaceID = Races(0, 0)
	
	If Request.Form.Item("submit_race") = "submit_race" Then
		lRaceID = Request.Form.Item("races")
	End If
End If

ReDim Records(4, 0)

Private Sub GetRcds(lThisRace, sMF)
    'get race distance
    sql = "SELECT Dist FROM RaceData WHERE RaceID = " & lRaceID
    Set rs = conn.Execute(sql)
    sRaceDist = rs(0).Value
    Set rs = Nothing

	'set up array with age groups and open
	iBegAge = 0

	i = 1
	Records(0, 0) = "Open"
	ReDim Preserve Records(4, i)
	sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sMF & "' AND RaceID = " & lRaceID & " ORDER BY EndAge"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		If iBegAge = 0 Then
			Records(0, i) = rs(0).Value & " and Under"
		ElseIf rs(0).Value = "110" Then
			Records(0, i) = iBegAge & " and Over"
		Else
			Records(0, i) = iBegAge & " - " & rs(0).Value
		End If
		iBegAge = rs(0).Value + 1
		i = i + 1
		ReDim Preserve Records(4, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
					
	'get open record
	BestTime(0) = "123"
	BestTime(1) = "0"
	BestTime(2) = "10800"
	BestTime(3) = "1/1/1900"
	sql = "SELECT p.ParticipantID, pr.Age, ir.FnlScnds, e.EventDate FROM Participant p INNER JOIN PartRace pr "
	sql = sql & "INNER JOIN IndResults ir INNER JOIN RaceData rd INNER JOIN Events e "
	sql = sql & "ON e.EventID = rd.EventID AND e.EventGrp = " & lEventGrp & " AND rd.Dist = '" & sRaceDist & "' "
	sql = sql & "ON ir.RaceID = rd.RaceID AND ir.FnlScnds > 0 "
	sql = sql & "ON pr.ParticipantID = ir.ParticipantID AND pr.RaceID = ir.RaceID "
	sql = sql & "ON p.ParticipantID = pr.ParticipantID WHERE p.Gender = '" & sMF & "'"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		If Not rs(2).Value & "" = "" Then
			If Not rs(1).Value = "99" Then
				sngThisTime = rs(2).Value
				If CSng(sngThisTime) < CSng(BestTime(2)) Then 
					BestTime(0) = GetPartName(rs(0).Value)
					BestTime(1) = rs(1).Value
					BestTime(2) = sngThisTime
					BestTime(3) = Year(rs(3).Value)
				End If
			End If
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing
					
	'see of a faster record exists in records table
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RcdSetter, Age, Record, RcdYear FROM Records WHERE EventGrp = " & lEventGrp
	sql = sql & " AND RaceDist = '" & sRaceDist & "' AND Gender = '" & sMF & "'"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		Do While Not rs.EOF
			sngThisTime = ConvertToSeconds(rs(2).Value)
			If CSng(sngThisTime) < CSng(BestTime(2)) Then 
				BestTime(0) = rs(0).Value
				BestTime(1) = rs(1).Value
				BestTime(2) = sngThisTime
				BestTime(3) = rs(3).Value
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
					
	Records(1, 0) = BestTime(0)
	Records(2, 0) = BestTime(1)
	Records(3, 0) = ConvertToMinutes(CSng(BestTime(2)))
	Records(4, 0) = BestTime(3)
					
	'get age group records
	For i = 1 to UBound(Records, 2) - 1
		BestTime(0) = "123"
		BestTime(1) = "0"
		BestTime(2) = "10800"
		BestTime(3) = "1/1/1900"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT p.ParticipantID, pr.Age, ir.FnlScnds, e.EventDate FROM Participant p INNER JOIN PartRace pr "
		sql = sql & "INNER JOIN IndResults ir INNER JOIN RaceData rd  INNER JOIN Events e "
		sql = sql & "ON e.EventID = rd.EventID AND e.EventGrp = " & lEventGrp & " AND rd.Dist = '" & sRaceDist & "' "
		sql = sql & "ON ir.RaceID = rd.RaceID AND ir.FnlScnds > 0 "
		sql = sql & "ON pr.ParticipantID = ir.ParticipantID AND pr.RaceID = ir.RaceID "
		sql = sql & "ON p.ParticipantID = pr.ParticipantID "
		sql = sql & "WHERE p.Gender = '" & sMF & "' AND pr.AgeGrp = '" & Records(0, i) & "' AND e.EventGrp = " & lEventGrp 
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then
			Do While Not rs.EOF
				If Not rs(2).Value & "" = "" Then
					If Not rs(1).Value = "99" Then
						sngThisTime = ConvertToSeconds(rs(2).Value)
						If CSng(sngThisTime) < CSng(BestTime(2)) Then 
							BestTime(0) = GetPartName(rs(0).Value)
							BestTime(1) = rs(1).Value
							BestTime(2) = sngThisTime
							BestTime(3) = Year(rs(3).Value)
						End If
					End If
				End If
				rs.MoveNext
			Loop
			Set rs = Nothing
		End If
					
		'see of a faster record exists in records table
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RcdSetter, Age, Record, RcdYear FROM Records WHERE EventGrp = " & lEventGrp
		sql = sql & " AND RaceDist = '" & sRaceDist & "' AND Gender = '" & sMF & "' AND AgeGroup = '" & Records(0, i) & "'"
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then
			Do While Not rs.EOF
				sngThisTime = ConvertToSeconds(rs(2).Value)
				If CSng(sngThisTime) < CSng(BestTime(2)) Then 
					BestTime(0) = rs(0).Value
					BestTime(1) = rs(1).Value
					BestTime(2) = sngThisTime
					BestTime(3) = rs(3).Value
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Set rs = Nothing
						
		If Not BestTime(2) = "10800" Then
			Records(1, i) = BestTime(0)
			Records(2, i) = BestTime(1)
			Records(3, i) = ConvertToMinutes(CSng(BestTime(2)))
			Records(4, i) = BestTime(3)
		End If
	Next
End Sub

Function GetPartName(lParticipantID)
	sql2 = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lParticipantID
	Set rs2 = conn.Execute(sql2)
	GetPartName = Replace(rs2(0).Value, "''", "'") & " " & Replace(rs2(1).Value, "''", "'")
	Set rs2 = Nothing
End Function
%>

<!--#include file = "../includes/convert_to_seconds.asp" -->
<!--#include file = "../includes/convert_to_minutes.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Fitness Events Records</title>
<meta name="description" content="GSE Fitness Events Records for road races, nordic ski, showshoe events, mountain bike, duathlon, and triathlon timing.">
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
    <img class="img-responsive" src="/graphics/html_header.png" alt="GSE Header">

	<h3 class="h3"><%=sEventName%> Records</h3>
		
	<%If UBound(Races, 2) > 1 Then%>
		<form class="form-inline" name="get_races" method="post" action="records.asp?event_id=<%=lEventID%>">
		<div class="form-group">	
			<label for="races">Select Race:</label>
			<select class ="form-control" name="races" id="races" onchange="this.form.get_race.click()">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Races, 2) - 1%>
					<%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
						<option value="<%=Races(0, i)%>" selected><%=Races(2, i)%></option>
					<%Else%>
						<option value="<%=Races(0, i)%>"><%=Races(2, i)%></option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" class ="form-control" name="submit_race" id="submit_race" value="submit_race">
			<input type="submit" class ="form-control" name="get_race" id="get_race" value="Get Race Info">
		</div>
		</form>
	<%End If%>
			
	<%If Not CLng(lRaceID) = 0 Then%>
		<div class="bg-critical">
			<a href="javascript:window.print()">Print These</a>
		</div>

		<div class="col-sm-6">
            <%Call GetRcds(lRaceID, "m")%>
            <h4 class="h4">Male Division Records</h4>
			<table class="table table-striped">
				<tr>
					<td>Division</td>
					<td>Record Setter</td>
					<td>Age</td>
					<td>Record</td>
					<td>Year</td>
				</tr>
				<%For i = 0 to UBound(Records, 2) - 1%>
					<tr>
						<td><%=Records(0, i)%></td>
						<td><%=Records(1, i)%></td>
						<td><%=Records(2, i)%></td>
						<td><%=Records(3, i)%></td>
						<td><%=Records(4, i)%></td>
					</tr>
                    <%If UBound(Records, 2) <= 2 Then Exit For%>
				<%Next%>
			</table>
		</div>
		<div class="col-sm-6">
            <%Call GetRcds(lRaceID, "f")%>
            <h4 class="h4">Female Division Records</h4>
			<table class="table table-striped">
				<tr>
					<th>Division</th>
					<th>Record Setter</th>
					<th>Age</th>
					<th>Record</th>
					<th>Year</th>
				</tr>
				<%For i = 0 to UBound(Records, 2) - 1%>
					<tr>
						<td><%=Records(0, i)%></td>
						<td><%=Records(1, i)%></td>
						<td><%=Records(2, i)%></td>
						<td><%=Records(3, i)%></td>
						<td><%=Records(4, i)%></td>
					</tr>
                    <%If UBound(Records, 2) <= 2 Then Exit For%>
				<%Next%>
			</table>
		</div>
	<%End If%>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
