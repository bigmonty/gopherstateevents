<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, sql2, rs2
Dim i, j, k, m, n, p
Dim lTeamID, lThisMeet, lThisRace, lMeetClass, lCoachID
Dim TeamsArr(), RosterArr(), DeleteArr(), RaceAssign(), MeetsReg(), TempArr(6), RaceArr(), BibRange(), MeetClasses()
Dim sCoachName, sGender,sGradeYear, sOrderBy, sMeetName, sSport, sMeetInfoSheet, sMapLink, sCourseMap, sPopulateBibs, sLockClasses, sHasRelay
Dim iGrade, iThisBib
Dim dShutdown, dMeetDate
Dim sErrMsg
Dim bDuplBib, bHasRelays

If Not (Session("role") = "coach" Or Session("role") = "team_staff") Then Response.Redirect "/default.asp?sign_out=y"

If Session("role") = "coach" Then
    lCoachID = Session("my_id")
Else
    lCoachID = Session("team_coach_id")
End If

lTeamID = Request.QueryString("team_id")
lThisMeet = Request.QueryString("meet_id")
sOrderBy = Request.QueryString("order_by")
 
'get year for roster grades
If Month(Date) <=5 Then
	sGradeYear = Right(CStr(Year(Date) - 1), 2)
Else
	sGradeYear = Right(CStr(Year(Date)), 2)	
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_class") = "submit_class" Then
    lMeetClass = Request.Form.Item("meet_classes")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet & " AND TeamsID = " & lTeamID
    rs.Open sql, conn, 1, 2
    If lMeetClass & "" = "" Then
        rs(0).Value = Null
    Else
        rs(0).Value = lMeetClass
    End If
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_populate") = "submit_populate" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PopulateBibs FROM Coaches WHERE CoachesID = " & lCoachID
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("populate_bibs")
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_lineup") = "submit_lineup" Then
	i = 0
	j = 0
	ReDim DeleteArr(0)
	ReDim RaceAssign(3, 0)
	sql = "SELECT r.RosterID FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE TeamsID = " & lTeamID & " AND r.Archive = 'n' "
	sql = sql & "ORDER BY r.LastName, g.Grade" & sGradeYear
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		If Request.Form.Item("race_" & rs(0).Value) = "0" Then
			DeleteArr(j) = rs(0).value			'mark for deletion from this meet if they are entered
			j = j + 1
			ReDim Preserve DeleteArr(j)
		Else
			RaceAssign(0, i) = Request.Form.Item("race_" & rs(0).Value)
			RaceAssign(1, i) = rs(0).Value
			RaceAssign(2, i) = "n"		'use as a flag to indicate that this exists and does not need to be inserted
			RaceAssign(3, i) = Request.Form.Item("bib_" & rs(0).Value)
			i = i + 1
			ReDim Preserve RaceAssign(3, i)
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	'overwrite existing race assignment for this meet/participant if they exist
	For i = 0 to UBound(RaceAssign, 2) - 1
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RacesID, Bib FROM IndRslts WHERE RosterID = " & RaceAssign(1, i) & " AND MeetsID = " & lThisMeet
		rs.Open sql, conn, 1, 2
		If rs.recordcount > 0 Then
			rs(0).Value = RaceAssign(0, i)
			RaceAssign(2, i) = "y"			'indicate that they have been entered
            If RaceAssign(3, i) & "" = "" Then
                rs(1).Value = 0
            Else
                If DuplBib(RaceAssign(3, i), RaceAssign(1, i)) = False Then 'prevent them from changing the bib of an existing entrant to an assigned bib
                    rs(1).Value = RaceAssign(3, i)
                End If
            End If
			rs.Update
		End If		
		rs.Close
		Set rs = Nothing
	Next
	
	'enter race if they do not exist
	For i = 0 to UBound(RaceAssign, 2) - 1
		If RaceAssign(2, i) = "n" Then
			If RaceAssign(3, i) & "" = "" Then
				sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID) VALUES (" & lThisMeet & ", " & RaceAssign(0, i)
				sql = sql & ", " & RaceAssign(1, i) & ")"
				Set rs = conn.Execute(sql)
				Set rs = Nothing
			Else
                If DuplBib(RaceAssign(3, i), RaceAssign(1, i)) = False Then
				    sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID, Bib) VALUES (" & lThisMeet & ", " & RaceAssign(0, i)
				    sql = sql & ", " & RaceAssign(1, i) & ", " & RaceAssign(3, i) & ")"
				    Set rs = conn.Execute(sql)
				    Set rs = Nothing
                Else
				    sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID) VALUES (" & lThisMeet & ", " & RaceAssign(0, i)
				    sql = sql & ", " & RaceAssign(1, i) & ")"
				    Set rs = conn.Execute(sql)
				    Set rs = Nothing
                End If
			End If
		End If
	Next
	
	'now delete those that are so marked
	For i = 0 to UBound(DeleteArr) - 1
		sql = "DELETE FROM IndRslts WHERE RosterID = " & DeleteArr(i) & " AND MeetsID = " & lThisMeet
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	Next
		
	'now check for duplicate bibs in the race for this team in this meet
	bDuplBib = False
	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT ir.Bib FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.MeetsID = " & lThisMeet 
	sql = sql & " AND r. TeamsID = " & lTeamID & " AND ir.Bib <> 0 ORDER BY ir.Bib"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		Do While Not rs.EOF
			If i = 0 Then
				iThisBib = rs(0).Value
			Else
				If CInt(iThisBib) = rs(0).Value Then
					bDuplBib = True
					Exit Do
				Else
					iThisBib = rs(0).Value
				End If
			End If
			i = 1
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
	
	If bDuplBib = True Then
		sErrMsg = "There is at least one duplicate bib number assigned to your participants for this meet.  The bib number that "
		sErrMsg = sErrMsg & "was initially identified is " & iThisBib & ".  There may be more.  This will cause issues with "
		sErrMsg = sErrMsg & "results processing and/or extra time on the part of our staff in preparing for this meet.  Please "
		sErrMsg = sErrMsg & "look your assignments over carefully and re-assign as needed."
		
		sOrderBy = "bib"
	End if
ElseIf Request.Form.Item("get_meet") = "get_meet" Then
	lThisMeet = Request.Form.Item("meets")
ElseIf Request.Form.Item("get_team") = "get_team" Then
	lTeamID = Request.Form.Item("teams")
End If

If CStr(lTeamID) = vbNullString Then lTeamID = 0
If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

'get coach info
sql = "SELECT FirstName, LastName, PopulateBibs FROM Coaches WHERE CoachesID = " & lCoachID
Set rs = conn.Execute(sql)
sCoachName = rs(0).Value & " " & rs(1).Value
sPopulateBibs = rs(2).Value
Set rs = Nothing

'get teams
i = 0
ReDim TeamsArr(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamName, Gender, Sport FROM Teams WHERE CoachesID = " & lCoachID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	TeamsArr(0, i) = rs(0).value 
	TeamsArr(1, i) = rs(1).Value & " (" & rs(2).Value & ")"
    TeamsArr(2, i) = rs(3).Value
	i = i + 1
	ReDim Preserve TeamsArr(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If UBound(TeamsArr, 2) = 1 Then lTeamID = TeamsArr(0, 0)

'get team gender
i = 0
ReDim RosterArr(6, 0)
ReDim MeetsReg(1, 0)
If Not CLng(lTeamID) = 0 Then
	sql = "SELECT Gender FROM Teams WHERE TeamsID = " & lTeamID
	Set rs = conn.Execute(sql)
	sGender = rs(0).Value
	Set rs = Nothing
	
	'convert gender to full word
	Select Case sGender
		Case "M"
			sGender = "Male"
		Case "F"
			sGender = "Female"
	End Select
	sql = "SELECT RosterID, FirstName, LastName, Gender FROM Roster WHERE TeamsID = " & lTeamID & " AND Archive = 'n' ORDER BY LastName, FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		RosterArr(0, i) = rs(0).Value
		RosterArr(1, i) = Replace(rs(1).Value, "''", "'")
		RosterArr(2, i) = Replace(rs(2).Value, "''", "'")
		RosterArr(3, i) = GetGrade(rs(0).Value)
		RosterArr(4, i) = rs(3).Value
		RosterArr(5, i) = BibToShow(rs(0).Value)
        RosterArr(6, i) = GetRace(rs(0).Value)
  		i = i + 1
		ReDim Preserve RosterArr(6, i)
		rs.MoveNext
	Loop
	Set rs = Nothing

	're-order if ordering by bib
	If sOrderBy = "bib" Then
		For i = 0 to UBound(RosterArr, 2) - 2
			For j = i + 1 to UBound(RosterArr, 2) - 1
				If RosterArr(5, i) & "" = "" Or CInt(RosterArr(5, i)) > CInt(RosterArr(5, j)) Then
					For k = 0 to 6
						TempArr(k) = RosterArr(k, i)
						RosterArr(k, i) = RosterArr(k, j)
						RosterArr(k, j) = TempArr(k)
					Next
				End IF
			Next
		Next
    ElseIf sOrderBy = "grade-name" Then
	    For i = 0 to UBound(RosterArr, 2) - 2
		    For j = i + 1 to UBound(RosterArr, 2) - 1
			    If CLng(RosterArr(3, i)) < CLng(RosterArr(3, j)) Then
				    For k = 0 to 6
					    TempArr(k) = RosterArr(k, i)
					    RosterArr(k, i) = RosterArr(k, j)
					    RosterArr(k, j) = TempArr(k)
				    Next
			    End IF
		    Next
	    Next
	End If

    'get bib range
    i = 0
    ReDim BibRange(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FirstBib, LastBib FROM TeamBibs WHERE TeamsID = " & lTeamID & " ORDER BY FirstBib"
    rs.Open sql, conn, 1,  2
    Do While Not rs.EOF
        BibRange(0, i) = rs(0).Value
        BibRange(1, i) = rs(1).Value
        i = i + 1
        ReDim Preserve BibRange(1, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get avail bib list
    Dim bBibFound
    Dim OurBibs(), AvailBibs()

    j = 0
    ReDim OurBibs(0)
    For i = 0 To UBound(BibRange, 2) - 1
        For k = BibRange(0, i) To BibRange(1, i)
            OurBibs(j) = k
            j = j + 1
            ReDim Preserve OurBibs(j)
        Next
    Next

    j = 0
    ReDim AvailBibs(0)
    For i = 0 To UBound(OurBibs) - 1
        bBibFound = False

        For k = 0 To UBound(RosterArr, 2) - 1
            If OurBibs(i) = RosterArr(5, k) Then
                bBibFound = True
                Exit For
            End If
        Next

        If bBibFound = False Then
            AvailBibs(j) = OurBibs(i)
            j = j + 1
            ReDim Preserve AvailBibs(j)
        End If
    Next

	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID "
	sql = sql & "WHERE mt.TeamsID = " & lTeamID & " ORDER BY m.MeetDate DESC"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		MeetsReg(0, i) = rs(0).Value
		MeetsReg(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
		i = i + 1
		ReDim Preserve MeetsReg(1, i)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End If

ReDim RaceArr(1, 0)

If Not CLng(lThisMeet) = 0 Then
	'get meet name
	sql = "SELECT MeetName, MeetDate, WhenShutdown, Sport, LockClasses FROM Meets WHERE MeetsID = " & lThisMeet
	Set rs = conn.Execute(sql)
	sMeetName = rs(0).Value & " on " & rs(1).Value 
	
	'get year for roster grades
	If Month(rs(1).Value) <=7 Then
		sGradeYear = Right(CStr(Year(rs(1).Value) - 1), 2)
	Else
		sGradeYear = Right(CStr(Year(rs(1).Value)), 2)	
	End If
	
	dShutdown = rs(2).Value
	sSport = rs(3).Value
    sLockClasses = rs(4).Value
	Set rs = Nothing

    If sSport = "Cross-Country" Then Response.Redirect "cc_lineup_mgr.asp?meet_id=" & lThisMeet & "&team_id=" & lTeamID
    i = 0
    ReDim MeetClasses(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClassesID, ClassName, Gender, Details FROM MeetClasses WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MeetClasses(0, i) = rs(0).Value
        MeetClasses(1, i) = Replace(rs(1).Value, "''", "'")
        MeetClasses(2, i) = rs(2).Value
        MeetClasses(3, i) = rs(3).Value
        i = i + 1
        ReDim Preserve MeetClasses(3, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet & " AND TeamsID = " & lTeamID
    rs.Open sql, conn, 1, 2
    lMeetClass = rs(0).Value
    rs.Close
    Set rs = Nothing

    If lMeetClass & "" = "" Then lMeetClass = 0

    sHasRelay = "n"
    i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RacesID, RaceDesc, IndivRelay FROM Races WHERE MeetsID = " & lThisMeet & " AND (Gender = '" & sGender & "' OR Gender = 'Open') "
    sql = sql & "AND (RaceClass IS NULL OR RaceClass = " & lMeetClass & ") ORDER BY ViewOrder"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
        If sHasRelay = "n" Then
            If rs(2).Value = "Relay" Then sHasRelay = "y"
        End If
		RaceArr(0, i) = rs(0).Value
		RaceArr(1, i) = rs(1).Value
		i = i + 1
		ReDim Preserve RaceArr(1, i)
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
End If

'get races this part is entered for this meet
Function GetRace(lThisPart)	
	GetRace = 0
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RacesID FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND RosterID = " & lThisPart
	rs2.Open sql2, conn, 1, 2
	If rs2.recordcount > 0 Then
		GetRace = rs2(0).Value
	Else
		GetRace = 0
	End If
	rs2.Close
	Set rs2 = Nothing
End Function

'get races this part is entered for this meet
Private Function BibToShow(lThisPart)	
	BibToShow = 0
	
	'first see if a bib has been assigned to this participant for this event
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND RosterID = " & lThisPart
	rs2.Open sql2, conn, 1, 2
	If rs2.recordcount > 0 Then BibToShow = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
	
	'if no bib has been assigned, then show their most recent bib if available
	If BibToShow = 0 Then
        If sPopulateBibs = "y" Then
		    Set rs2 = Server.CreateObject("ADODB.Recordset")
		    sql2 = "SELECT ir.Bib FROM IndRslts ir INNER JOIN Meets m ON ir.MeetsID = m.MeetsID WHERE RosterID = " 
		    sql2 = sql2 & lThisPart & " ORDER BY m.MeetDate DESC"
		    rs2.Open sql2, conn, 1, 2
		    If rs2.recordcount > 0 Then BibToShow = rs2(0).Value
		    rs2.Close
		    Set rs2 = Nothing
        End If
	End If
End Function
	
Private Function GetGrade(lMyID)
    GetGrade = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	If Month(Date) <= 5 Then
        sql2 = "SELECT Grade" & Right(CStr(Year(Date) - 1), 2) & " FROM Grades WHERE RosterID = " & lMyID
    Else
        sql2 = "SELECT Grade" & Right(CStr(Year(Date)), 2) & " FROM Grades WHERE RosterID = " & lMyID
    End If
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then  GetGrade = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function

Private Function DuplBib(iThisBib, lMyID)
    DuplBib = False

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Bib FROM IndRslts WHERE RosterID <> " & lMyID & " AND MeetsID = " & lThisMeet & " AND Bib = " & iThisBib
	rs2.Open sql2, conn, 1, 2
	If rs2.recordcount > 0 Then DuplBib = True
	rs2.Close
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Line-Up Manager-Coach Version</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<h4 class="h4">CC/Nordic Meet Line-Up Manager:  <%=sCoachName%>!</h4>
			<div class="row">
				<div class="col-sm-6">
					<form role="form" class="form-inline" name="get_team" method="post" action="lineup_mgr.asp">
					<label for="teams">Team:</label>
					<select class="form-control" name="teams" id="teams" onchange="this.form.submit2.click();">
						<option value="0">&nbsp;</option>
						<%For i = 0 to UBound(TeamsArr, 2) - 1%>
							<%If CLng(TeamsArr(0, i)) = CLng(lTeamID) Then%>
								<option value="<%=TeamsArr(0, i)%>" selected><%=TeamsArr(1, i)%> - <%=TeamsArr(2, i)%></option>
							<%Else%>
								<option value="<%=TeamsArr(0, i)%>"><%=TeamsArr(1, i)%> - <%=TeamsArr(2, i)%></option>
							<%End If%>
						<%Next%>
					</select>
					<input type="hidden" name="get_team" id="get_team" value="get_team">
					<input class="form-control" type="submit" name="submit2" id="submit2" value="Go">
					</form>
				</div>	
				<div class="col-sm-6">
					<%If Not CLng(lTeamID) = 0 Then%>
						<form role="form" class="form-inline" name="get_meet" method="post" action="lineup_mgr.asp?team_id=<%=lTeamID%>">
						<label for="meets">Meet:</label>
						<select class="form-control" name="meets" id="meets" onchange="this.form.submit3.click();">
							<option value="0">&nbsp;</option>
							<%For i = 0 to UBound(MeetsReg, 2) - 1%>
								<%If CLng(MeetsReg(0, i)) = CLng(lThisMeet) Then%>
									<option value="<%=MeetsReg(0, i)%>" selected><%=MeetsReg(1, i)%></option>
								<%Else%>
									<option value="<%=MeetsReg(0, i)%>"><%=MeetsReg(1, i)%></option>
								<%End If%>
							<%Next%>
						</select>
						<input type="hidden" name="get_meet" id="get_meet" value="get_meet">
						<input class="form-control" type="submit" name="submit3" id="submit3" value="Go">
						</form>
					<%End If%>
				</div>	
			</div>

			<%If Not CLng(lThisMeet) = 0 Then%>
				<%If UBound(MeetClasses, 2) > 0 Then%>
					<h4 class="h4">This Meet Has Meet Class(es)!</h4>
										
					<p class="bg-warning">
						This meet has classes for some or all of the races in the meet.  Each team in the meet must have their class designated in 
						order to enter participants in class-based races. PLEASE BE CAREFUL TO DO THIS ACCURATELY BASED ON THE CLASS DESCRIPTION.  
						Otherwise your team members could be entered into the incorrect race(s).
					</p>
					<table class="table table-striped">
						<tr>
							<td valign="top">
								<h4 class="h4">Meet Classes</h4>
								<ol>
									<%For i = 0 To UBound(MeetClasses, 2) -1%>
										<li><%=MeetClasses(1, i)%> (<%=MeetClasses(3, i)%>)</li>
									<%Next%>
								</ol>

								<p>This information is determined by official MSHSL numbers as 
									found <a href="http://www.mshsl.org/mshsl/enrollments15.asp" onclick="openThis(this.href,1024,768);return false;">here.</a></p>
							</td>
							<td valign="top">
								<h4 class="h4">This Team's Class</h4>
								<form role="form" class="form-inline" name ="designate_class" method="post" action="lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>&amp;order_by=<%=sOrderBy%>">
								<select class="form-control" name="meet_classes" id="meet_classes">
									<option value="">&nbsp;</option>
									<%For i = 0 To UBound(MeetClasses, 2) - 1%>
										<%If CLng(lMeetClass) = CLng(MeetClasses(0, i)) Then%>
											<option value="<%=MeetClasses(0, i)%>" selected><%=MeetClasses(1, i)%></option>
										<%Else%>
											<option value="<%=MeetClasses(0, i)%>"><%=MeetClasses(1, i)%></option>
										<%End If%>
									<%Next%>
								</select>
								<%If sLockClasses = "n" Then%>
									<input class="form-control" type="hidden" name="submit_class" id="submit_class" value="submit_class">
									<input class="form-control" type="submit" name="submit_team_class" id="submit_team_class" value="Submit Class">
								<%End If%>
								</form>
								<%If sLockClasses = "y" Then%>
									<p class="bg-danger">
										Classes for this meet have been locked.  Your team's class, if designated, should appear above.  
										If it looks incorrect, or if it is not designated within a week of the meet, please notify 
										<a href="mailto:bob.schneider@gopherstateevents.com">bob.schneider@gopherstateevents.com</a>.
									</p>
								<%End If%>
							</td>
						</tr>
					</table>
				<%End If%>

				<p class="bg-warning">
					<span style="font-weight:bold;">RE-ASSIGNING ATHLETES TO A DIFFERENT RACE:</span> To do this, remove them from the race that they are in.  That will put them in the 
					"Unassigned" pool at the bottom.  Then go in and re-assign them to the race you wish.
				</p>

				<%If sSport = "Nordic Ski" Then%>
					<div class="row">
						<div class="col-sm-4">
							<form role="form" class="form-inline" name="remember_bibs" method="post" action="lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>&amp;order_by=<%=sOrderBy%>">
							<div class="form-group" style="padding: 5px;">
								<label>Remember Past Bibs?</label>
								<select class="form-control" name="populate_bibs" id="populate_bibs" onchange="this.form.submit3a.click();">
									<%If sPopulateBibs = "y" Then%>
										<option value="n">No</option>
										<option value="y" selected>Yes</option>
									<%Else%>
										<option value="n">No</option>
										<option value="y">Yes</option>
									<%End If%>
								</select>
								<input class="form-control" type="hidden" name="submit_populate" id="submit_populate" value="submit_populate">
								<input class="form-control" type="submit" name="submit3a" id="submit3a" value="Submit This">
							</div>
							</form>
						</div>
						<div class="col-sm-8">
							<p>
								NOTE:  If an attempt is made to enter a bib that has already been assigned, no bib will be assigned to that
								participant and no warning will be posted.  Their bib window will be left blank.
							</p>
						</div>
					</div>
				<%End If%>

				<%If Not sErrMsg = vbNullString Then%>
					<p class="bg-danger"><%=sErrMsg%></p>
				<%End If%>

				<div class="row">					
					<ul class="nav">
						<%If sHasRelay = "y" Then%>
							<li class="nav-item"><a class="nav-link" href="javascript:pop('relay_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>',610,700)">Relay Manager</a></li>
						<%End If%>
						<li class="nav-item"><a class="nav-link" href="javascript:pop('meet_sheet.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>',750,700)">Meet Sheet</a></li>
						<%If sSport = "Nordic Ski" Then%>
							<li class="nav-item"><a class="nav-link" href="javascript:pop('bib_list.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>',750,700)">Bib List</a></li>
						<%End If%>
						<li class="nav-item"><a class="nav-link" href="javascript:pop('blnk_mt_sht.asp',750,700)">Blank Meet Sheet</a></li>
						<%If Not sMapLink = vbNullString Then%>
							<li class="nav-item"><a class="nav-link" href="javascript:pop('<%=sMapLink%>',1024,768)">Map Link</a></li>
						<%End If%>
						<%If Not sMeetInfoSheet = vbNullString Then%>
							<li class="nav-item"><a class="nav-link" href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/info_sheets/<%=sMeetInfoSheet%>',1024,768)">Info Sheet</a></li>
						<%End If%>
						<%If Not sCourseMap = vbNullString Then%>
							<li class="nav-item"><a class="nav-link" href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/course_maps/<%=sCourseMap%>',1024,768)">Course Map</a></li>
						<%End If%>
						<li class="nav-item"><a class="nav-link" href="lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>&amp;order_by=bib">Sort: Bib</a>
						<li class="nav-item"><a class="nav-link" href="lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>">Sort: Name</a></li>
						<li class="nav-item"><a class="nav-link" href="lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>&amp;order_by=grade-name">Sort: Grade-Name</a></li>
						<%If Date >= CDate(dMeetDate) Then%>
							<li class="nav-item"><a class="nav-link" href="our_rslts.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>">Our Results</a></li>
						<%End If%>
					</ul>
				</div>
				
				<div class="row">
					<div class="col-sm-8">
						<h4 class="h4">Manage Your Line-Up</h4>
						<form role="form" class="form" name="assign_races" method="post" action="lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>&amp;order_by=<%=sOrderBy%>">
						<div class="form-group">
							<table class="table table-striped">
								<tr>
									<td colspan="5">
										<input class="form-control" type="hidden" name="submit_lineup" id="submit_lineup" value="submit_lineup">
										<%If Now() >= dShutdown Then%>
											<input class="form-control" type="submit" name="submit2" id="submit2" value="Save Line-Up/Bib Changes" disabled>
										<%Else%>
											<input class="form-control" type="submit" name="submit2" id="submit2" value="Save Line-Up/Bib Changes">
										<%End If%>

										<%If Now() >= dShutdown Then%>
											<p class="bg-success">
												We're sorry.  Line-up changes for this event have closed.  Subject to local meet management
												approval, you may make changes at the meet.
											</p>
										<%End If%>
									</td>
								</tr>
								<%For m = 0 To UBound(RaceArr, 2) - 1%>
									<tr>
										<th colspan="5">
											<a href="javascript:pop('../../meet_dir/races/race_details.asp?race_id=<%=RaceArr(0, m)%>',400,250)"
												style="color: #892700;"><%=RaceArr(1, m)%></a>
										</th>
									</tr>
									<tr>
										<th>No</th>
										<th>Name</th>
										<th>Gr</th>
										<th>Bib</th>
										<th>Race</th>
									</tr>
									<%n = 0%>
									<%For i = 0 to UBound(RosterArr, 2) - 1%>
										<%If CLng(RosterArr(6, i)) = CLng(RaceArr(0, m)) Then%>
											<tr>
												<td><%=n + 1%>)</td>
												<td><%=RosterArr(2, i)%>,&nbsp;<%=RosterArr(1, i)%></td>
												<td><%=RosterArr(3, i)%></td>
												<td>
													<%If sSport = "Nordic Ski" Then%>
														<select class="form-control" name="bib_<%=RosterArr(0, i)%>" id="bib_<%=RosterArr(0, i)%>">
															<option value="">&nbsp;</option>
															<%For j = OurBibs(0) to OurBibs(UBound(OurBibs) - 1)%>
																<%If CInt(j) = CInt(RosterArr(5, i)) Then%>
																	<option value="<%=j%>" selected><%=j%></option>
																<%Else%>
																	<%For k = 0 To UBound(AvailBibs) - 1%>
																		<%If CInt(AvailBibs(k)) = CInt(j) Then%>
																			<option value="<%=j%>"><%=j%></option>
																			<%Exit For%>
																		<%End If%>
																	<%Next%>
																<%End If%>
															<%Next%>
														</select>
													<%Else%>    
														na
													<%End If%>
												</td>
												<td>
													<select class="form-control" name="race_<%=RosterArr(0, i)%>" id="race_<%=RosterArr(0, i)%>">
														<option value="0">Remove</option>
														<%For p = 0 To UBound(RaceArr, 2) - 1%>
															<%If CLng(RosterArr(6, i)) = CLng(RaceArr(0, p)) Then%>
																<option value="<%=RaceArr(0, p)%>" selected><%=RaceArr(1, p)%></option>
															<%Else%>
																<option value="<%=RaceArr(0, p)%>"><%=RaceArr(1, p)%></option>
															<%End If%>
														<%Next%>
													</select>
												</td>
											</tr>

											<%n=n+1%>
										<%End If%>
									<%Next%>
								<%Next%>
								<tr>
									<th colspan="5">Unassigned</th>
								</tr>
								<tr>
									<th>No</th>
									<th>Name</th>
									<th>Gr</th>
									<th>Bib</th>
									<th>Race</th>
								</tr>
								<%n = 0%>
								<%For i = 0 to UBound(RosterArr, 2) - 1%>
									<%If CLng(RosterArr(6, i)) = 0 Then%>
										<tr>
											<td><%=n + 1%>)</td>
											<td><%=RosterArr(2, i)%>,&nbsp;<%=RosterArr(1, i)%></td>
											<td><%=RosterArr(3, i)%></td>
											<td>
												<%If sSport = "Nordic Ski" Then%>
													<select class="form-control" name="bib_<%=RosterArr(0, i)%>" id="bib_<%=RosterArr(0, i)%>">
														<option value="">&nbsp;</option>
														<%For j = OurBibs(0) to OurBibs(UBound(OurBibs) - 1)%>
															<%If CInt(j) = CInt(RosterArr(5, i)) Then%>
																<option value="<%=j%>" selected><%=j%></option>
															<%Else%>
																<%For k = 0 To UBound(AvailBibs) - 1%>
																	<%If CInt(AvailBibs(k)) = CInt(j) Then%>
																		<option value="<%=j%>"><%=j%></option>
																		<%Exit For%>
																	<%End If%>
																<%Next%>
															<%End If%>
														<%Next%>
													</select>
												<%Else%>    
													na
												<%End If%>
											</td>
											<td>
												<select class="form-control" name="race_<%=RosterArr(0, i)%>" id="race_<%=RosterArr(0, i)%>">
													<option value="0">&nbsp;</option>
													<%For m = 0 To UBound(RaceArr, 2) - 1%>
														<option value="<%=RaceArr(0, m)%>"><%=RaceArr(1, m)%></option>
													<%Next%>
												</select>
											</td>
										</tr>

										<%n=n+1%>
									<%End If%>
								<%Next%>
							</table>
						</div>
						</form>
					</div>
					<div class="col-sm-4">
						<%If sSport = "Cross-Country" Then%>
							<div class="bg-danger text-danger">
								<h4 class="h4">Avoiding Missed RFID Reads</h4>

								<p>
									Our rfid system usually has about a 99.5% read rate or higher.  We often don't miss a read for several meets in a row.  This is 
									very good but any missed read will affect the results until resolved.  Help avoid them by following these guidelines.
								</p>
								<ol class="list-group">
									<li class="nav-item">Have them wear the bib horizontally on the front of the jersey.</li>
									<li class="nav-item">They should place the bib as low on the jersey as is comfortable.</li>
									<li class="nav-item">They should use four pins.</li>
									<li class="nav-item">When self-timing, they should keep their arms away from their body.</li>
									<li class="nav-item">They should not crumple or modify the bib in any way.</li>
									<li class="nav-item">They should compete in the race they are assigned and wear the bib they are assigned.  We can make changes easily.</li>
								</ol>
							</div>
						<%End If%>

						<h4 class="h4">Available Bibs</h4>

						<ul class="list-group">
							<%For i = 0 To UBound(AVailBibs) - 1%>
								<li class="nav-item"><%=AvailBibs(i)%></li>
							<%Next%>
						</ul>
					</div>
				</div>
			<%End If%>
		</div>
	</div
</div>
<!--#include file = "../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
