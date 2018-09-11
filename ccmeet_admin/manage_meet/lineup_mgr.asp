<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, sql2, rs2
Dim i, j, k
Dim lTeamID, lThisMeet, lMyID, lCoachID
Dim RosterArr(), DeleteArr(), RaceAssign(), TempArr(6), BibRange()
Dim sGender, sGradeYear, sOrderBy, sMeetName, sSport, sMeetInfoSheet, sMapLink, sCourseMap, sTeamName, sPopulateBibs
Dim iGrade, iNumRaces, iThisBib
Dim dShutdown, dMeetDate
Dim sErrMsg
Dim bDuplBib

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

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

'get coach info
sql = "SELECT CoachesID FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
lCoachID = rs(0).Value
Set rs = Nothing

If Request.Form.Item("submit_populate") = "submit_populate" Then
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
		If Request.Form.Item("race_" & rs(0).Value) = "none" Then
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
			RaceAssign(2, i) = "y"			'indicate that this was already applied
			If Not RaceAssign(3, i) = vbNullString Then rs(1).Value = RaceAssign(3, i)
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
				sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID, Bib) VALUES (" & lThisMeet & ", " & RaceAssign(0, i)
				sql = sql & ", " & RaceAssign(1, i) & ", " & RaceAssign(3, i) & ")"
				Set rs = conn.Execute(sql)
				Set rs = Nothing
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
End If

'get coach info
sql = "SELECT PopulateBibs FROM Coaches WHERE CoachesID = " & lCoachID
Set rs = conn.Execute(sql)
sPopulateBibs = rs(0).Value
Set rs = Nothing

'get team gender
i = 0
ReDim RosterArr(6, 0)
sql = "SELECT Gender, TeamName FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sGender = rs(0).Value
sTeamName = Replace(rs(1).Value, "''", "'")
Set rs = Nothing
	
'convert gender to full word
Select Case sGender
	Case "M"
		sGender = "Male"
	Case "F"
		sGender = "Female"
End Select

sql = "SELECT RosterID, FirstName, LastName, Gender FROM Roster WHERE TeamsID = " & lTeamID
sql = sql & " AND Archive = 'n' ORDER BY LastName, FirstName"
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
ElseIf sOrderBy = "race-name" Then
	For i = 0 to UBound(RosterArr, 2) - 2
		For j = i + 1 to UBound(RosterArr, 2) - 1
			If CLng(RosterArr(6, i)) > CLng(RosterArr(6, j)) Then
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
sql = "SELECT FirstBib, LastBib FROM TeamBibs WHERE TeamsID = " & lTeamID
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

iNumRaces = 1
i = 0
ReDim RaceArr(1, 0)
'get meet name
sql = "SELECT MeetName, MeetDate, WhenShutdown, Sport FROM Meets WHERE MeetsID = " & lThisMeet
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
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisMeet & " AND (Gender = '" & sGender & "' OR Gender = 'Open') ORDER BY ViewOrder"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	iNumRaces = iNumRaces + 1
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
	
Private Function UpdateGrade(lMyID)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
	rs2.Open sql2, conn, 1, 2
	rs2(0).Value = Request.Form.Item("grade_" & lMyID)
	rs2.Update
	rs2.Close
	Set rs2 = Nothing
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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country Line-Up Manager</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body style="background-color: #fff;">
<div style="margin: 10px;padding: 10px;font-size: 0.9em;text-align: left;">
    <h4 class="h4">Gopher State Events Cross-Country Line-Up Manager:  <%=sTeamName%></h4>
			
	<%If Not sErrMsg = vbNullString Then%>
		<p><%=sErrMsg%></p>
	<%End If%>

	<form class="form-inline" name="remember_bibs" method="post" action="lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>&amp;order_by=<%=sOrderBy%>">
    <label for="populate_bibs">Remember Past Bibs?</label>
    <select class="form-control" name="populate_bibs" id="populate_bibs" onchange="this.form.submit1.click();">
        <%If sPopulateBibs = "y" Then%>
            <option value="n">No</option>
            <option value="y" selected>Yes</option>
        <%Else%>
            <option value="n">No</option>
            <option value="y">Yes</option>
        <%End If%>
    </select>
 	<input type="hidden" name="submit_populate" id="submit_populate" value="submit_populate">
	<input type="submit" class="form-control" name="submit1" id="submit1" value="Submit This">
    </form>
    <br>				
	<ul class="list-inline bg-success">
		<li>
            <a href="javascript:pop('/cc_meet/coach/meets/relay_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>',610,700)">Relays</a>
		</li>
		<li>
            <a href="javascript:pop('meet_sheet.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>',750,700)">Meet Sheet</a>
		</li>
		<li>
            <a href="javascript:pop('blnk_mt_sht.asp',750,700)">Blank Meet Sheet</a>
        </li>
		<%If Not sMapLink = vbNullString Then%>
			<li>
			    <a href="javascript:pop('<%=sMapLink%>',1024,768)">MapQuest Link to Site</a>
            </li>
		<%End If%>
		<%If Not sMeetInfoSheet = vbNullString Then%>
			<li>
			    <a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/info_sheets/<%=sMeetInfoSheet%>',1024,768)">Info Sheet</a>
            </li>
		<%End If%>
		<%If Not sCourseMap = vbNullString Then%>
			<li>
			    <a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/course_maps/<%=sCourseMap%>',1024,768)">Course Map</a>
            </li>
		<%End If%>
		<li>
		    <a href="lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>&amp;order_by=bib">Sort By Bib</a>
		</li>
		<li>
            <a href="lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>">Sort By Name</a>
		</li>
		<li>
            <a href="lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>&amp;order_by=race-name">Sort By Race-Name</a>
		</li>
		<li>
            <a href="lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>&amp;order_by=grade-name">Sort By Grade-Name</a>
        </li>
	</ul>

    <div class="col-xs-10">
	    <form class="form" name="assign_races" method="post" action="lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>&amp;order_by=<%=sOrderBy%>">
	    <table class="table table-striped">
		    <tr>
			    <td colspan="<%=iNumRaces + 4%>">
				    <input type="hidden" name="submit_lineup" id="submit_lineup" value="submit_lineup">
				    <input type="submit" class="form-control" name="submit2" id="submit2" value="Click Here To Save Line-Up Changes">
			    </td>
		    </tr>
		    <tr>
			    <th rowspan="2">No</th>
			    <th rowspan="2">Name</th>
			    <th rowspan="2">Gr</th>
			    <th rowspan="2">Bib</th>
			    <th colspan="<%=iNumRaces + 1%>">Race(s)</th>
		    </tr>
		    <tr>
			    <td>None</td>
			    <%For i = 0 to UBound(RaceArr, 2) - 1%>
				    <td><a href="javascript:pop('../../meet_dir/races/race_details.asp?race_id=<%=RaceArr(0, i)%>',400,250)"><%=RaceArr(1, i)%></a></td>
			    <%Next%>
		    </tr>
		    <%For i = 0 to UBound(RosterArr, 2) - 1%>
				<tr>
					<td><%=i + 1%>)</td>
					<td><%=RosterArr(2, i)%>,&nbsp;<%=RosterArr(1, i)%></td>
					<td><%=RosterArr(3, i)%></td>
					<td>
						<%If sSport = "Cross-Country" Then%>
                            &nbsp;
                        <%Else%>
                            <%If UBound(OurBibs) > 0 Then%>
                                <select name="bib_<%=RosterArr(0, i)%>" id="bib_<%=RosterArr(0, i)%>">
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
                                &nbsp;
                            <%End If%>
                        <%End If%>
					</td>
					<td>
						<input type="radio" name="race_<%=RosterArr(0, i)%>" id="race_<%=RosterArr(0, i)%>" value="none" 
									checked>
					</td>
					<%For j = 0 to UBound(RaceArr, 2) - 1%>
						<td>
							<input type="radio" name="race_<%=RosterArr(0, i)%>" id="race_<%=RosterArr(0, i)%>" 
										value="<%=RaceArr(0, j)%>"
								<%If CLng(GetRace(RosterArr(0, i))) = CLng(RaceArr(0, j)) Then%>
								checked
								<%End If%>
							>
						</td>
					<%Next%>
				</tr>
		    <%Next%>
	    </table>
	    </form>
    </div>
    <div class="col-xs-2 bg-info">
        <h5 class="h5">Avail Bibs</h5>

        <ul class="list-group">
            <%For i = 0 To UBound(AvailBibs) - 1%>
                <li class="list-group-item"><%=AvailBibs(i)%></li>
            <%Next%>
        </ul>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
