<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisMeet
Dim i, x, y
Dim sName, sPhone, sEmail, sCourseMap, sMeetInfoSheet, sRsltsOfficial, sDynamicRaceAssign, sSrvrIDs, sStartTime
Dim sArchive13, sGradeYear, sSport
Dim Meets(), MeetArray(16), MeetDir(), TableFields(2, 28), ArchiveArr()
Dim bRaceFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

sSrvrIDs = Request.QueryString("srvr_ids")
If sSrvrIDs = vbNullString Then sSrvrIDs = "n"

sArchive13 = Request.QueryString("archive_13")
If sArchive13 = vbNullString Then sArchive13 = "n"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get year for roster grades
If Month(Date) <=5 Then
	sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
Else
	sGradeYear = Right(CStr(Year(Date)), 2)	
End If

If sSrvrIDs = "y" Then
    'tables and fields to check
    TableFields(0, 0) = "Roster"
    TableFields(1, 0) = "RosterSrvrID"
    TableFields(2, 0) = "RosterID"
    
    TableFields(0, 1) = "Roster"
    TableFields(1, 1) = "TeamSrvrID"
    TableFields(2, 1) = "TeamsID"
    
    TableFields(0, 2) = "Teams"
    TableFields(1, 2) = "TeamSrvrID"
    TableFields(2, 2) = "TeamsID"
    
    TableFields(0, 3) = "Teams"
    TableFields(1, 3) = "CoachSrvrID"
    TableFields(2, 3) = "CoachesID"
    
    TableFields(0, 4) = "Races"
    TableFields(1, 4) = "RaceSrvrID"
    TableFields(2, 4) = "RacesID"
    
    TableFields(0, 5) = "Coaches"
    TableFields(1, 5) = "CoachSrvrID"
    TableFields(2, 5) = "CoachesID"
    
    TableFields(0, 6) = "MeetTeams"
    TableFields(1, 6) = "TeamSrvrID"
    TableFields(2, 6) = "TeamsID"
    
    TableFields(0, 7) = "MeetTeams"
    TableFields(1, 7) = "MeetTeamsSrvrID"
    TableFields(2, 7) = "MeetTeamsID"
    
    TableFields(0, 8) = "IndRslts"
    TableFields(1, 8) = "RaceSrvrID"
    TableFields(2, 8) = "RacesID"
    
    TableFields(0, 9) = "IndRslts"
    TableFields(1, 9) = "RosterSrvrID"
    TableFields(2, 9) = "RosterID"
    
    TableFields(0, 10) = "TmRslts"
    TableFields(1, 10) = "TeamSrvrID"
    TableFields(2, 10) = "TeamsID"
    
    TableFields(0, 11) = "TmRslts"
    TableFields(1, 11) = "RaceSrvrID"
    TableFields(2, 11) = "RacesID"
    
    TableFields(0, 12) = "RaceLaps"
    TableFields(1, 12) = "RaceSrvrID"
    TableFields(2, 12) = "RacesID"
    
    TableFields(0, 13) = "TimeAdjust"
    TableFields(1, 13) = "RaceSrvrID"
    TableFields(2, 13) = "RacesID"
    
    TableFields(0, 14) = "RFIDReads"
    TableFields(1, 14) = "RaceSrvrID"
    TableFields(2, 14) = "RacesID"
    
    TableFields(0, 15) = "RaceDelay"
    TableFields(1, 15) = "RaceSrvrID"
    TableFields(2, 15) = "RacesID"
     
    TableFields(0, 16) = "RaceWinners"
    TableFields(1, 16) = "RaceSrvrID"
    TableFields(2, 16) = "RacesID"
    
    TableFields(0, 17) = "RunOrder"
    TableFields(1, 17) = "RaceSrvrID"
    TableFields(2, 17) = "RacesID"
    
    TableFields(0, 18) = "Pursuit"
    TableFields(1, 18) = "RaceSrvrID"
    TableFields(2, 18) = "RacesID"
    
    TableFields(0, 19) = "Grades"
    TableFields(1, 19) = "RosterSrvrID"
    TableFields(2, 19) = "RosterID"
    
    TableFields(0, 20) = "Grades"
    TableFields(1, 20) = "GradeSrvrID"
    TableFields(2, 20) = "GradesID"
    
    TableFields(0, 21) = "Advancement"
    TableFields(1, 21) = "RaceSrvrID"
    TableFields(2, 21) = "RacesID"
    
    TableFields(0, 22) = "Advancement"
    TableFields(1, 22) = "AdvancementSrvrID"
    TableFields(2, 22) = "AdvancementID"
     
    TableFields(0, 23) = "Relays"
    TableFields(1, 23) = "RaceSrvrID"
    TableFields(2, 23) = "RacesID"
    
    TableFields(0, 24) = "RelayRslts"
    TableFields(1, 24) = "RelaySrvrID"
    TableFields(2, 24) = "RelaysID"
    
    TableFields(0, 25) = "Relays"
    TableFields(1, 25) = "RelaySrvrID"
    TableFields(2, 25) = "RelaysID"
     
    TableFields(0, 26) = "TeamBibs"
    TableFields(1, 26) = "TeamSrvrID"
    TableFields(2, 26) = "TeamsID"
    
    TableFields(0, 27) = "ScoreByPts"
    TableFields(1, 27) = "RaceSrvrID"
    TableFields(2, 27) = "RacesID"
    
    TableFields(0, 28) = "StageSettings"
    TableFields(1, 28) = "RaceSrvrID"
    TableFields(2, 28) = "RacesID"

    For y = 0 To UBound(TableFields, 2)
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT " & TableFields(2, y) & ", " & TableFields(1, y) & " FROM " & TableFields(0, y) & " WHERE " 
		sql = sql & TableFields(1, y)
		sql = sql & " = 0"
		rs.Open sql, conn, 1, 2
		Do While Not rs.EOF
			rs(1).Value = rs(0).Value
			rs.Update
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
    Next
ElseIf sArchive13 = "y" Then
	i = 0
	ReDim ArchiveArr(0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RosterID FROM Grades WHERE Grade" & sGradeYear & " >= 13"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		ArchiveArr(i) = rs(0).Value
		i = i + 1
		ReDim Preserve ArchiveArr(i)
		rs.MoveNext
	Loop
	rs.Close
	SEt rs = Nothing

	For i = 0 To UBound(ArchiveArr) - 1
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT Archive FROM Roster WHERE RosterID = " & ArchiveArr(i) & " AND Archive = 'n'"
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then
			rs(0).Value = "y"
			rs.Update
		End If
		rs.Close
		Set rs = Nothing
	Next
ElseIf Request.QueryString("update_grades") = "y" Then        
    Dim iThisYear

    iThisYear = Right(Year(Date), 2)

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Grade" & CInt(iThisYear) - 1 & ", Grade" & iThisYear & " FROM Grades WHERE Grade" & CInt(iThisYear) - 1 & " < 16"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Not rs(0).Value = 0 Then rs(1).Value = rs(0).Value + 1
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamPlace FROM IndRslts WHERE TeamPlace IS NULL"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	rs(0).Value = 0
	rs.Update
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_this") = "submit_this" Then
    If Request.Form.item("delete_meet") = "on" Then
        sql = "DELETE FROM MeetTeams WHERE MeetsID = " & lThisMeet
        Set rs = conn.Execute(sql)
        Set rs = Nothing

		sql = "DELETE FROM Meets WHERE MeetsID = " & lThisMeet 
		Set rs = conn.Execute(sql)
		Set rs = Nothing

        lThisMeet = 0
    Else
		Dim lRaceID

		sDynamicRaceAssign = Request.Form.Item("dynamic_race_assign")

	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT MeetName, MeetDate, MeetHost, WebSite, MeetSite, Weather, Comments, Sport, ShowOnline, WhenShutdown, MeetDirID,  "
	    sql = sql & "Invoice, DynamicRaceAssign FROM Meets WHERE MeetsID = " & lThisMeet
        rs.Open sql, conn, 1, 2
        If Request.Form.Item("meet_name") & "" = "" Then
		    rs(0).Value = rs(0).OriginalValue
	    Else
   		    rs(0).Value = Replace(Request.Form.Item("meet_name"), "'", "''")
	    End If
	
        If Request.Form.Item("meet_date") & "" = "" Then
		    rs(1).Value = rs(1).OriginalValue
	    Else
    	    rs(1).Value = Request.Form.Item("meet_date")
	    End If
	
        If Request.Form.Item("meet_host") & "" = "" Then
		    rs(2).Value = Null
	    Else
		    rs(2).Value = Replace(Request.Form.Item("meet_host"), "'", "''")
	    End If
	
        rs(3).Value = Request.Form.Item("website")
	
        If Request.Form.Item("meet_site") & "" = "" Then
		    rs(4).Value = Null
	    Else
    	    rs(4).Value = Replace(Request.Form.Item("meet_site"), "'", "''")
	    End If
	
        If Request.Form.Item("weather") & "" = "" Then
		    rs(5).Value = Null
	    Else
    	    rs(5).Value = Replace(Request.Form.Item("weather"), "'", "''")
	    End If
	
        If Request.Form.Item("comments") & "" = "" Then
		    rs(6).Value = Null
	    Else
    	    rs(6).Value = Replace(Request.Form.Item("comments"), "'", "''")
	    End If
	
        rs(7).Value = Request.Form.Item("sport")
		sSport = Request.Form.Item("sport")

        rs(8).Value = Request.Form.Item("show_online")
	
        If Request.Form.Item("when_shutdown") & "" = "" Then
		    rs(9).Value = rs(9).OriginalValue
	    Else
		    rs(9).Value = Request.Form.Item("when_shutdown")
	    End If
	
        rs(10).Value = Request.Form.Item("meet_dir_id")
	    rs(11).Value = Request.Form.Item("invoice")
        rs(12).Value = sDynamicRaceAssign

        rs.Update
        rs.Close
        Set rs = Nothing

		'check for b-tbd races
		bRaceFound = False
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RaceName FROM Races WHERE MeetsID = " & lThisMeet & " AND RaceName = 'B-TBD'"
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then bRaceFound = True
		rs.Close
		Set rs = Nothing

		If sDynamicRaceAssign = "y" Then
			'add if not there
			If bRaceFound = False Then
				sql = "INSERT INTO Races(MeetsID, RaceName, RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, "
				sql = sql & "ScoreMethod, NumAllow, NumScore, StartType, IndivRelay, ViewOrder) VALUES (" & lThisMeet 
				sql = sql & ", 'B-TBD', 'B-TBD', '8:00 AM', 0, 'meters', 'Male', 'Place', 0, 5, 'Mass', 'Ind', 0)"
				Set rs = conn.Execute(sql)
				Set rs = Nothing

				'get race id ServerID
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT RacesID, RaceSrvrID FROM Races WHERE MeetsID = " & lThisMeet & " AND RaceName = 'B-TBD'"
				rs.Open sql, conn, 1, 2
				lRaceID = rs(0).Value
				rs(1).Value = rs(0).Value
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
						Set rs = New ADODB.Recordset
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
					Else
						sql = "DELETE FROM Pursuit WHERE RacesID = " & lRaceID
						Set rs = conn.Execute(sql)
						Set rs = Nothing
					End If
				End If
			End If
		Else
			'delete if there
			If bRaceFound = True Then
				sql = "DELETE FROM Races WHERE RaceName = 'B-TBD' AND MeetsID = " & lThisMeet
				Set rs = conn.Execute(sql)
				Set rs = Nothing
			End If
		End If

		'check for g-tbd races
		bRaceFound = False
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RaceName FROM Races WHERE MeetsID = " & lThisMeet & " AND RaceName = 'G-TBD'"
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then bRaceFound = True
		rs.Close
		Set rs = Nothing

		If sDynamicRaceAssign = "y" Then
			'add if not there
			If bRaceFound = False Then
				sql = "INSERT INTO Races(MeetsID, RaceName, RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, "
				sql = sql & "ScoreMethod, NumAllow, NumScore, StartType, IndivRelay, ViewOrder) VALUES (" & lThisMeet 
				sql = sql & ", 'G-TBD', 'G-TBD', '8:00 AM', 0, 'meters', 'Female', 'Place', 0, 5, 'Mass', 'Ind', 0)"
				Set rs = conn.Execute(sql)
				Set rs = Nothing

				'get race id ServerID
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT RacesID, RaceSrvrID FROM Races WHERE MeetsID = " & lThisMeet & " AND RaceName = 'G-TBD'"
				rs.Open sql, conn, 1, 2
				lRaceID = rs(0).Value
				rs(1).Value = rs(0).Value
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
						Set rs = New ADODB.Recordset
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
					Else
						sql = "DELETE FROM Pursuit WHERE RacesID = " & lRaceID
						Set rs = conn.Execute(sql)
						Set rs = Nothing
					End If
				End If
			End If
		Else
			'delete if there
			If bRaceFound = True Then
				sql = "DELETE FROM Races WHERE RaceName = 'G-TBD' AND MeetsID = " & lThisMeet
				Set rs = conn.Execute(sql)
				Set rs = Nothing
			End If
		End If

	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lThisMeet
        rs.Open sql, conn, 1, 2
	    rs(0).Value = Request.form.Item("map_link")
        rs.Update
        rs.Close
        Set rs = Nothing

	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT BibStart, BibEnd FROM BibRange WHERE MeetsID = " & lThisMeet
        rs.Open sql, conn, 1, 2
        If Request.Form.Item("bib_start") & "" = "" Then
		    rs(0).Value = rs(0).OriginalValue
	    Else
		    rs(0).Value = Request.Form.Item("bib_start")
	    End If
	
        If Request.Form.Item("bib_end") & "" = "" Then
		    rs(1).Value = rs(1).OriginalValue
	    Else
		    rs(1).Value = Request.Form.Item("bib_end")
	    End If
        rs.Update
        rs.Close
        Set rs = Nothing
	
	    If Request.Form.Item("official_rslts") = "y" Then
		    sRsltsOfficial = "n"
		
		    Set rs = Server.CreateObject("ADODB.Recordset")
		    sql = "SELECT MeetsID FROM OfficialRslts WHERE MeetsID = " & lThisMeet
		    rs.Open sql, conn, 1, 2
		    If rs.RecordCount > 0 Then sRsltsOfficial = "y"
		    rs.Close
		    Set rs = Nothing
		
		    If sRsltsOfficial = "n" Then
			    sql = "INSERT INTO OfficialRslts (MeetsID) VALUES (" & lThisMeet & ")"
			    Set rs = conn.Execute(sql)
			    Set rs = Nothing
		    End If
	    Else
		    sql = "DELETE FROM OfficialRslts WHERE MeetsID = " & lThisMeet 
		    Set rs = conn.Execute(sql)
		    Set rs = Nothing
	    End If
    End If
ElseIf Request.Form.Item("get_meet") = "get_meet" Then
	lThisMeet = Request.Form.Item("meets")
End If

If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

i = 0
ReDim Meets(2, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets ORDER BY MeetDate DESC"
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

i = 0
ReDim MeetDir(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetDirID, FirstName, LastName FROM MeetDir ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	MeetDir(0, i) = rs(0).Value
	MeetDir(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve MeetDir(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Not CLng(lThisMeet) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT m.MeetName, m.MeetDate, m.MeetHost, m.WebSite, m.MeetSite, ml.MapLink, m.Weather, m.Comments, m.Sport, m.ShowOnline, m.WhenShutdown, "
	sql = sql & "br.BibStart, br.BibEnd, m.MeetDirID, m.Invoice, m.DynamicRaceAssign FROM Meets m INNER JOIN MapLinks ml ON m.MeetsID = ml.MeetsID "
	sql = sql & "INNER JOIN BibRange br on br.MeetsID = m.MeetsID WHERE m.MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	MeetArray(0) = Replace(rs(0).Value, "''", "'")
	MeetArray(1) = rs(1).Value
	MeetArray(2) = rs(2).Value
	If Not rs(3).Value & "" = "" Then MeetArray(3) = Replace(rs(3).Value, "''", "'")
	If Not rs(4).Value & "" = "" Then MeetArray(4) = Replace(rs(4).Value, "''", "'")
	If Not rs(5).Value & "" = "" Then MeetArray(5) = Replace(rs(5).Value, "''", "'")
	If Not rs(6).Value & "" = "" Then MeetArray(6) = Replace(rs(6).Value, "''", "'")
	If Not rs(7).Value & "" = "" Then MeetArray(7) = Replace(rs(7).Value, "''", "'")
	MeetArray(8) = rs(8).Value
	MeetArray(9) = rs(9).Value
	MeetArray(10) = rs(10).Value
	MeetArray(11) = rs(11).Value
	MeetArray(12) = rs(12).Value
	MeetArray(13) = rs(13).Value
	MeetArray(15) = rs(14).Value
    MeetArray(16) = rs(15).Value
    rs.Close
	Set rs = Nothing
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Phone, Email FROM MeetDir WHERE MeetDirID = " & MeetArray(13)
	rs.Open sql, conn, 1, 2
	sName =  Replace(rs(0).Value, "''", "'") & " " &  Replace(rs(1).Value, "''", "'")
	sPhone = rs(2).Value
	sEmail = rs(3).Value
	rs.Close
	Set rs = Nothing
	
	MeetArray(14) = "n"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MeetsID FROM OfficialRslts WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then MeetArray(14) = "y"
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
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT StartTime FROM RFIDSettings WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sStartTime = rs(0).Value
	rs.Close
	Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE  Admin Manage CC Meet</title>

<script>
function chkFlds() {
 	if (document.update_info.meet_name.value == '' || 
	 	document.update_info.meet_site.value == '')
		{
  		alert('Please fill in all required fields!');
  		return false
  		}
	else
		if (isNaN(document.update_info.date_month.value) ||
		   isNaN(document.update_info.date_day.value) ||
		   isNaN(document.update_info.date_year.value))
    		{
			alert('All event date fields must be numeric values');
			return false
			} 	
	else
   		return true
}
</script>
</head>
<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		
		<div class="col-sm-10">
			<%If CLng(lThisMeet) > 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>
			
			<h4 class="h4">Cross-Country/Nordic Ski Meet Manager</h4>
						
			<form name="get_meet" method="post" action="manage_meet.asp">
			Select Meet:
			<select name="meets" id="meets" onchange="this.form.submit1.click();">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Meets, 2) - 1%>
					<%If CLng(lThisMeet) = 0 Then%>
						<option value="<%=Meets(0, i)%>"><%=Meets(1, i)%> (<%=Meets(2, i)%>)</option>
					<%Else%>
						<%If CLng(Meets(0, i)) = CLng(lThisMeet) Then%>
							<option value="<%=Meets(0, i)%>" selected><%=Meets(1, i)%> (<%=Meets(2, i)%>)</option>
						<%Else%>
							<option value="<%=Meets(0, i)%>"><%=Meets(1, i)%> (<%=Meets(2, i)%>)</option>
						<%End If%>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="get_meet" id="get_meet" value="get_meet">
			<input type="submit" name="submit1" id="submit1" value="Get Meet">
			</form>
			
			<%If Not CLng(lThisMeet) = 0 Then%>
				<div style="text-align:right;font-size:0.85em;margin-top:10px;">
                    <a href="manage_meet.asp?meet_id=<%=lThisMeet%>&amp;archive_13=y">Archive Grades 13+</a>
                    &nbsp;|&nbsp;
                    <a href="manage_meet.asp?meet_id=<%=lThisMeet%>&amp;srvr_ids=y" style="color:red;">Update Srvr IDs</a>
                    &nbsp;|&nbsp;
                    <a href="javascript:pop('http://www.gopherstateevents.com/events/ccmeet_info.asp?meet_id=<%=lThisMeet%>',900,700)">Info Link</a>
                    &nbsp;|&nbsp;
					<a href="/ccmeet_admin/manage_meet/clone_prev.asp?meet_id=<%=lThisMeet%>">Clone Previous</a>
					&nbsp;|&nbsp;
					<a href="javascript:pop('/ccmeet_admin/manage_meet/upload_run_order.asp?meet_id=<%=lThisMeet%>&amp;file_sent=n',800,300)">Upload Run Order</a>
					&nbsp;|&nbsp;
					<a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/course_maps/upload_map.asp?meet_id=<%=lThisMeet%>',800,300)">Upload Course Map</a>
					&nbsp;|&nbsp;
					<a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/info_sheets/upload_info.asp?meet_id=<%=lThisMeet%>',800,300)">Upload Info Sheet</a>
					<%If Not sMeetInfoSheet = vbNullString Then%>
						&nbsp;|&nbsp;
						<a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/info_sheets/<%=sMeetInfoSheet%>',1024,768)">Info Sheet</a>
					<%End If%>
					<%If Not sCourseMap = vbNullString Then%>
						&nbsp;|&nbsp;
						<a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/course_maps/<%=sCourseMap%>',1024,768)">Course Map</a>
					<%End If%>
				</div>
				
				<form name="meet_info" method="post" action="manage_meet.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkFlds()">
				<table>
					<tr>	
						<th><span style="color:#d62002">*</span>Meet Name:</th>
						<td><input name="meet_name" id="meet_name" maxlength="30" size="45" value="<%=MeetArray(0)%>"></td>
						<th><span style="color:#d62002">*</span>Meet Date:</th>
						<td><input name="meet_date" id="meet_date" maxLength="20" size="20" value="<%=MeetArray(1)%>"></td>
					</tr>
					<tr>	
						<th><span style="color:#d62002">*</span>Meet Host:</th>
						<td><input type="text" name="meet_host" id="meet_host" size="45" value="<%=MeetArray(2)%>"></td>
						<th>Web Site:</th>
						<td><input type="text" name="website" id="website" size="45" value="<%=MeetArray(3)%>"></td>
					</tr>
					<tr>	
						<th valign="top" rowspan="2"><span style="color:#d62002">*</span>Meet Site:</th>
						<td rowspan="2"><textarea name="meet_site" id="meet_site" rows="2" cols="32"><%=MeetArray(4)%></textarea></td>
						<th valign="top">Map Link:</th>
						<td valign="top"><input type="text" name="map_link" id="map_link" size="45" value="<%=MeetArray(5)%>"></td>
					</tr>
					<tr>
						<th valign="top">Rslts Official:</th>
						<td valign="top">
							<select name="official_rslts" id="official_rslts">
								<%If MeetArray(14) = "n" Then%> 
									<option value="n" selected>n</option>
									<option value="y">y</option>
								<%Else%>
									<option value="n">n</option>
									<option value="y" selected>y</option>
								<%End If%>
							</select>
						</td>
					</tr>
					<tr>	
						<th><span style="color:#d62002">*</span>Sport:</th>
						<td>
							<select name="sport" id="sport">
								<%If MeetArray(8) = "Cross-Country" Then%> 
									<option value="Cross-Country" selected>Cross-Country</option>
									<option value="Nordic Ski">Nordic Ski</option>
								<%Else%>
									<option value="Cross-Country">Cross-Country</option>
									<option value="Nordic Ski" selected>Nordic Ski</option>
								<%End If%>
							</select>
						</td>
						<th>Show Online:</th>
						<td>
							<select name="show_online" id="show_online">
								<%If MeetArray(9) = "n" Then%> 
									<option value="n" selected>n</option>
									<option value="y">y</option>
								<%Else%>
									<option value="n">n</option>
									<option value="y" selected>y</option>
								<%End If%>
							</select>
						</td>
					</tr>
					<tr>	
						<th><span style="color:#d62002">*</span>Shutdown:</th>
						<td><input type="text" name="when_shutdown" id="when_shutdown" size="45" value="<%=MeetArray(10)%>"></td>
						<th>Bib Range:</th>
						<td>
							From:&nbsp;<input type="text" name="bib_start" id="bib_start" size="4" value="<%=MeetArray(11)%>">
							To:&nbsp;<input type="text" name="bib_end" id="bib_end" size="4" value="<%=MeetArray(12)%>">
						</td>
					</tr>
					<tr>
						<th valign="top">Weather:</th>
						<td><textarea name="weather" id="weather" rows="2" cols="32"><%=MeetArray(6)%></textarea></td>
						<th valign="top">Comments:</th>
						<td><textarea name="comments" id="comments" rows="2" cols="32"><%=MeetArray(7)%></textarea></td>
					</tr>
					<tr>	
						<th valign="top"><span style="color:#d62002">*</span>Meet Director:</th>
						<td valign="top">
							<select name="meet_dir_id" id="meet_dir_id">
								<%For i = 0 To UBound(MeetDir, 2) - 1%>
									<%If CLng(MeetArray(13)) = CLng(MeetDir(0, i)) Then%> 
										<option value="<%=MeetDir(0, i)%>" selected><%=MeetDir(1, i)%></option>
									<%Else%>
										<option value="<%=MeetDir(0, i)%>"><%=MeetDir(1, i)%></option>
									<%End If%>
								<%Next%>
							</select>
						</td>
						<td colspan="2" rowspan="2">
							<ul style="list-style:none;">
								<li style="font-weight:bold;"><%=sName%></li>
								<li><%=sPhone%></li>
								<li><a href="mailto:<%=sEmail%>"><%=sEmail%></a></li>
							</ul>
						</td>
					</tr>
                    <tr>
						<th>Dynamic Race Assign:</th>
						<td>
							<select name="dynamic_race_assign" id="dynamic_race_assign">
								<%If MeetArray(16) = "y" Then%> 
									<option value="y" selected>y</option>
									<option value="n">n</option>
								<%Else%>
									<option value="y">y</option>
									<option value="n" selected>n</option>
								<%End If%>
							</select>
						</td>
					</tr>
  					<tr>
						<th valign="top">Invoice:</th>
						<td>
							$<input type="text" name="invoice" id="invoice" size="5" value="<%=MeetArray(15)%>">
						</td>
						<th valign="top">Start Time:</th>
                        <td>
                            <%=sStartTime%>
                        </td>
                    </tr>
					<tr>
						<td style="text-align:right;" colspan="2">
							<input type="checkbox" name="delete_meet" id="delete_meet">&nbsp;Delete This Meet
						</td>
					</tr>
					<tr>
						<td colspan="4">
							<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
							<input type="submit" name="submit2" id="submit2" value="Save Changes">
						</td>
					</tr>
				</table>
				</form>
			<%End If%>
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
