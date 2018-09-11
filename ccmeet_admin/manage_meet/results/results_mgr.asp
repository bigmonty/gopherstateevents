<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k, m, n
Dim lThisMeet, lThisRace, lRacesID
Dim iNumEntrants, iTtlFin, iBibToFind, iFirstBib, iFirstPlace, iSecondBib, iSecondPlace, iInsertBib, iInsertPl
Dim BibRslts(5), Races(), RsltsArr(), AvailBibs(), AsgndBibs(), RawRslts(), RenumRslts, SortArr(1)
Dim sMeetName, sGradeYear, sOrderResultsBy, sRaceName, sScoreMethod, sInsert, sDelete, sExclude, sInsrtTime, sSetDelay, sRaceGender, sRenumber
Dim sRefreshTimes, sMinTime, sManualTime, sErrMsg, sFirstTime, sSecondTime, sRaceTime, sFinishTime, sMyStart, sRaceDelay
Dim sngFirstTime, sngSecondTime, sngFnlScnds
Dim dMeetDate
Dim bFound, bErrExists

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

iBibToFind = 0

lThisMeet = Request.QueryString("meet_id")
lThisRace = Request.QueryString("this_race")

sDelete = Request.QueryString("delete_fin")
If sDelete = vbNullString Then sDelete = "n"

sRenumber = Request.QueryString("renumber")
If sRenumber = vbNullString Then sRenumber = "n"

sRefreshTimes = Request.QueryString("refresh_times")
If sRefreshTimes = vbNullString Then sRefreshTimes = "n"

sExclude = Request.QueryString("exclude")
If sExclude = vbNullString Then sExclude = "n"

sInsert = Request.QueryString("insert")
If sInsert = vbNullString Then sInsert = "n"

sSetDelay = Request.QueryString("set_delay")
If sSetDelay = vbNullString Then sSetDelay = "n"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT IndRsltsID FROM IndRslts WHERE MeetsID = " & lThisMeet
rs.Open sql, conn, 1, 2
iNumEntrants = rs.RecordCount
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT IndRsltsID FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND FnlScnds > 0 AND Place > 0 AND RaceTime > '00:00'"
rs.Open sql, conn, 1, 2
iTtlFin = rs.RecordCount
rs.Close
Set rs = Nothing

'get order by
sql = "SELECT OrderBy FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
If rs(0).Value = "Place" Then
    sOrderResultsBy = "Place"
Else
    sOrderResultsBy = "Time"
End If
Set rs = Nothing

'get year for roster grades
If Month(dMeetDate) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

i = 0    
ReDim Races(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisMeet & " ORDER BY ViewOrder, RaceTime"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
    Races(1, i) = Replace(rs(1).Value, "''", "'")
    Races(2, i) = GetDelay(rs(0).Value)
	i = i + 1
	ReDim Preserve Races(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If CStr(lThisRace) = vbNullString Then lThisRace = Races(0, 0)

If Request.Form.Item("submit_swap") = "submit_swap" Then
    bErrExists = False
    
    If Request.Form.Item("bib_1") = vbNullString Then bErrExists = True
    If Request.Form.Item("bib_2") = vbNullString Then bErrExists = True
    
    If bErrExists = False Then
        If Not IsNumeric(Request.Form.Item("bib_1")) Then bErrExists = True
    End If
    
    If bErrExists = False Then
        If Not IsNumeric(Request.Form.Item("bib_2")) Then bErrExists = True
    End If
    
    If bErrExists = False Then
        iFirstBib = Request.Form.Item("bib_1")
        iSecondBib = Request.Form.Item("bib_2")
        
        iFirstPlace = 0
        iSecondPlace = 0
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Place, ElpsdTime FROM IndRslts WHERE Bib = " & iFirstBib & " AND MeetsID = " & lThisMeet
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            iFirstPlace = rs(0).Value
            sFirstTime = rs(1).Value
        Else
            bErrExists = True
        End If
        rs.Close
        Set rs = Nothing
        
        If bErrExists = False Then
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT Place, ElpsdTime FROM IndRslts WHERE Bib = " & iSecondBib & " AND MeetsID = " & lThisMeet
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then
                iSecondPlace = rs(0).Value
                sSecondTime = rs(1).Value
            Else
                bErrExists = True
            End If
            rs.Close
            Set rs = Nothing
        End If
    End If

    'note error
    If bErrExists = True Then
        sErrMsg = "Oops!  Make sure you have entered two numeric bib numbers to swap and they both exist in the results."
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Place, ElpsdTime, RaceTime, IndDelay, FnlScnds, RacesID FROM IndRslts WHERE Bib = " & iFirstBib
        sql = sql & " AND MeetsID = " & lThisMeet
        rs.Open sql, conn, 1, 2
        rs(0).Value = iSecondPlace
        rs(1).Value = sSecondTime
        sRaceTime = GetRaceTime(sSecondTime, rs(3).Value, rs(5).Value)  'elpsdtime - racedelay - inddelay
        rs(2).Value = sRaceTime
        rs(4).Value = ConvertToSeconds(sRaceTime)
        rs.Update
        rs.Close
        Set rs = Nothing
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Place, ElpsdTime, RaceTime, IndDelay, FnlScnds, RacesID FROM IndRslts WHERE Bib = " & iSecondBib
        sql = sql & " AND MeetsID = " & lThisMeet
        rs.Open sql, conn, 1, 2
        rs(0).Value = iFirstPlace
        rs(1).Value = sFirstTime
        sRaceTime = GetRaceTime(sFirstTime, rs(3).Value, rs(5).Value)  'elpsdtime - racedelay - inddelay
        rs(2).Value = sRaceTime
        rs(4).Value = ConvertToSeconds(sRaceTime)
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_race") = "submit_race" Then  
    lThisRace = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_restore") = "submit_restore" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Excludes FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND Bib = " & Request.Form.Item("restores")
    rs.Open sql, conn, 1, 2
    rs(0).Value = "y"
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_exclude") = "submit_exclude" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Excludes FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND Bib = " & Request.Form.Item("exclude_bib")
    rs.Open sql, conn, 1, 2
    rs(0).Value = "n"
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_delay") = "submit_delay" Then
    Set rs = Server.CreateOBject("ADODB.Recordset")
    sql = "SELECT RaceDelay FROM RaceDelay WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("delay")
    rs.Update
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateOBject("ADODB.Recordset")
    sql = "SELECT MinTime FROM Races WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("min_time")
    rs.Update
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateOBject("ADODB.Recordset")
    sql = "SELECT FirstBib, ManualTime FROM RaceWinners WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("first_bib")
    rs(1).Value = Request.Form.Item("manual_time")
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_delete") = "submit_delete" Then
    If Not CStr(Request.Form.Item("delete_bib")) = vbNullString Then
        Set rs = Server.CreateOBject("ADODB.Recordset")
        sql = "SELECT ElpsdTime, RaceTime, Place, FnlScnds FROM IndRslts WHERE RacesID = " & lThisRace & " AND Bib = " & Request.Form.Item("delete_bib")
        rs.Open sql, conn, 1, 2
        rs(0).Value = "00:00"
        rs(1).Value = "00:00"
        rs(2).Value = "0"
        rs(3).Value = "0"
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_insert") = "submit_insert" Then
    iInsertBib = Request.Form.Item("insert_bib")
    iInsertPl = Request.Form.Item("insert_pl")
    sFinishTime = Request.Form.Item("insert_time") 
    sMyStart = Request.Form.Item("insert_start")

    'get race delay
    sql = "SELECT RaceDelay FROM RaceDelay WHERE RacesID = " & lThisRace
    Set rs = conn.execute(sql)
    sRaceDelay = rs(0).Value
    Set rs = Nothing

    'determine final time
     sngFnlScnds = ConvertToSeconds(sFinishTime) - ConvertToSeconds(sRaceDelay) - ConvertToSeconds(sMyStart)
     sngFnlScnds = Round(sngFnlScnds, 3)

     sRaceTime = ConvertToMinutes(sngFnlScnds)

    'get all finishers in this event and add one to all event places beyond that
    Set rs = Server.CreateOBject("ADODB.Recordset")
    sql = "SELECT Place FROM IndRslts WHERE RacesID = " & lThisRace & " AND Place >= " & iInsertPl & " ORDER BY FnlScnds"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        rs(0).Value = CInt(rs(0).Value) + 1
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'enter this place
    Set rs = Server.CreateOBject("ADODB.Recordset")
    sql = "SELECT Place FROM IndRslts WHERE Bib = " & iInsertBib & " AND RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    If rs.REcordCount > 0 Then
        rs(0).Value = iInsertPl
        rs.Update
    Else
        sErrMsg = "Oops.  That bib must not be entered in this race."
    End If
    rs.Close
    Set rs = Nothing

    If sErrMsg = vbNullString Then
        'enter this time
        Set rs = Server.CreateOBject("ADODB.Recordset")
        sql = "SELECT ElpsdTime, RaceTime, IndDelay, FnlScnds FROM IndRslts WHERE RacesID = " & lThisRace & " AND Bib = " & iInsertBib
        rs.Open sql, conn, 1, 2
        rs(0).Value = sFinishTime
        rs(1).Value = sRaceTime
        rs(2).Value = sMyStart
        rs(3).Value = sngFnlScnds
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.form.Item("submit_bib") = "submit_bib" Then
    iBibToFind = Request.Form.Item("bib_to_find")
End If

If Not CInt(iBibToFind) = 0 Then
    sql = "SELECT r.RosterID, r.FirstName, r.LastName, t.TeamName, r.Gender,ir.RaceTime, g.Grade" & sGradeYear & " FROM IndRslts ir "
	sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
	sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lThisRace & " AND ir.Bib = " & iBibToFind
    Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
        sErrMsg = "Oops.  That bib number was not found in thre results for this race.  Please check another race in this meet."
    Else
        BibRslts(0) = GetPlace(rs(0).Value)
        BibRslts(1) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        BibRslts(2) = rs(3).Value
        BibRslts(3) = rs(4).Value
        BibRslts(4) = rs(6).Value
        BibRslts(5) = rs(5).Value
    End If
    Set rs = Nothing
End If

If sRenumber = "y" Then
    i = 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Place FROM IndRslts WHERE RacesID = " & lThisRace & " AND Place > 0 AND ElpsdTime > '00:00' AND FnlScnds > 0 "
    sql = sql & "ORDER BY FnlScnds"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        rs(0).Value = i
        i = i + 1
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

If sRefreshTimes = "y" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.ElpsdTime, rd.RaceDelay, ir.RaceTime, ir.IndDelay, ir.FnlScnds FROM IndRslts ir INNER JOIN RaceDelay rd "
    sql = sql & "ON rd.RacesID = ir.RacesID WHERE ir.MeetsID = " & lThisMeet & " ORDER BY ir.Place"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngFnlScnds = ConvertToSeconds(rs(0).Value) - ConvertToSeconds(rs(1).Value) - CSng(rs(3).Value)
        rs(2).Value = ConvertToMinutes(sngFnlScnds)
        rs(4).Value = sngFnlScnds
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

sql = "SELECT RaceName, ScoreMethod, Gender, MinTime FROM Races WHERE RacesID = " & lThisRace
Set rs = conn.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'")
sScoreMethod = rs(1).Value
sRaceGender = rs(2).Value
sMinTime = rs(3).Value
Set rs = Nothing

'get race delay
sql = "SELECT RaceDelay FROM RaceDelay WHERE RacesID = " & lThisRace
Set rs = conn.execute(sql)
sRaceDelay = rs(0).Value
Set rs = Nothing

sql = "SELECT FirstBib, ManualTime FROM RaceWinners WHERE RacesID = " & lThisRace
Set rs = conn.Execute(sql)
If rs.BOF And rs.EOF Then
    bFound = False
Else
    iFirstBib = rs(0).Value
    sManualTime = rs(1).Value
    bFound = True
End If
Set rs = Nothing
   
If bFound = False Then
    sql = "INSERT INTO RaceWinners(RacesID) VALUES (" & lThisRace & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
        
    iFirstBib = "0"
    sManualTime = "00:00"
End If

i = 0
ReDim RsltsArr(12, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.FirstName, r.LastName, t.TeamName, r.RosterID, r.Gender, ir.RaceTime, ra.RaceDist, "
sql = sql & "ra.RaceUnits, ir.Excludes, ir.TeamPlace, ir.Bib, ir.ElpsdTime, ir.Place FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID "
sql = sql & "INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID INNER JOIN Races ra "
sql = sql & "ON ra.RacesID = ir.RacesID WHERE ir.RacesID = " & lThisRace & " AND ir.Place > 0 AND ir.ElpsdTime > '00:00' AND ir.FnlScnds > 0 "
sql = sql & "ORDER BY ir.FnlScnds, ir.Place"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	RsltsArr(0,i) = rs(10).Value & "-" & Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
	RsltsArr(1,i) = Replace(rs(2).Value, "''", "'")
	RsltsArr(2,i) = GetGrade(rs(3).Value)
	RsltsArr(3,i) = rs(4).Value
	RsltsArr(4,i) = rs(5).Value
	RsltsArr(5,i) = rs(6).Value
	RsltsArr(6,i) = rs(7).Value
	RsltsArr(7,i) = rs(8).Value
	If CInt(rs(9).Value) = 0 Then
		RsltsArr(8,i) = "---"
	Else
		RsltsArr(8,i) = rs(9).Value
	End If
    RsltsArr(9, i) = rs(10).Value
    RsltsArr(10, i) = rs(3).Value
    RsltsArr(11, i) = rs(11).Value
    RsltsArr(12, i) = rs(12).Value
	i = i + 1
	ReDim Preserve RsltsArr(12, i)

	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

i = 0
ReDim AsgndBibs(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND Place <> 0 AND Bib <> 0 ORDER BY Bib"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	AsgndBibs(i) = rs(0).Value
	i = i + 1
	ReDim Preserve AsgndBibs(i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

k = 0  
Dim bBibFound    
ReDim AvailBibs(0)
For i = 1 To 1500
    bBibFound = False

    For j = 0 To UBound(AsgndBibs) - 1
        If CInt(i) = CInt(AsgndBibs(j)) Then
            bBibFound = True
            Exit For
        End If
    Next

    If bBibFound = False Then
        AvailBibs(k) = i
        k = k + 1
        ReDim Preserve AvailBibs(k)
    End If
Next

Function GetPlace(lRosterID)
	GetPlace = 0
	If sOrderResultsBy = "Time" Then
        sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRace & " AND Place > 0 ORDER BY FnlScnds"
    Else
        sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRace & " AND Place > 0 ORDER BY Place"
    End If
	Set rs2 = conn.Execute(sql2)
	Do While Not rs2.EOF
		GetPlace = GetPlace + 1
		If CLng(rs2(0).Value) = CLng(lRosterID) Then Exit Do
		rs2.MoveNext
	Loop
	Set rs2 = Nothing
End Function

%>
<!--#include file = "../../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../../includes/convert_to_minutes.asp" -->
<%	

Private Function GetGrade(lMyID)
    GetGrade = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & Right(CStr(Year(Date)), 2) & " FROM Grades WHERE RosterID = " & lMyID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetGrade = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function

Private Function GetDelay(lThisRace)
     Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT RaceDelay FROM RaceDelay WHERE RacesID = " & lThisRace
    rs2.Open sql2, conn, 1, 2
    GetDelay = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetRaceTime(sElpsdTime, sngIndDelay, lRaceID)
    Dim sngElpsdTime
    
    sngElpsdTime = ConvertToSeconds(sElpsdTime)
    
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT RaceDelay FROM RaceDelay WHERE RacesID = " & lRaceID
    rs2.Open sql2, conn, 1, 2
    If ConvertToSeconds(rs2(0).Value) = 0 And sngIndDelay = 0 Then
        GetRaceTime = sElpsdTime
    Else
        GetRaceTime = ConvertToMinutes(ConvertToSeconds(sElpsdTime) - ConvertToSeconds(rs2(0).Value) - sngIndDelay)
    End If
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE CC/Nordic Results Manager</title>

<script>
function chkInsert(){
 	if (document.insrt_fnshr.insert_pl.value == '' || 
 	    document.insrt_fnshr.insert_bib.value == '' ||
 	    document.insrt_fnshr.insert_race.value == '')
		{
  		alert('Place, bib , and race are required.');
  		return false
  		}
	else
   		return true
}

function chkFlds2() {
 	if (document.find_bib.bib_to_find.value == '')
		{
  		alert('You must submit a bib number to look for.');
  		return false
  		}
 	else
		if (isNaN(document.find_bib.bib_to_find.value))
    		{
			alert('The bib number can not contain non-numeric values');
			return false
			}
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "../manage_meet_nav.asp" -->
			<%End If%>

			<h4 class="h4">Manage Individual Results for <%=sMeetName%> on <%=dMeetDate%>:&nbsp;<%=sRaceName%></h4>

            <nav class="navbar navbar-expand-sm bg-light">
                <ul class="navbar-nav">
                    <li class="nav-item">
                        <a class="nav-link" href="#">Total Entries: <%=iNumEntrants%></a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="#">Total Finishers: <%=iTtlFin%></a>
                    </li>
                </ul>
            </nav>

			<form class="form-inline bg-success" name="get_races" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>" 
                style="padding: 5px 0 5px 15px;margin-bottom: 0;">
			<label for="races">Select Race:</label>
			<select class="form-control" name="races" id="races" onchange="this.form.get_race.click();">
				<%For i = 0 to UBound(Races, 2) - 1%>
					<%If CLng(lThisRace) = CLng(Races(0, i)) Then%>
						<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
					<%Else%>
						<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
			<input type="submit" class="form-control" name="get_race" id="get_race" value="Get Results">
			</form>

            <nav class="navbar navbar-expand-sm bg-light">
                <ul class="navbar-nav">
                    <%If sScoreMethod="Pursuit" Then%>
                        <li class="nav-item">
                            <a class="nav-link" href="pursuit_results.asp?meet_id=<%=lThisMeet%>">Pursuit-Formatted Results</a>
                        </li>
                    <%End If%>
                    <li class="nav-item">
                        <a class="nav-link" href="javascript:pop('/results/cc_rslts/print_rslts.asp?meet_id=<%=lThisMeet%>&amp;race_id=<%=lThisRace%>',1024,768)">Print These</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="/results/cc_rslts/comp_rslts.asp?meet_id=<%=lThisMeet%>" onclick="openThis(this.href,1024,768);return false;">Comprehensive Results</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="/results/cc_rslts/dwnld_overall.asp?meet_id=<%=lThisMeet%>&amp;race_id=<%=lThisRace%>" 
						onclick="openThis(this.href,800,600);return false;">Dwnld This</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="/results/cc_rslts/dual_rslts.asp?meet_id=<%=lThisMeet%>&amp;race_id=<%=lThisRace%>" 
						onclick="openThis(this.href,800,600);return false;">Dual Meet-Formated Results</a>
                    </li>
                </ul>
            </nav>

			<!--#include file = "results_nav.asp" -->				
			
            <div class="row">
                <div class="col-sm-3 bg-warning">
                    <h4 class="h4">Insert Finisher</h4>
				    <form class="form-horizontal" name="insrt_fnshr" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>" 
                        onsubmit="return chkInsert();">
                    <div class="form-group">
                        <label for="insert_pl" class="control-label col-xs-6">Overall Pl:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="insert_pl" id="insert_pl">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="insert_bib" class="control-label col-xs-6">Bib:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="insert_bib" id="insert_bib">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="insert_time" class="control-label col-xs-6">Elpsd Time:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="insert_time" id="insert_time">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="insert_start" class="control-label col-xs-6">Ind. Start:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="insert_start" id="insert_start" value="0">
                        </div>
                    </div>
                    <div class="form-group">
                        <input type="hidden" name="submit_insert" id="submit_insert" value="submit_insert">
				        <input type="submit" class="form-control" name="submit2" id="submit2" value="Submit">
                    </div>
                    </form>
                </div>
                <div class="col-sm-3 bg-success">
                    <h4 class="h4">Race Data</h4>

				    <form class="form-horizontal" name="r_delay" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>&amp;set_delay=y">
                    <div class="form-group">
                        <label for="delay" class="control-label col-xs-6">Race Delay:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="delay" id="delay" value="<%=sRaceDelay%>">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="first_bib" class="control-label col-xs-6">First Bib:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="first_bib" id="first_bib" value="<%=iFirstBib%>">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="manual_time" class="control-label col-xs-6">First Time:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="manual_time" id="manual_time" value="<%=sManualTime%>">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="min_time" class="control-label col-xs-6">Min Time:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="min_time" id="min_time" value="<%=sMinTime%>">
                        </div>
                    </div>
                    <div class="form-group">
                        <input type="hidden" name="submit_delay" id="submit_delay" value="submit_delay">
				        <input type="submit" class="form-control" name="submit6" id="submit6" value="Submit">
                    </div>
                    </form>
                </div>
                <div class="col-sm-3 bg-danger">
                    <h4 class="h4">Delete Finisher</h4>
				    <form class="form-horizontal" name="delete_fnshr" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>&amp;delete_fin=y">
                    <div class="form-group">
                        <label for="delete_bib" class="control-label col-xs-6">Bib:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="delete_bib" id="delete_bib">
                        </div>
                    </div>
                    <input type="hidden" name="submit_delete" id="submit_delete" value="submit_delete">
				    <input type="submit" class="form-control" name="submit3" id="submit3" value="Submit">
                    </form>
                    <br><br>
                    <h4 class="h4">Exclude Finisher</h4>
				    <form class="form-horizontal" name="delete_fnshr" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>&amp;exclude=y">
                    <div class="form-group">
                        <label for="exclude_bib" class="control-label col-xs-6">Bib:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="exclude_bib" id="exclude_bib">
                        </div>
                    </div>
                    <input type="hidden" name="submit_exclude" id="submit_exclude" value="submit_exclude">
				    <input type="submit" class="form-control" name="submit4" id="submit4" value="Submit">
                    </form>
                </div>
                <div class="col-sm-3 bg-warning">
                    <h4 class="h4" style="margin-bottom: 0;">Swap Finishers:</h4>
                    <form class="form-horizontal" name="swap_finishers" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>">
                    <div class="form-group">
                        <label for="bib_1" class="control-label col-xs-6">Bib 1:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="bib_1" id="bib_1">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="bib_2" class="control-label col-xs-6">Bib 2:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="bib_2" id="bib_2">
                        </div>
                    </div>
 				    <input type="hidden" name="submit_swap" id="submit_swap" value="submit_swap">
				    <input type="submit" class="form-control" name="submit2a" id="submit2a" value="Submit">
                    </form>     
                               
                    <h4 class="h4">Restore Finisher:</h4>
                    <form class="form-horizontal" name="restore_bib" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>">
                    <div class="form-group">
                        <label for="restores" class="control-label col-xs-6">Bib:</label>
                        <div class="col-xs-6">
                            <input type="text" class="form-control" name="restores" id="restores">
                        </div>
                    </div>
				    <input type="hidden" name="submit_restore" id="submit_restore" value="submit_restore">
				    <input type="submit" class="form-control" name="submit2" id="submit2" value="Submit">
                    </form>
                </div>
            </div>
			
			<h4 class="h4">Individual Results</h4>

            <div class="col-sm-5 bg-info">
                <form class="form-inline" name="find_bib" method="post" action="results_mgr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>" 
                    onsubmit="return chkFlds2();">
                <label for="bib_to_find">Bib To Find:</label>
                <input type="text" class="form-control" name="bib_to_find" id="bib_to_find" size="3" value="<%=iBibToFind%>">
                <input type="hidden" name="submit_bib" id="submit_bib" value="submit_bib">
                <input type="submit" class="form-control" name="submit_lookup" id="submit_lookup" value="Find Bib">
                </form>
            </div>
            <div class="col-sm-7 bg-info">
                <%If Not CInt(iBibToFind) = 0 Then%>
                    <%If sErrMsg = vbNullString Then%>
                        <table class="table">
                            <tr><th>Pl</th><th>Name</th><th>School</th><th>MF</th><th>Gr</th><th>Time</th></tr>
                            <tr>
                                <%For i = 0 To 5%>
                                    <td><%=BibRslts(i)%></td>
                                <%Next%>
                            </tr>
                        </table>
                    <%Else%>
                        <p><%=sErrMsg%></p>
                    <%End If%>
                <%End If%>
            </div>

			<table class="table table-striped">
				<tr>
					<th>Pl</th>
					<th>Tm</th>
					<th>Bib-Name</th>
					<th>Team</th>
					<th>Gr</th>
					<th>M/F</th>
					<th>Time</th>
					<th>Elpsd</th>
					<th>EvntPl</th>
				</tr>
				<%k = 1%>
				<%For i = 0 to UBound(RsltsArr, 2) - 1%>
					<tr>
						<td>
							<%If RsltsArr(7, i) = "y" Then%>
								-
							<%Else%>
								<%=k%>
								<%k = k + 1%>
							<%End If%>
						</td>
						<td>
							<%=RsltsArr(8, i)%>
						</td>
						<td>
                            <a href="javascript:pop('edit_indiv.asp?meet_id=<%=lThisMeet%>&amp;this_part=<%=RsltsArr(10, i)%>',800,300)"><%=RsltsArr(0, i)%></a>
                        </td>
						<td><%=RsltsArr(1, i)%></td>
						<td><%=RsltsArr(2, i)%></td>
						<td><%=RsltsArr(3, i)%></td>
						<td><%=RsltsArr(4, i)%></td>
						<td><%=RsltsArr(11, i)%></td>
						<td><%=RsltsArr(12, i)%></td>
					</tr>
				<%Next%>
			</table>
		</div>
    </div>
      	<!--#include file = "../../../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
