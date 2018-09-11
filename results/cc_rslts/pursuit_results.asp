<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, rs, conn, rs2, sql2
Dim i, j, k, m, x
Dim lPrelimID
Dim sPursuitName, sPursuitDesc, sPrelimDesc
Dim sPursuitDist, sPrelimDist, dPrelimDate
Dim sPrelimMeet
Dim TmRsltsArr(), TempArr(), IndRsltsArr(), RankArr(), Races()
Dim iFileNum, fName, a, fs
Dim sPartName, sTeamName
Dim lMeetID, sMeetName, dMeetDate, sMeetSite, sWeather
Dim sGradeYear

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetName, MeetDate, MeetSite, Weather FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value
If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
If Not rs(3).Value & "" = "" Then sWeather = Replace(rs(3).Value, "''", "'")
Set rs = Nothing
	
i = 0
ReDim Races(0)
sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(i) = rs(0).Value
	i = i + 1
	ReDim Preserve Races(i)
	rs.MoveNext
Loop
Set rs = Nothing

'get year for roster grades
If Month(dMeetDate) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

Set fs=Server.CreateObject("Scripting.FileSystemObject")
fname = "C:\Inetpub\h51web\gopherstateevents\dwnlds\" & sMeetName & "_" & Year(dMeetDate) & "_pursuit.txt"
Set a=fs.CreateTextFile(fname, True)

'write overall
a.WriteBlankLines (2)
a.writeline ("Pursuit Results for " & sMeetName & " on " & dMeetDate)
a.WriteLine("Meet Site: " & sMeetSite)
If Not sWeather = vbNullString Then a.WriteLine("Weather: " & sWeather)
a.WriteLine("Generated: " & Now())
          
For x = 0 To UBound(Races) - 1
	sPursuitDist = vbNullString
	sPursuitDesc = vbNullString
	sPursuitName = vbNullString
	
	lPrelimID = 0
	
    sPrelimDist = vbNullString
    dPrelimDate = vbNullString
    sPrelimDesc = vbNullString
    sPrelimMeet = vbNullString
	
	'get pursuit data
	sql = "SELECT RaceDist, RaceUnits, RaceDesc, RaceName, Technique FROM Races WHERE RacesID = " & Races(x)
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
       '---
    Else
	    sPursuitDist = rs(0).Value & " " & rs(1).Value
	    sPursuitDesc = rs(4).Value
	    sPursuitName = rs(3).Value
    End If
	Set rs = Nothing
	      
	'get prelim race id
	sql = "SELECT PrelimRace FROM Pursuit WHERE RacesID = " & Races(x)
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
       lPrelimID = 0
    Else
	    lPrelimID = rs(0).Value
    End If
	Set rs = Nothing
	      
	'get prelim race info
	sql = "SELECT r.RaceDist, r.RaceUnits, m.MeetDate, r.RaceDesc, m.MeetName, r.Technique FROM Races r INNER JOIN Meets m "
	sql = sql & "ON r.MeetsID = m.MeetsID WHERE RacesID = " & lPrelimID
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
       '---
    Else
	    sPrelimDist = rs(0).Value & " " & rs(1).Value
	    dPrelimDate = rs(2).Value
	    sPrelimDesc = rs(5).Value
	    sPrelimMeet = rs(4).Value
    End If
	Set rs = Nothing

	a.WriteBlankLines (2)
	a.writeline (sPursuitDesc & " (" & sPursuitDist & ")")
	a.writeline ("(Prelim Race: " & sPrelimMeet & " " & sPrelimDesc & " on " & dPrelimDate & " (" & sPrelimDist & "))")
	a.WriteBlankLines (1)
	        
	'get team results
	a.writeline ("Team Results")
	a.WriteBlankLines (1)
	         
	a.writeline ("PL" & vbTab & "TEAM" & Space(14) & vbTab & "SCORE" & vbTab & "R1 " & vbTab & "R2 " & vbTab & "R3 " & vbTab & "R4 " & vbTab & "R5 " & vbTab & "R6 " & vbTab & "R7 ")
	a.writeline ("------------------------------------------------------------------------------------------")
	k = 0
	ReDim TmRsltsArr(9, 0)
	sql = "SELECT t.TeamsID, t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
	sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & Races(x)
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    For j = 0 To 9
	        TmRsltsArr(j, k) = Trim(rs(j).Value)
	    Next
	    k = k + 1
	    ReDim Preserve TmRsltsArr(9, k)
	    rs.MoveNext
	Loop
	Set rs = Nothing
	
	'sort the arrays
	For k = 0 To UBound(TmRsltsArr, 2) - 2
	    For j = k + 1 To UBound(TmRsltsArr, 2) - 1
	        If ConvertToSeconds(TmRsltsArr(2, k)) < ConvertToSeconds(TmRsltsArr(2, j)) Then
	            For m = 0 To 9
	                TempArr(m) = TmRsltsArr(m, k)
	                TmRsltsArr(m, k) = TmRsltsArr(m, j)
	                TmRsltsArr(m, j) = TempArr(m)
	            Next
	        End If
	    Next
	Next
	       
	'now write to the file
	For k = 0 To UBound(TmRsltsArr, 2) - 1
		sTeamName = TmRsltsArr(1, k)
		If Len(sTeamName) < 20 Then
			sTeamName = sTeamName & Space(20 - Len(sTeamName))
		Else
			sTeamName = Left(sTeamName, 20)
		End If
	            
	   a.writeline (k + 1 & vbTab & sTeamName & vbTab & TmRsltsArr(2, k) & vbTab & TmRsltsArr(3, k) & vbTab & TmRsltsArr(4, k) & vbTab & TmRsltsArr(5, k) & vbTab & TmRsltsArr(6, k) & vbTab & TmRsltsArr(7, k) & vbTab & TmRsltsArr(8, k) & vbTab & TmRsltsArr(9, k))
	Next
	        
	'get a finishers array for this race
	j = 0
	ReDim IndRsltsArr(8, 0)
	sql = "SELECT r.RosterID, r.FirstName, r.LastName, t.TeamName, ir.RaceTime, ir.Bib, g.Grade" & sGradeYear & " FROM IndRslts ir INNER JOIN Roster r"
	sql = sql & " ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
	sql = sql & "WHERE ir.RacesID = " & Races(x) & " AND ir.Place > 0 AND ir.RaceTime > '00:00' ORDER BY ir.Place"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    IndRsltsArr(0, j) = rs(0).Value                 'name
	    IndRsltsArr(1, j) = rs(5).Value & "-" & Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	    IndRsltsArr(2, j) = rs(6).Value                 'grade
	    IndRsltsArr(3, j) = rs(3).Value                 'team
	    IndRsltsArr(8, j) = rs(4).Value                 'pursuit time
	                
	    j = j + 1
	    ReDim Preserve IndRsltsArr(8, j)
	    rs.MoveNext
	Loop
	Set rs = Nothing
	        
	'create rank array
	ReDim RankArr(2, UBound(IndRsltsArr, 2))
	ReDim TempArr(1)
	For j = 0 To UBound(IndRsltsArr, 2) - 1
	    RankArr(0, j) = IndRsltsArr(0, j)
	    RankArr(1, j) = IndRsltsArr(8, j)
	Next
	        
	'sort the rank array
	For j = 0 To UBound(RankArr, 2) - 2
	    For k = j + 1 To UBound(RankArr, 2) - 1
	        If ConvertToSeconds(RankArr(1, j)) > ConvertToSeconds(RankArr(1, k)) Then
	            For m = 0 To 1
	                TempArr(m) = RankArr(m, j)
	                RankArr(m, j) = RankArr(m, k)
	                RankArr(m, k) = TempArr(m)
	            Next
	        End If
	    Next
	Next
	        
	For j = 0 To UBound(RankArr, 2) - 1
	    RankArr(2, j) = j + 1
	Next
	        
	'sort the rank array
	For j = 0 To UBound(IndRsltsArr, 2) - 1
	    For k = 0 To UBound(RankArr, 2) - 1
	        If CLng(IndRsltsArr(0, j)) = CLng(RankArr(0, k)) Then
	            IndRsltsArr(7, j) = RankArr(2, k)
	            Exit For
	        End If
	    Next
	Next
	        
	'now get prelim time and rank
	j = 1
	sql = "SELECT RosterID, RaceTime FROM IndRslts WHERE RacesID = " & lPrelimID & " AND RaceTime <> '00:00' AND Place <> 0 ORDER BY RaceTime"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    For k = 0 To UBound(IndRsltsArr, 2) - 1
	        If CLng(IndRsltsArr(0, k)) = CLng(rs(0).Value) Then
	            IndRsltsArr(4, k) = ConvertToMinutes(Round(ConvertToSeconds(rs(1).Value) + ConvertToSeconds(IndRsltsArr(8, k)), 2))
	            IndRsltsArr(5, k) = j
	            IndRsltsArr(6, k) = rs(1).Value
	            j = j + 1
	        End If
	    Next
	    rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	        
	'write to text file
	a.WriteBlankLines (2)
	a.writeline ("Individual Pursuit Results")
	a.WriteBlankLines (1)
	
	a.writeline (Space(72) & vbTab & Space(4) & sPrelimDesc & Space(6) & vbTab & Space(4) & sPursuitDesc)
	a.writeline (Space(72) & vbTab & "----------------" & vbTab & "---------------")
	a.writeline ("PL" & vbTab & "BIB-NAME" & Space(12) & vbTab & "GR" & vbTab & "TEAM" & Space(16) & _
	             vbTab & "COMB TIME  " & vbTab & "RANK" & vbTab & " TIME  " & vbTab & vbTab & _
	             "RANK" & vbTab & " TIME")
	a.writeline ("-----------------------------------------------------------------------------------------------------------------------")
	            
	For j = 0 To UBound(IndRsltsArr, 2) - 1
		sPartName = IndRsltsArr(1, j)
		If Len(sPartName) < 20 Then
			sPartName = sPartName & Space(20 - Len(sPartName))
		Else
			sPartName = Left(sPartName, 20)
		End If
		
		sTeamName = IndRsltsArr(3, j)
		If Len(sTeamName) < 20 Then
			sTeamName = sTeamName & Space(20 - Len(sTeamName))
		Else
			sTeamName = Left(sTeamName, 20)
		End If
	            
	    a.writeline (j + 1 & vbTab & sPartName & vbTab & IndRsltsArr(2, j) & vbTab & sTeamName & vbTab & IndRsltsArr(4, j) & vbTab & IndRsltsArr(5, j) & vbTab & IndRsltsArr(6, j) & vbTab & IndRsltsArr(7, j) & vbTab & IndRsltsArr(8, j))
	Next
Next
	        
a.Close
Set a = Nothing
Set fs = Nothing

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<%

'begin download
Response.Redirect "../../dwnlds/" & sMeetName & "_" & Year(dMeetDate) & "_pursuit.txt"
%>
<!DOCTYPE html>
<html>
<head>
<title>GSE Pursuit-Formatted Nordic Ski Results</title>
<meta name="description" content="GSE pursuit-formatted nordic ski results.">
</head>
<body>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
