<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lThisMeet
Dim iGrade
Dim sMeetName, sName, sTeam, sBib, sMF, sRace, sGradeYear
Dim dMeetDate
Dim BibArray()
Dim fs, fname, sFileName

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

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
 
'get year for roster grades
If Month(dMeetDate) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If
	
If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

i = 0
ReDim BibArray(5, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.FirstName, r.LastName, t.TeamsID, t.Gender, r.Gender, g.Grade" & sGradeYear & ", ir.RacesID, ir.Bib FROM Roster r INNER JOIN Grades g "
sql = sql & "ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID WHERE ir.MeetsID = " 
sql = sql & lThisMeet & " AND ir.Bib <> 0 ORDER BY t.TeamName, t.Gender, ir.Bib"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	BibArray(0, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(0).Value, "''", "'") 
	BibArray(1, i) = TeamName(rs(2).Value) & " (" & rs(3).Value & ")"
	BibArray(2, i) = rs(4).Value
	BibArray(3, i) = rs(5).Value 
	BibArray(4, i) = RaceName(rs(6).Value)
	BibArray(5, i) = rs(7).Value
	i = i + 1
	ReDim Preserve BibArray(5, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function TeamName(lTeamID)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
	rs2.Open sql2, conn, 1, 2
	TeamName = Replace(rs2(0).Value, "''", "'")
	rs2.Close
	Set rs2 = Nothing
End Function

Private Function RaceName(lRaceID)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RaceName FROM Races WHERE RacesID = " & lRaceID
	rs2.Open sql2, conn, 1, 2
	RaceName = Replace(rs2(0).Value, "''", "'")
	rs2.Close
	Set rs2 = Nothing
End Function

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\ccmeet_admin\manage_meet\downloads\bib_list_" & sMeetName & "_" & Year(CDate(dMeetDate)) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine(UCase("Bib List for " & sMeetName & "-" & Year(CDate(dMeetDate))))
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("NAME" & Space(20) & vbTab & "TEAM" & Space(20) & vbTab & "M/F" & vbTab & "GR" & vbTab & "RACE" & vbTab & "BIB ")
For i = 0 to UBound(BibArray, 2) - 1
	sName = BibArray(0, i)
	If Len(sName) < 24 Then
		sName = sName & Space(24 - Len(sName))
	Else
		sName = Left(sName, 24)
	End If
		
	sTeam = BibArray(1, i)
	If Len(sTeam) < 24 Then
		sTeam = sTeam & Space(24 - Len(sTeam))
	Else
		sTeam = Left(sTeam, 24)
	End If
		
	sMF = BibArray(2, i)
	iGrade = BibArray(3, i)
	sRace = BibArray(4, i)
	sBib = BibArray(5, i)
		
	fname.WriteLine(sName & vbTab & sTeam & vbTab & sMF & vbTab & iGrade & vbTab & sRace & vbTab & sBib)
Next

'begin download
Response.Redirect "downloads/bib_list_" & sMeetName & "_" & Year(CDate(dMeetDate)) & ".txt"

fname.Close
Set fname=nothing
Set fs=nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title><%=sMeetName%> Bib List</title>
<!--#include file = "../../includes/meta2.asp" -->

<script type="text/javascript" src="../../misc/scripts.js"></script>
<link rel="stylesheet" type="text/css" href="../../misc/styles.css">

</head>
<body style="background-image:none">
	&nbsp;
</body>
<%
conn.Close
Set conn=Nothing
%>
</html>
