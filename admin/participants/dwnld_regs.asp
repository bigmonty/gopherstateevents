<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID
Dim sEventName, sEventRaces, sRace
Dim PartArray
Dim fs, fname, sFileName
Dim dEventDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = rs(0).Value
dEventDate = rs(1).Value
Set rs = Nothing

If CLng(lRaceID) = 0 Then
    sRace = "All Races"

    sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        sEventRaces = sEventRaces & rs(0).Value & ", "
	    rs.MoveNext
    Loop
    Set rs = Nothing

    If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

	sql="SELECT p.ParticipantID, p.FirstName, p.LastName, p.Gender, rc.Age, p.DOB, p.Phone, p.City, p.St, p.Email, rg.ShrtSize, rc.Bib, rc.RaceID FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces & ") AND rg.RaceID IN (" 
	sql = sql & sEventRaces & ") ORDER BY p.LastName, p.FirstName"
	Set rs=conn.Execute(sql)
	PartArray = rs.GetRows()
	Set rs=Nothing

    For i  = 0 To UBound(PartArray, 2)
        PartArray(12, i) = GetRaceName(PartArray(12, i))
    Next
Else
    sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
    Set rs = conn.Execute(sql)
    sRace = Replace(rs(0).Value, "''", "'")
    Set rs = Nothing

	sql="SELECT p.ParticipantID, p.FirstName, p.LastName, p.Gender, rc.Age, p.DOB, p.Phone, p.City, p.St, p.Email, rg.ShrtSize, rc.Bib, rc.RaceID FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE (rc.RaceID = " & lRaceID & " AND rg.RaceID = " 
	sql = sql & lRaceID & ") ORDER BY p.LastName, p.FirstName"
	Set rs=conn.Execute(sql)
	PartArray = rs.GetRows()
	Set rs=Nothing

    For i  = 0 To UBound(PartArray, 2)
        PartArray(12, i) = sRace
    Next
End If

Private Function GetRaceName(lThisRace)
    sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lThisRace
    Set rs = conn.Execute(sql)
    GetRaceName = Replace(rs(0).Value, "''", "'")
    Set rs = Nothing
End Function

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\admin\participants\downloads\registrants_" & sEventName & "_" & sRace & "_" & Year(CDate(dEventDate)) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine(UCase("Participant Registrations for " & sEventName & " " & sRace & "-" & Year(CDate(dEventDate))))
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("FIRST" & vbTab & "LAST" & vbTab & "M/F" & vbTab & "AGE" & vbTab & "DOB" & vbTab & "PHONE" & vbTab & "CITY" & vbTab & "ST" & vbTab & "EMAIL" & vbTab & "SIZE" & vbTAb & "BIB" & vbTab & "RACE")
For i = 0 to UBound(PartArray, 2)
	fname.WriteLine(PartArray(1, i) & vbTab & PartArray(2, i) & vbTab & PartArray(3, i) & vbTab & PartArray(4, i) & vbTab & PartArray(5, i) & vbTab & PartArray(6, i)  & vbTab & PartArray(7, i) & vbTab & PartArray(8, i) & vbTab & PartArray(9, i) & vbTab & PartArray(10, i) & vbTab & PartArray(11, i) & vbTab & PartArray(12, i))
Next

'begin download
Response.Redirect "downloads/registrants_" & sEventName & "_" & sRace & "_" & Year(CDate(dEventDate)) & ".txt"

fname.Close
Set fname=nothing
Set fs=nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title><%=sEventName%> Participants</title>
<!--#include file = "../../includes/meta2.asp" -->

<script type="text/javascript" src="../misc/scripts.js"></script>
<link rel="stylesheet" type="text/css" href="../misc/styles.css">

</head>
<body style="background-image:none">
	&nbsp;
</body>
<%
conn.Close
Set conn=Nothing
%>
</html>
