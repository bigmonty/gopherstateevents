<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisMeet
Dim sMeetName
Dim dMeetDate
Dim TeamArray()
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

i = 0
ReDim TeamArray(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT t.TeamName, c.FirstName, c.LastName, t.Gender FROM Teams t  INNER JOIN Coaches c ON t.CoachesID = c.CoachesID "
sql = sql & "INNER JOIN MeetTeams mt ON mt.TeamsID = t.TeamsID WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	TeamArray(0, i) = rs(0).Value & " (" & rs(3).Value & ")"
    TeamArray(1, i) = rs(2).Value & ",  " & rs(1).Value
	i = i + 1
	ReDim Preserve TeamArray(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\ccmeet_admin\manage_meet\downloads\team_labels.txt"

Response.Write sFileName

Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("Team Labels for " & sMeetName & " on " & dMeetDate)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("Team" & vbTab & "Coach")
For i = 0 to UBound(TeamArray, 2) - 1
	fname.WriteLine(TeamArray(0, i) & vbTab & TeamArray(1, i))
Next

'begin download
Response.Redirect "downloads/team_labels.txt"

fname.Close
Set fname=nothing
Set fs=nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title><%=sMeetName%> Team Labels</title>
<!--#include file = "../../includes/meta2.asp" -->

</head>
<body style="background-image:none">
	&nbsp;
</body>
<%
conn.Close
Set conn=Nothing
%>
</html>
