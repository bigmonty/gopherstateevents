<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet
Dim sMeetName
Dim dMeetDate
Dim TeamArray(), MClasses(), fClasses()
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

'get meet classes id
i = 0
ReDim MClasses(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetClassesID, ClassName FROM MeetClasses WHERE MeetsID = " & lThisMeet & " AND (Gender = 'M' or Gender = 'Both') ORDER BY ClassName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    MClasses(0, i) = rs(0).Value
    MClasses(1, i) = UCASE(rs(1).Value)
    i = i + 1
    ReDim Preserve MClasses(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

i = 0
ReDim FClasses(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetClassesID, ClassName FROM MeetClasses WHERE MeetsID = " & lThisMeet & " AND (Gender = 'F' or Gender = 'Both') ORDER BY ClassName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    FClasses(0, i) = rs(0).Value
    FClasses(1, i) = UCASE(rs(1).Value)
    i = i + 1
    ReDim Preserve FClasses(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetTeams(sGender, lThisClass)
    Dim x

    x = 0
    ReDim TeamArray(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT t.TeamName FROM Teams t  INNER JOIN Coaches c ON t.CoachesID = c.CoachesID INNER JOIN MeetTeams mt ON mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND t.Gender = '" & sGender & "' AND mt.MeetClass = '" & lThisClass & "' ORDER BY t.TeamName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    TeamArray(x) = Replace(rs(0).Value, "''", "'")
	    x = x + 1
	    ReDim Preserve TeamArray(x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\ccmeet_admin\manage_meet\downloads\teams_by_class.txt"

Response.Write sFileName

Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("Teams By Class for " & sMeetName & " on " & dMeetDate)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(2)

fname.WriteLine("MALE TEAMS")
fname.WriteBlankLines(2)

For i = 0 to UBound(MClasses, 2) - 1
	fname.WriteLine(MClasses(1, i))

    Call GetTeams("M", MClasses(0, i))

    For j = 0 To UBound(TeamArray) - 1
        fname.WriteLine(j + 1 & vbTab & TeamArray(j))
    Next
    fname.WriteBlankLines(2)
Next

fname.WriteLine("FEMALE TEAMS")
fname.WriteBlankLines(2)
For i = 0 to UBound(FClasses, 2) - 1
	fname.WriteLine(FClasses(1, i))

    Call GetTeams("F", FClasses(0, i))

    For j = 0 To UBound(TeamArray) - 1
        fname.WriteLine(j + 1 & vbTab & TeamArray(j))
    Next
    fname.WriteBlankLines(2)
Next

'begin download
Response.Redirect "downloads/teams_by_class.txt"

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
