<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim CoachArr()
Dim fs, fname, sFileName

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim CoachArr(6, 0)
sql = "SELECT CoachesID,  FirstName, LastName, Email, Phone, UserID, Password FROM Coaches ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	CoachArr(0, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	CoachArr(1, i) = GetTeam(rs(0).Value)
	CoachArr(2, i) = GetSport(rs(0).Value)
	CoachArr(3, i) = rs(3).Value
	CoachArr(4, i) = rs(4).Value
    CoachArr(5, i) = rs(5).Value
    CoachArr(6, i) = rs(6).Value
	i = i + 1
	ReDim Preserve CoachArr(6, i)
	rs.MoveNext
Loop
Set rs = Nothing

Function GetSport(lCoachID)
    GetSport = "unknown"
    Set rs2 = SErver.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Sport FROM Teams WHERE CoachesID = " & lCoachID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetSport = rs2(0).Value
    rs2.Close
	Set rs2 = Nothing
End Function

Function GetTeam(lCoachID)
    GetTeam = "unknown"
    Set rs2 = SErver.CreateObject("ADODB.Recordset")
	sql2 = "SELECT TeamName FROM Teams WHERE CoachesID = " & lCoachID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetTeam = rs2(0).Value
    rs2.Close
	Set rs2 = Nothing
End Function

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\ccmeet_admin\manage_coach\downloads\coach_data.txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine(UCase("GSE Coach Data"))
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("NAME" & vbTab & "TEAM(S)" & vbTab & "SPORT" & vbTab & "PHONE" & vbTab & "EMAIL" & vbTab & "USER ID" & vbTab & "PASSWORD")
For i = 0 to UBound(CoachArr, 2) - 1
	fname.WriteLine(CoachArr(0, i) & vbTab & CoachArr(1, i) & vbTab & CoachArr(2, i) & vbTab & CoachArr(3, i) & vbTab & CoachArr(4, i) & vbTab & CoachArr(5, i) & vbTab & CoachArr(6, i))
Next

'begin download
Response.Redirect "downloads/coach_data.txt"

fname.Close
Set fname=nothing
Set fs=nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>CCMeet Coach Data</title>

<!--#include file = "../../includes/js.asp" -->
</head>
<body>

<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
