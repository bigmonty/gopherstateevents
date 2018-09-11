<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim MeetDirArr()
Dim lMeetDirID
Dim fs, fname, sFileName

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim MeetDirArr(5, 0)
sql = "SELECT MeetDirID,  FirstName, LastName, Email, Phone, UserID, Password FROM MeetDir ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetDirArr(0, i) = rs(0).Value
	MeetDirArr(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	MeetDirArr(2, i) = rs(3).Value
	MeetDirArr(3, i) = rs(4).Value
	MeetDirArr(4, i) = rs(5).Value
	MeetDirArr(5, i) = rs(6).Value
	i = i + 1
	ReDim Preserve MeetDirArr(5, i)
	rs.MoveNext
Loop
Set rs = Nothing

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\inetpub\h51web\gopherstateevents\dwnlds\meet_directors.txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE Meet Directors")
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(2)

fname.WriteLine("ID NUM" & vbTab & "NAME " & vbTab & "EMAIL ADDRESS" & vbTab & "TELEPHONE" & vbTab & "USER NAME" & vbTab & "PASSWORD")
For i = 0 to UBound(MeetDirArr, 2) - 1
	fname.WriteLine(MeetDirArr(0, i) & vbTab & MeetDirArr(1, i) & vbTab & MeetDirArr(2, i) & vbTab & MeetDirArr(3, i) & vbTab & MeetDirArr(4, i) & vbTab & MeetDirArr(5, i))
Next

fname.Close
Set fname=nothing
Set fs=nothing

'begin download
Response.Redirect "../../dwnlds/meet_directors.txt"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE CC Meet Director Data</title>
<!--#include file = "../../includes/meta2.asp" -->

</head>
<body>
&nbsp;
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
