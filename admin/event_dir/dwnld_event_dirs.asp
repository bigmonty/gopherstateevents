<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim EventDir()
Dim i, j
Dim fs, fname, sFileName

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim EventDir(7, 0)
sql = "SELECT FirstName, LastName, Address, City, State, Zip, Phone, Email FROM EventDir WHERE Active = 'y' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventDir(0, i) = Replace(rs(0).Value, "''", "'")
	EventDir(1, i) = Replace(rs(1).Value, "''", "'")
	If Not rs(2).Value & "" = "" Then EventDir(2, i) = Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then EventDir(3, i) = Replace(rs(3).Value, "''", "'")
	EventDir(4, i) = rs(4).Value
	EventDir(5, i) = rs(5).Value
	EventDir(6, i) = rs(6).Value
	EventDir(7, i) = rs(7).Value
	i = i + 1
	ReDim Preserve EventDir(7, i)
	rs.MoveNext
Loop
Set rs = Nothing

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\event_directors.txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("Event Directors")
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("FIRST" & vbTab & "LAST" & vbTab & "ADDRESS" & vbTab & "CITY" & vbTab & "ST" & vbTab & "ZIP" & vbTab & "PHONE" & vbTab & "EMAIL")
For i = 0 to UBound(EventDir, 2) - 1
	fname.WriteLine(EventDir(0, i) & vbTab & EventDir(1, i) & vbTab & EventDir(2, i) & vbTab & EventDir(3, i) & vbTab & EventDir(4, i) & vbTab & EventDir(5, i) & vbTab & EventDir(6, i) & vbTab & EventDir(7, i))
Next

'begin download
Response.Redirect "/dwnlds/event_directors.txt"

fname.Close
Set fname=nothing
Set fs=nothing

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE&copy; Event Directors Download</title>

<!--#include file = "../../includes/js.asp" -->

</head>

<body>
&nbsp;
</body>
</html>
