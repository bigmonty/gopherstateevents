<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim iYear
Dim sEventDir, sPhone, sEmail
Dim Events()
Dim fs, fname, sFileName

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
		
iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")
		
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(10, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate, Location, EventType, Website, WhenShutdown, EventDirID FROM Events "
sql  = sql & "WHERE EventDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear & "' ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Call RaceDirInfo(rs(7).Value)
    Events(0, i) = rs(0).Value	
    Events(1, i) = Replace(rs(1).Value, "''", "'")
    Events(2, i) = rs(2).Value	
    Events(3, i) = rs(3).Value	
    Events(4, i) = EventType(rs(4).Value)
    Events(5, i) = rs(5).Value	
    Events(6, i) = rs(6).Value	
    Events(7, i) = sEventDir
    Events(8, i) = sPhone
    Events(9, i) = sEmail
    Events(10, i) = Races(rs(0).Value)
    i = i + 1
	ReDim Preserve Events(10, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\events.txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("Event Directors")
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("EVENTID" & vbTab & "NAME" & vbTab & "DATE" & vbTab & "LOCATION" & vbTab & "TYPE" & vbTab & "WEBSITE" & vbTab & "DEADLINE" & vbTab & "DIRECTOR" & vbTab & "PHONE" & vbTab & "EMAIL" & vbTab & "RACES")
For i = 0 to UBound(Events, 2) - 1
	fname.WriteLine(Events(0, i) & vbTab & Events(1, i) & vbTab & Events(2, i) & vbTab & Events(3, i) & vbTab & Events(4, i) & vbTab & Events(5, i) & vbTab & Events(6, i) & vbTab & Events(7, i) & vbTab & Events(8, i) & vbTab & Events(9, i) & vbTab & Events(10, i))
Next

'begin download
Response.Redirect "/dwnlds/events.txt"

fname.Close
Set fname=nothing
Set fs=nothing
    
Private Function Races(lEventID)
    Dim sRaces

    Races = vbNullString

    sql2 = "SELECT RaceName FROM RaceData WHERE EventID = " & lEventID
    Set rs2 = conn.Execute(sql2)
    Do While Not rs2.EOF
        sRaces = sRaces & rs2(0).Value & ", "
        rs2.MoveNext
    Loop
    Set rs2 = Nothing

    If Not sRaces = vbNullString Then Races = Left(sRaces, Len(sRaces) - 2)
End Function

Private Function EventType(lThisType)
    sql2 = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lThisType
    Set rs2 = conn.Execute(sql2)
    EventType = rs2(0).Value
    Set rs2 = Nothing
End Function

Private Sub RaceDirInfo(lEventDir)
    sql2 = "SELECT FirstName, LastName, Email, Phone FROM EventDir WHERE EventDirID = " & lEventDir
    Set rs2 = conn.Execute(sql2)
    sEventDir = Replace(rs2(0).Value, "''", "'") & " " & Replace(rs2(1).Value, "''", "'")
    sEmail = rs2(2).Value
    sPhone = rs2(3).Value
    Set rs2 = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE Download Event List</title>
<!--#include file = "../../includes/meta2.asp" -->
</head>

<body>
    &nbsp;
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>