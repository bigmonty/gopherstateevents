<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lPartID
Dim sPartName
Dim MyRaces()

lPartID = Request.QueryString("part_id")
If CStr(lPartID) & "" = "" Then lPartID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

ReDim MyRaces(1, 0)
If Not CLng(lPartID) = 0 Then
	sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
	Set rs = conn.Execute(sql)
	sPartName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
	Set rs = Nothing

    i = 0
    Set rs= Server.CreateObject("ADODB.Recordset")
    sql = "SELECT e.EventName, e.EventDate FROM Events e INNER JOIN RaceData r ON e.EventID = r.EventID INNER JOIN PartRace pr ON r.RaceID = pr.RaceID "
    sql = sql & "WHERE pr.ParticipantID = " & lPartID & " ORDER BY e.EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MyRaces(0, i) = Replace(rs(0).Value, "''", "'")
        MyRaces(1, i) = rs(1).Value
        i =  i + 1
        ReDim Preserve MyRaces(1, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; User History</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
    <h3 class="h3">User History: <%=sPartName%> (<%=lPartID%>)</h3>

    <ul class="list">
        <%For i = 0 To UBound(MyRaces, 2) - 1%>
            <li><%=MyRaces(0, i)%>&nbsp;(<%=MyRaces(1, i)%>)</li>
        <%Next%>
    </ul>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
