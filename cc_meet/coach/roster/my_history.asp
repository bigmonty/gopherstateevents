<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lTeamID, lRosterID
Dim i, j, k
Dim sMyName, sGender, sGradeYear, sTeamName
Dim MyRslts(), SortArr(9)

lRosterID = Request.QueryString("roster_id")
If CStr(lRosterID) = vbNullString Then lRosterID = 0
If Not IsNumeric(lRosterID) Then Response.Redirect "http://www.google.com"
If CLng(lRosterID) < 0 Then Response.Redirect "http://www.google.com"

'get year for roster grades
If Month(Date) <=7 Then
	sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
Else
	sGradeYear = Right(CStr(Year(Date)), 2)	
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Gender, TeamsID FROM Roster WHERE RosterID = " & lRosterID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sMyName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
    sGender = rs(2).Value
    lTeamID = rs(3).Value
End If
rs.Close
Set rs = Nothing

sql = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = Replace(rs(0).Value, "''", "'")
Set rs = Nothing

i = 0
ReDim MyRslts(9, 0)
sql = "SELECT MeetsID, RacesID, RaceTime FROM IndRslts WHERE RosterID = " & lRosterID & " AND FnlScnds > 0 AND Place > 0"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    MyRslts(0, i) = rs(0).Value
    MyRslts(1, i) = rs(1).Value
    MyRslts(2, i) = RacePlace(rs(0).Value, lRosterID)
    MyRslts(3, i) = rs(2).Value
    MyRslts(4, i) = PlaceOnTeam(rs(0).Value, lRosterID)
    i = i + 1
    ReDim Preserve MyRslts(9, i)
    rs.MoveNext
Loop
Set rs = Nothing

For i = 0 To UBound(MyRslts, 2) - 1
    sql = "SELECT MeetName, MeetDate, MeetSite FROM Meets WHERE MeetsID = " & MyRslts(0, i)
    Set rs = conn.Execute(sql)
    MyRslts(5, i) = Replace(rs(0).Value, "''", "'")
    MyRslts(6, i) = rs(1).Value
    MyRslts(7, i) = Replace(rs(2).Value, "''", "'")
    Set rs = Nothing

    sql = "SELECT RaceDesc, RaceDist, RaceUnits FROM Races WHERE RacesID = " & MyRslts(1, i)
    Set rs = conn.Execute(sql)
    MyRslts(8, i) = Replace(rs(0).Value, "''", "'")
    MyRslts(9, i) = rs(1).Value & " " & rs(2).Value
    Set rs = Nothing
Next

'sort by date
If UBound(MyRslts, 2) > 0 Then
    For i = 0 To UBound(MyRslts, 2) - 2
        For j = i + 1 To UBound(MyRslts, 2) - 1 
            If CDate(MyRslts(6, i)) < CDate(MyRslts(6, j)) Then
                For k = 0 To 9
                    SortArr(k) = MyRslts(k, i)
                    MyRslts(k, i) = MyRslts(k, j)
                    MyRslts(k, j) = SortArr(k)
                Next
            End If
        Next
    Next
End If

Private Function PlaceOnTeam(lThisMeet, lMyID)
    Dim x
    Dim lThisTeam

    sql2 = "SELECT TeamsID FROM Roster WHERE RosterID = " & lMyID 
    Set rs2 = conn.Execute(sql2)
    lThisTeam = rs2(0).Value
    Set rs2 = Nothing

    x = 0
    sql2 = "SELECT ir.RosterID, r.TeamsID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.MeetsID = " & lThisMeet 
    sql2 = sql2 & " AND r.TeamsID = " & lThisTeam & " AND ir.FnlScnds > 0 AND ir.Place > 0 ORDER BY ir.FnlScnds"
    Set rs2 = conn.Execute(sql2)
    Do While Not rs2.EOF
        x = x + 1
        If CLng(rs2(0).Value) = CLng(lMyID) Then
            PlaceOnTeam = x
            Exit Do
        End If
        rs2.MoveNext
    Loop
    Set rs2 = Nothing
End Function

Private Function RacePlace(lThisMeet, lMyID)
    Dim x

    x = 0
    sql2 = "SELECT RosterID FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND FnlScnds > 0 AND Place > 0 ORDER BY FnlScnds"
    Set rs2 = conn.Execute(sql2)
    Do While Not rs2.EOF
        x = x + 1
        If CLng(rs2(0).Value) = CLng(lMyID) Then
            RacePlace = x
            Exit Do
        End If
        rs2.MoveNext
    Loop
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Individual History</title>
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="My CC/Nordic History">
	<h4 class="h4">GSE Cross-Country/Nordic Ski Individual History - <%=sTeamName%></h4>

    <a href="javascript:window.print();">Print</a>

    <h5 class="h5"><%=sMyName%> (<%=sGender%>)</h5>
    <table class="table table-striped">
        <tr>
            <th>No.</th>
            <th>Meet</th>
            <th>Date</th>
            <th>Location</th>
            <th>Race</th>
            <th>Dist</th>
            <th>Time</th>
            <th>Race Pl</th>
            <th>Team Pl</th>
        </tr>
        <%For i = 0 To UBound(MyRslts, 2) - 1%>
            <tr>
                <td><%=i + 1%>)</td>
                <td><%=MyRslts(5, i)%></td>
                <td><%=MyRslts(6, i)%></td>
                <td><%=MyRslts(7, i)%></td>
                <td><%=MyRslts(8, i)%></td>
                <td><%=MyRslts(9, i)%></td>
                <td><%=MyRslts(3, i)%></td>
                <td><%=MyRslts(2, i)%></td>
                <td><%=MyRslts(4, i)%></td>
            </tr>
        <%Next%>
    </table>
</div>
<!--#include file = "../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
