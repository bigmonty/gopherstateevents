<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim sMeetName, sTeamName, sGender, sOrderBy, sGradeYear
Dim RacePartsArr()
Dim i, j, k
Dim lMeetID, lTeamID

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lMeetID = Request.QueryString("meet_id")
lTeamID = Request.QueryString("team_id")
	
sOrderBy = Request.QueryString("order_by")
If sOrderBy = vbNullString Then sOrderBy = "r.LastName, r.FirstName"

If sOrderBy = "name" Then sOrderBy = "r.LastName, r.FirstName"
If sOrderBy = "bib" Then sOrderBy = "ir.Bib"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value & " on " & rs(1).Value 
If Month(rs(1).Value) <=7 Then
	sGradeYear = Right(CStr(Year(rs(1).Value) - 1), 2)
Else
	sGradeYear = Right(CStr(Year(rs(1).Value)), 2)	
End If
Set rs = Nothing

'get team name, gender
sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = rs(0).Value
sGender = rs(1).Value
Set rs = Nothing

'convert gender to full word
Select Case sGender
	Case "M"
		sGender = "Male"
	Case "F"
		sGender = "Female"
End Select

j = 0
ReDim RacePartsArr(5, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.Bib, ir.IndDelay, ir.Gate, ir.RacesID FROM Roster r INNER JOIN Grades g "
sql = sql & "ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
sql = sql & "WHERE TeamsID = " & lTeamID & " AND ir.MeetsID = " & lMeetID & " AND r.Archive = 'n' ORDER BY " & sOrderBy
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	RacePartsArr(0, j) = rs(1).Value & ", " & rs(0).Value
	RacePartsArr(1, j) = rs(2).Value
    RacePartsArr(2, j) = ConvertToMinutes(rs(4).Value)
    RacePartsArr(3, j) = rs(5).Value
    RacePartsArr(4, j) = rs(3).Value
    RacePartsArr(5, j) = GetRaceName(rs(6).Value)
	j = j + 1
	ReDim Preserve RacePartsArr(5, j)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If sglScnds <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(sglScnds \ 60)
    sSecondPart = CStr(((sglScnds / 60) - (sglScnds \ 60)) * 60)
    
    'add leading zero to seconds if necessary
    If CSng(sSecondPart) < 10 Then
        sSecondPart = "0" & sSecondPart
    End If
    
    'make sure there are exactly two decimal places
    If Len(sSecondPart) < 5 Then
        If Len(sSecondPart) = 2 Then
            sSecondPart = sSecondPart & ".00"
        ElseIf Len(sSecondPart) = 4 Then
            sSecondPart = sSecondPart & "0"
        End If
    Else
        sSecondPart = Left(sSecondPart, 5)
    End If
    
    'do the conversion
    If CInt(sMinutePart) <= 60 Then
        ConvertToMinutes = sMinutePart & ":" & sSecondPart
    Else
        sHourPart = CStr(CSng(sMinutePart) \ 60)
        sMinutePart = CStr(CSng(sMinutePart) Mod 60)

        If Len(sMinutePart) = 1 Then
            sMinutePart = "0" & sMinutePart
        End If

        ConvertToMinutes = sHourPart & ":" & sMinutePart & ":" & sSecondPart
    End If
End Function

Private Function GetRaceName(lThisRace)
    sql2 = "SELECT RaceDesc FROM Races WHERE RacesID = " & lThisRace
    Set rs2 = conn.Execute(sql2)
    GetRaceName = Replace(rs2(0).Value, "''", "'")
    Set rs2 = Nothing
End Function
%>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Bib List</title>
</head>
<body>
<div class="container">
    <h3 class="h3">Bib List for <%=sTeamName%> (<%=sGender%>)</h3>
    <h4 class="h4"><%=sMeetName%></h4>
    <ul class="nav">
        <li class="nav-item"><a class="nav-link" href="bib_list.asp?meet_id=<%=lMeetID%>&amp;team_id=<%=lTeamID%>&amp;order_by=bib">Order By Bib</a></li>
        <li class="nav-item"><a class="nav-link" href="bib_list.asp?meet_id=<%=lMeetID%>&amp;team_id=<%=lTeamID%>&amp;order_by=name">Order By Name</a></li>
        <li class="nav-item"><a class="nav-link" href="javascript:window.print()">Print</a></li>
    </ul>

    <table class="table table-striped">
        <tr>
            <th>No.</th>
            <th>Name</th>
            <th>Bib</th>
            <th>Gr</th>
            <th>Race</th>
            <th>Start</th>
            <th>Gate</th>
        </tr>
		<%For j = 0 to UBound(RacePartsArr, 2) - 1%>
			<tr>
				<td><%=j + 1%></td>
				<td><%=RacePartsArr(0, j)%></td>
				<td><%=RacePartsArr(4, j)%></td>
                <td><%=RacePartsArr(1, j)%></td>
                <td><%=RacePartsArr(5, j)%></td>
				<td><%=RacePartsArr(2, j)%></td>
				<td><%=RacePartsArr(3, j)%></td>
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
