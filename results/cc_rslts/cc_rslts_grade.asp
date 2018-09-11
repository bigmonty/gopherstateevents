<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn2, rs, rs2, sql2
Dim i, j, k
Dim lMeetID, lRaceID
Dim sRaceName, sMeetName, sGradeYear, sOrderResultsBy
Dim RsltsArr, Grades(1, 5)
Dim dMeetDate

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

Grades(0, 0) = "7"
Grades(1, 0) = "7th Grade"
Grades(0, 1) = "8"
Grades(1, 1) = "8th Grade"
Grades(0, 2) = "9"
Grades(1, 2) = "9th Grade"
Grades(0, 3) = "10"
Grades(1, 3) = "10th Grade"
Grades(0, 4) = "11"
Grades(1, 4) = "11th Grade"
Grades(0, 5) = "12"
Grades(1, 5) = "12th Grade"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn2 = Server.CreateObject("ADODB.connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
	
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn2.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value
Set rs = Nothing
	
'get year for roster grades
If Month(dMeetDate) <= 7 Then
	sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
	sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

sql = "SELECT RaceDesc, OrderBy FROM Races WHERE RacesID = " & lRaceID
Set rs = conn2.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'")
sOrderResultsBy = rs(1).Value
Set rs = Nothing

'get all results and then parse them out later by grade
If sOrderResultsBy = "time" Then
	sql = "SELECT ir.Bib, r.LastName, r.FirstName, t.TeamName, ir.RaceTime, g.Grade" & sGradeYear & " FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID "
    sql = sql & "INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID 
    sql = sql & " AND ir.Place > 0 AND ir.RaceTime > '00:00' AND Excludes = 'n' ORDER BY ir.FnlScnds, ir.Place"
Else
	sql = "SELECT ir.Bib, r.LastName, r.FirstName, t.TeamName, ir.RaceTime, g.Grade" & sGradeYear & " FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID "
    sql = sql & "INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID 
    sql = sql & " AND ir.Place > 0 AND ir.RaceTime > '00:00' AND Excludes = 'n' ORDER BY ir.Place"
End If
Set rs = conn2.Execute(sql)
If rs.BOF and rs.EOF Then
   ReDim RsltsArr(5, 0)
Else
    RsltsArr = rs.GetRows()
End If
Set rs = Nothing

For i = 0 To UBound(RsltsArr, 2)
    RsltsArr(1, i) = Replace(RsltsArr(1, i), "''", "'")
    RsltsArr(2, i) = Replace(RsltsArr(2, i), "''", "'")
    RsltsArr(3, i) = Replace(RsltsArr(3, i), "''", "'")
'    If Not RsltsArr(5, i) & "" = "" Then RsltsArr(5, i) = GetGrade(RsltsArr(5,i))
Next
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Gopher State Events CC/Nordic Results By Grade</title>
<meta name="description" content="Cross-Country & Nordic Ski Results by Gopher State Events, a conventional timing service offererd by H51 Software, LLC in Minnetonka, MN.">
 
<!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <div class="row">
        <img src="/graphics/html_header.png" class="img-responsive" alt="Individual Results">
	    <h4 class="h4">Grade Level Results<br><%=sMeetName%> on <%=dMeetDate%><br><%=sRaceName%></h4>
		<br>
	    <div>
		    <a href="javascript:window.print();">Print</a>
	    </div>
    </div>

    <div class="row">
         <%For j = 5 To 0 Step -1%>
	        <table class="table table-striped">
                <tr><th colspan="5"><%=Grades(1, j)%></th></tr>
		        <tr>
			        <th>Pl</th>
			        <th>Bib-Name</th>
			        <th>Team</th>
			        <th>Time</th>
                    <th>Grade</th>
		        </tr>
			    <%k = 0%>
			    <%For i = 0 to UBound(RsltsArr, 2)%>
                    <%If CInt(RsltsArr(5, i)) = CInt(Grades(0, j)) Then%>
					    <%k = k + 1%>
						<tr>
							<td><%=k%></td>
							<td><%=RsltsArr(0, i)%> - <%=RsltsArr(2, i)%>, <%=RsltsArr(1, i)%></td>
							<td><%=RsltsArr(3, i)%></td>
							<td><%=RsltsArr(4, i)%></td>
                            <td><%=RsltsArr(5, i)%></td>
						</tr>
                    <%End If%>
			    <%Next%>
		    </table>
        <%Next%>
    </div>
</div>
<%
conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
