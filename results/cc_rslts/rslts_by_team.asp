<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs
Dim i, j
Dim lMeetID, lRaceID
Dim sRaceName, sMeetName, sGradeYear
Dim MeetTeams, OurRslts
Dim dMeetDate

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
	
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
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

sql = "SELECT t.TeamsID, t.TeamName, t.Gender FROM Teams t INNER JOIN MeetTeams mt ON t.TeamsID = mt.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lMeetID & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
MeetTeams = rs.GetRows()
Set rs = Nothing

sql = "SELECT RaceDesc FROM Races WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'")
Set rs = Nothing

Private Sub GetTeamRslts(lThisTeam)
	sql = "SELECT r.RosterID, ir.Bib, r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.RaceTime FROM IndRslts ir INNER JOIN Roster r"
	sql = sql & " ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID "
	sql = sql & "WHERE ir.RacesID = " & lRaceID & " AND r.TeamsID = " & lThisTeam & " AND ir.Place > 0 AND ir.RaceTime > '00:00' AND ir.Excludes = 'n' "
	sql = sql & "ORDER BY ir.FnlScnds, ir.Place"
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
	    ReDim OurRslts(5, 0)
    Else
	    OurRslts = rs.GetRows()
    End If
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Gopher State Events CC/Nordic Results By Team</title>
<meta name="description" content="Cross-Country & Nordic Ski Results by Gopher State Events, a conventional timing service offererd by H51 Software, LLC in Minnetonka, MN.">
  
<!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <div class="row">
        <img src="/graphics/html_header.png" class="img-responsive" alt="Individual Results">
	    <h4 class="h4">Individual Results By Team <br><small><%=sMeetName%> on <%=dMeetDate%><br><%=sRaceName%></small></h4>

	    <div class="bg-warning">
		    <a href="javascript:window.print();">Print</a>
	    </div>
    </div>

    <div class="row">
        <%For j = 0 To UBound(MeetTeams, 2)%>
            <%Call GetTeamRslts(MeetTeams(0, j))%>
            
            <%If UBound(OurRslts, 2) > 0 Then%>
                <h4 class="h4 bg-info"><%=MeetTeams(1, j)%></h4>

	            <table class="table table-striped">
		            <tr>
			            <th style="text-align:right">Pl</th>
			            <th>Bib-Name</th>
			            <th>Gr</th>
			            <th style="text-align:center">Time</th>
		            </tr>
			        <%For i = 0 to UBound(OurRslts, 2)%>
					    <tr>
						    <td><%=i + 1%></td>
						    <td><%=OurRslts(1, i)%>-<%=OurRslts(3, i)%>, <%=OurRslts(2, i)%></td>
						    <td><%=OurRslts(4, i)%></td>
						    <td><%=OurRslts(5, i)%></td>
					    </tr>
			        <%Next%>
		        </table>
            <%End If%>
        <%Next%>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
