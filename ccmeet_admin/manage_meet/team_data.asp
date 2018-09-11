<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2, rs3, sql3
Dim i, j, k
Dim lThisMeet
Dim sMeetName, sCoachName, sCoachEmail, sTeamClass, sSortBy
Dim iNumEntrants, iMTotal, iFTotal, iTotalEntrants
Dim MTeams(), FTeams(), MeetArr(), SortArr(5)
Dim dMeetDate

Dim sMapLink, sMeetInfoSheet, sCourseMap
Dim dWhenShutdown

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

sSortBy = Request.QueryString("sort_by")
If sSortBy = vbNullString Then sSortBy = "team"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Not CLng(lThisMeet) = 0 Then
    sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
    Set rs = conn.Execute(sql)
    sMeetName = Replace(rs(0).Value, "''", "'")
    dMeetDate = rs(1).Value
    Set rs = Nothing

	'get maplink
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMapLink = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT InfoSheet FROM MeetInfo WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMeetInfoSheet = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get course map
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Map FROM CourseMap WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sCourseMap = rs(0).Value
	rs.Close
	Set rs = Nothing

    'get participating teams info	
    ReDim MTeams(5, 0)
    i = 0
    sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'M' ORDER BY t.TeamName"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        Call GetTeamData(rs(0).Value)

	    MTeams(0, i) = rs(0).Value
	    MTeams(1, i) = Replace(rs(1).Value, "''", "'")
        MTeams(2, i) = sCoachName
        MTeams(3, i) = sCoachEmail
        MTeams(4, i) = sTeamClass
        MTeams(5, i) = iNumEntrants

        iMTotal = CInt(iMTotal) + iNumEntrants
        iTotalEntrants = CInt(iTotalEntrants) + iNumEntrants

	    i = i + 1
	    ReDim Preserve MTeams(5, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing

    ReDim FTeams(5, 0)
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'F' ORDER BY t.TeamName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Call GetTeamData(rs(0).Value)

	    FTeams(0, i) = rs(0).Value
	    FTeams(1, i) = Replace(rs(1).Value, "''", "'")
        FTeams(2, i) = sCoachName
        FTeams(3, i) = sCoachEmail
        FTeams(4, i) = sTeamClass
        FTeams(5, i) = iNumEntrants

        iFTotal = CInt(iFTotal) + iNumEntrants
        iTotalEntrants = CInt(iTotalEntrants) + iNumEntrants

	    i = i + 1
	    ReDim Preserve FTeams(5, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing
End If

Private Sub GetTeamData(lThisTeam)
    sCoachName = vbNullString
    sCoachEmail = vbNullString
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT c.FirstName, c.LastName, c.Email FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID WHERE t.TeamsID = " & lThisTeam
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then
        sCoachName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
        sCoachEmail = rs2(2).Value
    End If
    rs2.Close
    Set rs2 = Nothing

    sTeamClass = "n/a"
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet & " AND TeamsID = " & lThisTeam
    rs2.Open sql2, conn, 1, 2
    If Not rs2(0).Value & "" = "" Then sTeamClass = ClassName(rs2(0).Value)
    rs2.Close
    Set rs2 = Nothing

    iNumEntrants = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT ir.RosterID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.MeetsID = " & lThisMeet 
    sql2 = sql2 & " AND r.TeamsID = " & lThisTeam
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then iNumEntrants = rs2.RecordCount
    rs2.Close
    Set rs2 = Nothing
End Sub

Private Function ClassName(lThisClass)
    Set rs3 = Server.CreateObject("ADODB.Recordset")
    sql3 = "SELECT ClassName FROM MeetClasses WHERE MeetClassesID = " & lThisClass 
    rs3.Open sql3, conn, 1, 2
    ClassName = rs3(0).Value
    rs3.Close
    Set rs3 = Nothing
End Function

If sSortBy = "num_parts" Then
    For i = 0 To UBound(MTeams, 2) - 2
        For j = i + 1 To UBound(MTeams, 2) - 1
            If CInt(MTeams(5, i)) < CInt(MTeams(5, j)) Then
                For k = 0 To 5
                    SortArr(k) = MTeams(k, i)
                    MTeams(k, i) = MTeams(k, j)
                    MTeams(k, j)= SortArr(k)
                Next
            End If
        Next
    Next

    For i = 0 To UBound(FTeams, 2) - 2
        For j = i + 1 To UBound(FTeams, 2) - 1
            If CInt(FTeams(5, i)) < CInt(FTeams(5, j)) Then
                For k = 0 To 5
                    SortArr(k) = FTeams(k, i)
                    FTeams(k, i) = FTeams(k, j)
                    FTeams(k, j)= SortArr(k)
                Next
            End If
        Next
    Next
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Team Data</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			
			    <h4 class="h4">Team Data: <%=sMeetName%> on <%=dMeetDate%></h4>
			
                <span style="font-weight: bold;">Total Entrants:</span>&nbsp;<%=iTotalEntrants%><br>

                <ul class="nav">
                   <li class="nav-item">Sort By:</li>
                    <li class="nav-item"><a class="nav-link" href="team_data.asp?sort_by=team&amp;meet_id=<%=lThisMeet%>">Team</a></li>
                    <li class="nav-item"><a class="nav-link" href="team_data.asp?sort_by=num_parts&amp;meet_id=<%=lThisMeet%>">Number of Entrants</a></li>
                </ul>

				<h4 class="h4">Female Teams</h4>
				<table class="table table-striped">
                    <tr><th>No.</th><th>Team</th><th>Coach</th><th>Email</th><th>Class</th><th>Ent</th></tr>
                    <%For i = 0 To UBound(FTeams, 2) - 1%>
                        <tr>
                            <td style="text-align: right;"><%=i + 1%>)</td>
                            <td><%=FTeams(1, i)%></td>
                            <td><%=FTeams(2, i)%></td>
                            <td><a href="mailto:<%=FTeams(3, i)%>"><%=FTeams(3, i)%></a></td>
                            <td><%=FTeams(4, i)%></td>
                            <td style="text-align: right;"><%=FTeams(5, i)%></td>
                        </tr>
                    <%Next%>
                    <tr><th style="text-align: right;" colspan="5">Total:</th><th style="text-align: right;"><%=iFTotal%></th></tr>
                </table>

				<h4 class="h4">Male Teams</h4>
				<table class="table table-striped">
                    <tr><th>No.</th><th>Team</th><th>Coach</th><th>Email</th><th>Class</th><th>Ent</th></tr>
                    <%For i = 0 To UBound(MTeams, 2) - 1%>
                        <tr>
                            <td style="text-align: right;"><%=i + 1%>)</td>
                            <td><%=MTeams(1, i)%></td>
                            <td><%=MTeams(2, i)%></td>
                            <td><a href="mailto:<%=MTeams(3, i)%>"><%=MTeams(3, i)%></a></td>
                            <td><%=MTeams(4, i)%></td>
                            <td style="text-align: right;"><%=MTeams(5, i)%></td>
                        </tr>
                    <%Next%>
                    <tr><th style="text-align: right;" colspan="5">Total:</th><th style="text-align: right;"><%=iMTotal%></th></tr>
                </table>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
