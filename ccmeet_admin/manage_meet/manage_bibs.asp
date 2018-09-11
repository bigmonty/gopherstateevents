<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k, p, n
Dim lThisMeet
Dim iGrade, iFirstBib, iLastBib
Dim sMeetName, sGradeYear, sAssign, sTeamRange, sRemove
Dim dMeetDate
Dim BibArray(), MeetTeams(), AssgndBibs(), AvailBibs(), TeamParts()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

sAssign = Request.QueryString("assign")
sTeamRange = Request.QueryString("team_range")
sRemove = Request.QueryString("remove")

Server.ScriptTimeout=1200

'Response.Buffer = False
Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RosterID, Bib FROM IndRslts WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
        rs(1).Value = Request.Form.Item("bib_" & rs(0).Value)
	    rs.Update
        rs.MoveNext
    Loop
	rs.Close
	Set rs = Nothing
End If

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

'get meet teams array
i = 0
ReDim MeetTeams(1, 0)
sql = "SELECT mt.TeamsID, t.TeamName, t.Gender FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName, t.Gender"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve MeetTeams(1, i)
	rs.MoveNext
Loop
Set rs = Nothing
 
'get year for roster grades
If Month(dMeetDate) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If
	
If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

If sRemove = "all" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		rs(0).Value = 0
		rs.Update
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End If

Call GetAvailBibs

If sAssign = "all" Then
	If sTeamRange = "y" Then
	Else
		For n = 0 To UBound(MeetTeams, 2) - 1
			Call GetTeamParts(MeetTeams(0, n))
			
            If UBound(TeamParts) > 0 Then
			    For p = 0 To UBound(TeamParts) - 1
				    Set rs2 = Server.CreateObject("ADODB.Recordset")
				    sql2 = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND RosterID = " & TeamParts(p)
				    rs2.Open sql2, conn, 1, 2
				    rs2(0).Value = AvailBibs(p)
				    rs2.Update
				    rs2.Close
				    Set rs2 = Nothing
			    Next
			
			    Call GetAvailBibs
            End If
		Next
	End If
End If

Private Sub GetAvailBibs()
	'get meet bib range
	sql = "SELECT BibStart, BibEnd FROM BibRange WHERE MeetsID = " & lThisMeet
	Set rs = conn.Execute(sql)
	iFirstBib = rs(0).Value
	iLastBib = rs(1).Value
	Set rs = Nothing
	
	i = 0
	ReDim AssgndBibs(0)
	sql = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND Bib > 0 ORDER BY Bib"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    AssgndBibs(i) = rs(0).Value
	    i = i + 1
	    ReDim Preserve AssgndBibs(i)
	    rs.MoveNext
	Loop
	Set rs = Nothing

	k = 0
	ReDim AvailBibs(0)
	For i = iFirstBib To iLastBib
		If UBound(AssgndBibs) = 0 Then
			AvailBibs(k) = i
			k = k + 1
			ReDim Preserve AvailBibs(k)
		Else
			For j = 0 To UBound(AssgndBibs) - 1
				If CInt(AssgndBibs(j)) = CInt(i) Then 
					Exit For
				Else
					If j = UBound(AssgndBibs) - 1 Then
						AvailBibs(k) = i
						k = k + 1
						ReDim Preserve AvailBibs(k)
					End If
				End If
		    Next
		End If
	Next
End Sub

Private Sub GetTeamParts(lTeamID)
	Dim x
	
	ReDim TeamParts(0)
	sql2 = "Select ir.RosterID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE r.TeamsID = " & lTeamID 
	sql2 = sql2 & " AND MeetsID = " & lThisMeet & " AND ir.Bib = 0 ORDER BY r.LastName, r.FirstName"
	Set rs2 = conn.Execute(sql2)
	Do While Not rs2.EOF
		TeamParts(x) = rs2(0).Value
		x = x + 1
		ReDim Preserve TeamParts(x)
		rs2.MoveNext
	Loop
	Set rs2 = Nothing
End Sub

i = 0
ReDim BibArray(6, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.FirstName, r.LastName, t.TeamName, t.Gender, r.Gender, g.Grade" & sGradeYear & ", ir.RacesID, ir.Bib, r.RosterID "
sql = sql & "FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
sql = sql & "INNER JOIN Teams t ON r.TeamsID = t.TeamsID WHERE ir.MeetsID = " & lThisMeet & " ORDER BY t.TeamName, t.Gender, r.LastName, r.FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	BibArray(0, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(0).Value, "''", "'") 
	BibArray(1, i) = rs(2).Value & " (" & rs(3).Value & ")"
	BibArray(2, i) = rs(4).Value
	BibArray(3, i) = rs(5).Value 
	BibArray(4, i) = RaceName(rs(6).Value)
	BibArray(5, i) = rs(7).Value
    BibArray(6, i) = rs(8).Value
	i = i + 1
	ReDim Preserve BibArray(6, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function TeamName(lTeamID)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
	rs2.Open sql2, conn, 1, 2
	TeamName = Replace(rs2(0).Value, "''", "'")
	rs2.Close
	Set rs2 = Nothing
End Function

Private Function RaceName(lRaceID)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RaceName FROM Races WHERE RacesID = " & lRaceID
	rs2.Open sql2, conn, 1, 2
	RaceName = Replace(rs2(0).Value, "''", "'")
	rs2.Close
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CCMeet Bib Manager</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<!--#include file = "manage_meet_nav.asp" -->
			
			<h4 class="h4">CCMeet Bib Manager: <%=sMeetName%></h4>
			
            <ul class="nav bg-info">
				<li class="nav-item">
					<a class="nav-link" href="/ccmeet_admin/manage_meet/dwnld_bib_list.asp?meet_id=<%=lThisMeet%>"
                	onclick="openThis(this.href,1024,768);return false;" style="color:#fff;">Download Bib List</a>
				</li>
            </ul>

			<div>
                <ul class="nav">
					<li class="nav-item" style="font-weight:bold;padding-top:8px;">Batch Assign:</li>
				    <li class="nav-item">
                        <a class="nav-link" href="manage_bibs.asp?meet_id=<%=lThisMeet%>&amp;assign=all&amp;team_range=n">All Missing</a>
                    </li>
				    <li class="nav-item">
                        <a class="nav-link" href="manage_bibs.asp?meet_id=<%=lThisMeet%>&amp;assign=all&amp;team_range=y">All Missing (Use Team Bib Range)</a>
                    </li>
				    <li class="nav-item">
                        <a class="nav-link" href="manage_bibs.asp?meet_id=<%=lThisMeet%>&amp;remove=all">Remove All (NO UNDO!)</a>
                    </li>
                </ul>
			</div>

			<h4 class="h4">Bib List</h4>
				
            <form class="form" name="assign_these" method="post" action="manage_bibs.asp?meet_id=<%=lThisMeet%>">
            <table class="table table-striped">
                <tr>
                    <td style="text-align: center;background-color:#ececec;" colspan="7">
                        <input type="hidden" name="submit_this" id="submit_this" value="submit_this">
                        <input type="submit" class="form-control" name="submit1" id="submit1" value="Save Changes">
                    </td>
                </tr>
                <tr><th>No.</th><th>Name</th><th>Team (Gender)</th><th>MF</th><th>Gr</th><th>Race</th><th>Bib</th></tr>
                <%For i = 0 To UBound(BibArray, 2) - 1%>
                    <tr>
                        <td><%=i + 1%>)</td>
                        <%For j = 0 To 4%>
                            <td><%=BibArray(j, i)%></td>
                        <%Next%>
                        <td><input type="text" class="form-control" name="bib_<%=BibArray(6, i)%>" id="" value="<%=BibArray(5, i)%>" size="3"></td>
                    </tr>
                <%Next%>
            </table>
            </form>
        </div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.Close
Set conn=Nothing
%>
</html>
