<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lTeamID
Dim RosterArr()
Dim i, j
Dim sSport, sGradeYear, sTeamName
Dim bNullGrade

If Not (Session("role") = "coach" Or Session("role") = "team_staff") Then Response.Redirect "/default.asp?sign_out=y"

lTeamID = Request.QueryString("team_id")
 
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
sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
rs.Open sql, conn, 1, 2
sTeamName = rs(0).Value & " (" & rs(1).Value & ")"
rs.Close
Set rs = Nothing

i = 0
ReDim RosterArr(6, 0)
sql = "SELECT RosterID, FirstName, LastName, Gender, Email, CellPhone FROM Roster WHERE TeamsID = " & lTeamID
sql = sql & " AND Archive = 'n' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RosterArr(0, i) = rs(0).Value
	RosterArr(1, i) = Replace(rs(1).Value, "''", "'")
	RosterArr(2, i) = Replace(rs(2).Value, "''", "'")
	RosterArr(3, i) = GetGrade(rs(0).Value)
	RosterArr(4, i) = rs(3).Value
    RosterArr(5, i) = rs(4).Value
    RosterArr(6, i) = rs(5).Value
	i = i + 1
	ReDim Preserve RosterArr(6, i)
	rs.MoveNext
Loop
Set rs = Nothing
	
Private Function GetGrade(lMyID)
    GetGrade = "0"

    bNullGrade = False
	Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then 
        GetGrade = rs2(0).Value
        bNullGrade = True
    End If
	rs2.Close
	Set rs2 = Nothing

    If bNullGrade = False Then
        sql2 = "INSERT INTO Grades (RosterID,  Grade" & sGradeYear & ") Values (" & lMyID & ", 0)"
        Set rs2 = conn.Execute(sql2)
        Set rs2 = Nothing
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Print Roster</title>
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="My CC/Nordic History">
	<h4 class="h4">Print GSE Roster: <%=sTeamName%></h4>

	<table class="table table-striped table-condensed table-responsive">
		<tr>
			<th>No.</th>
            <th>Roster ID</th>
			<th>Name (click to edit)</th>
			<th>Grade</th>
			<th>M/F</th>
            <th>Email</th>
            <th>Cell</th>
		</tr>
		<%For i = 0 to UBound(RosterArr, 2) - 1%>
			<tr>
				<td><%=i +1%>)</td>
                <td><%=RosterArr(0, i)%></td>
				<td><a href="javascript:pop('edit_part.asp?roster_id=<%=RosterArr(0, i)%>',1000,600)"><%=RosterArr(2, i)%>,<%=RosterArr(1, i)%></a></td>
				<td><%=RosterArr(3, i)%></td>
				<td><%=RosterArr(4, i)%></td>
				<td><a href="mailto:<%=RosterArr(5, i)%>"><%=RosterArr(5, i)%></a></td>
				<td><%=RosterArr(6, i)%></td>
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
