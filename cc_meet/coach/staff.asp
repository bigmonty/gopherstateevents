<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim TeamStaff()

If Not Session("role") = "coach" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim TeamStaff(6, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamStaffID, FirstName, LastName, Email, MobilePhone, UserName, Password FROM TeamStaff WHERE CoachesID = " & Session("my_id")
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    For j = 0 To 6
        TeamStaff(j, i) = rs(j).Value
    Next
    i = i + 1
    ReDim Preserve TeamStaff(6, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE View Staff</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
            <h4 class="h4">Team Staff</h4>
                    
            <p>
                Team Staff are folks that have almost daily responsibilities with your team.  This utility is designed to make our site a more functional 
                site for you the coach.  It's immediate purpose is so that members of your staff can receive results emails and assist you in managing 
                rosters and line-ups.
            </p>

            <p>
                Team Staff Members are "attached" to coaches, not teams, so they have a connection to all teams managed by that coach.  They can be given 
                access to team data or not as the coach desires.  NOTE:  Only the team's head coach will have access to any staff functionality.
            </p>

            <table class="table table-striped">
                <tr>
                    <th>Name (click to edit)</th>
                    <th>Email</th>
                    <th>Mobile</th>
                    <th>UserName</th>
                    <th>Password</th>
                </tr>
                <%For i = 0 To UBound(TeamStaff, 2) - 1%>
                    <tr>
                        <td><a href="javascript:pop('edit_staff.asp?staff_id=<%=TeamStaff(0, i)%>',800,400)"><%=TeamStaff(2, i)%>, <%=TeamStaff(1, i)%></a></td>
                        <td><a href="mailto:<%=TeamStaff(3, i)%>"><%=TeamStaff(3, i)%></a></td>
                        <td><%=TeamStaff(4, i)%></td>
                        <td><%=TeamStaff(5, i)%></td>
                        <td><%=TeamStaff(6, i)%></td>
                    </tr>
                <%Next%>
            </table>
        </div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" --> 
<%	
conn.Close
Set conn = Nothing
%>
</body>
</html>
