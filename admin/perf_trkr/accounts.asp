<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim sTeamName, sSport
Dim Accounts()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Accounts(8, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT pt.PerfTrkrID, r.FirstName, r.LastName, r. Gender, pt.Email, pt.CellPhone, pt.WhenSubscr, r.TeamsID, "
sql = sql & "Expiration FROM PerfTrkr pt INNER JOIN Roster r ON pt.RosterID = r.RosterID ORDER BY r.LastName, r.FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Call TeamData(rs(7).Value)
	Accounts(0, i) = rs(0).Value
    Accounts(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    Accounts(2, i) = rs(3).Value
    Accounts(3, i) = sTeamName
    Accounts(4, i) = sSport
	Accounts(5, i) = rs(4).Value
	Accounts(6, i) = rs(5).Value
	Accounts(7, i) = rs(6).Value
	Accounts(8, i) = rs(8).Value
	i = i + 1
	ReDim Preserve Accounts(8, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub TeamData(lTeamID)
	Dim rs2, sql2

	sTeamName = "unknown"
	sSport = "unknown"

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT TeamName, Sport FROM Teams WHERE TeamsID = " & lTeamID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then
		sTeamName = Replace(rs2(0).Value, "''", "'")
		sSport = rs2(1).Value
	End If
	rs2.Close
	Set rs2 = Nothing
End Sub
%>
<<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE&copy; Performance Tracker Accounts</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <!--#include file = "perf_trkr_nav.asp" -->

		    <h4 class="h4">GSE Performance Tracker Accounts</h4>
		
		    <table class="table table-striped">
			    <tr>
				    <th>Name</th>
                    <th>M/F</th>
 					<th>Team</th>
					<th>Sport</th>
					<th>Email</th>
					<th>Mobile</th>
				    <th>When Reg</th>
					<th>Expires</th>
			    </tr>
			    <%For i = 0 to UBound(Accounts, 2) - 1%>
				    <tr>
						<%For j = 1 to 8%>
							<td style="white-space:nowrap;">
								<%Select Case j%>
									<%Case 1%>
										<a href="javascript:pop('edit_accnt.asp?accnt_id=<%=Accounts(0, i)%>',800,400)"><%=Accounts(1, i)%></a>
									<%Case 5%>
										<a href="mailto:<%=Accounts(5, i)%>"><%=Accounts(5, i)%></a>
									<%Case Else%>
										<%=Accounts(j, i)%>
								<%End Select%>
							</td>
						<%Next%>
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
%></body>
</html>
