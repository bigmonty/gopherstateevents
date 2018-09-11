<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim Staff()
Dim i, j

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(11, 0)
sql = "SELECT StaffID, FirstName, LastName, Address, City, State, Zip, Phone, UserID, Password, Tech, Support, Email FROM Staff "
sql = sql & "ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then Staff(2, i) = Replace(rs(3).Value, "''", "'")
	If Not rs(4).Value & "" = "" Then Staff(3, i) = Replace(rs(4).Value, "''", "'")
	Staff(4, i) = rs(5).Value
	Staff(5, i) = rs(6).Value
	Staff(6, i) = rs(7).Value
	Staff(7, i) = rs(8).Value
	Staff(8, i) = rs(9).Value
	Staff(9, i) = rs(10).Value
	Staff(10, i) = rs(11).Value
	Staff(11, i) = rs(12).Value
	i = i + 1
	ReDim Preserve Staff(11, i)
	rs.MoveNext
Loop
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events Staff</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
 		    <h4 class="h4">Staff Information</h4>
		
		    <table class="table table-striped">
			    <tr>
				    <th>No.</th>
				    <th>Name (Email)</th>
				    <th>Address</th>
				    <th>City</th>
				    <th>ST</th>
				    <th>Zip</th>
				    <th>Phone</th>
				    <th>User ID</th>
				    <th>Password</th>
                    <th>Tech</th>
                    <th>Support</th>
			    </tr>
			    <%For i = 0 to UBound(Staff, 2) - 1%>
				    <tr>
						<td style="text-align:right;">
							<%=i + 1%>)
						</td>
						<%For j = 1 to 10%>
							<td style="white-space:nowrap;">
								<%If j = 1 Then%>
									<a href="mailto:<%=Staff(9, i)%>"><%=Staff(1, i)%></a>
								<%Else%>
									<%=Staff(j, i)%>
								<%End If%>
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
