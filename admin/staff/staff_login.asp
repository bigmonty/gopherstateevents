<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim Staff(), Logins()
Dim i, j

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(3, 0)
sql = "SELECT StaffID, FirstName, LastName, Phone, Email FROM Staff WHERE Active  = 'y' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	Staff(2, i) = rs(3).Value
	Staff(3, i) = rs(4).Value
	i = i + 1
	ReDim Preserve Staff(3, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub StaffLogins(lStaffID)
    Dim x

    x = 0
    ReDim Logins(0)
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT WhenVisit FROM StaffLogin WHERE StaffID = " & lStaffID & " ORDER BY WhenVisit DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Logins(x) = rs(0).Value
        x = x + 1
        ReDim Preserve Logins(x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Staff Logins</title>

<!--#include file = "../../includes/js.asp" -->


<style type="text/css">
    td,th{padding-right: 5px;}
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">Staff Logins</h4>
		
		    <table style="font-size:0.85em;margin:10px 0 0 0;">
			    <tr>
                    <%For i = 0 To UBound(Staff, 2) -1%>
                        <%Call StaffLogins(Staff(0, i))%>
                        <td style="white-space: nowrap;border: 1px solid #ececd8;padding:5px 10px 5px 10px;" valign="top">
                            <h4 class="h4"><%=Staff(1, i)%></h4>

                            <ul>
                                <%For j = 0 To UBound(Logins) - 1%>
                                    <li><%=Logins(j)%></li>
                                <%Next%>
                            </ul>
                        </td>
                    <%Next%>
			    </tr>
		    </table>
        </div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%></body>
</html>
