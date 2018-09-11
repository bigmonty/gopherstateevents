<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim Followers()

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get participating teams info	
i = 0
ReDim Followers(8, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FollowersID, FirstName, LastName, Email, UserName, Password, MobilePhone, Provider, WhenReg, Status FROM Followers "
sql = sql & "ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Followers(0, i) = rs(0).Value
	Followers(1, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
    Followers(2, i) = rs(3).Value
    Followers(3, i) = rs(4).Value
    Followers(4, i) = rs(5).Value
    Followers(5, i) = rs(6).Value
    Followers(6, i) = rs(7).Value
    Followers(7, i) = rs(8).Value
    Followers(8, i) = rs(9).Value
	i = i + 1
	ReDim Preserve Followers(8, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Followers</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    th{
        padding-right: 5px;
    }
</style>
</head>
<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE Followers</h4>

			<table style="font-size: 1.0em;">
                <tr>
                    <th>No.</th>
                    <th>Follower</th>
                    <th>Email</th>
                    <th>User Name</th>
                    <th>Password</th>
                    <th>Mobile Phone</th>
                    <th>Provider</th>
                    <th>When Reg</th>
                    <th>Status</th>
                </tr>
                <%For i = 0 To UBound(Followers, 2) - 1%>
                    <tr>
                        <%If i mod 2 = 0 Then%>
                            <td class="alt" style="text-align: right;"><%=i + 1%>)</td>
                            <td class="alt"><%=Followers(1, i)%></td>
                            <td class="alt"><a href="mailto:<%=Followers(2, i)%>"><%=Followers(2, i)%></a></td>
                            <td class="alt"><%=Followers(3, i)%></td>
                            <td class="alt" ><%=Followers(4, i)%></td>
                            <td class="alt"><%=Followers(5, i)%></td>
                            <td class="alt" ><%=Followers(6, i)%></td>
                            <td class="alt"><%=Followers(7, i)%></td>
                            <td class="alt" ><%=Followers(8, i)%></td>
                        <%Else%>
                            <td style="text-align: right;"><%=i + 1%>)</td>
                            <td><%=Followers(1, i)%></td>
                            <td><a href="mailto:<%=Followers(2, i)%>"><%=Followers(2, i)%></a></td>
                            <td><%=Followers(3, i)%></td>
                            <td><%=Followers(4, i)%></td>
                            <td><%=Followers(5, i)%></td>
                            <td><%=Followers(6, i)%></td>
                            <td><%=Followers(7, i)%></td>
                            <td><%=Followers(8, i)%></td>
                        <%End If%>
                    </tr>
                <%Next%>
            </table>
		</div>
	</div>
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
