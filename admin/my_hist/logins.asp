<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2, conn2
Dim sFirstName, sLastName, sScreenName
Dim Logins()
Dim i

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Logins(4, 0)
Set rs = SErver.CreateObject("ADODB.Recordset")
sql = "SELECT MyHistID, WhenVisit, IPAddress, Browser FROM MyHistLogin ORDER BY WhenVisit DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Logins(0, i) = rs(0).Value
    Logins(1, i) = GetUserName(rs(0).Value)
    Logins(2, i) = rs(1).Value
    Logins(3, i) = rs(2).Value
    Logins(4, i) = Left(rs(3).Value, 50)
    i = i + 1
    ReDim Preserve Logins(4, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function GetUserName(lMyHistID)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT pd.FirstName, pd.LastName, pd.ScreenName FROM PartData pd INNER JOIN PartProfile pp "
    sql2 = sql2 & "ON pd.PartID = pp.PartID WHERE MyHistID = " & lMyHistID
	rs2.Open sql2, conn2, 1, 2
    If rs2.RecordCount > 0 Then 
        If Not rs2(1).Value & "" = "" Then sLastName = Replace(rs2(1).Value, "''", "'")
        If Not rs2(0).Value & "" = "" Then sFirstName = Replace(rs2(0).Value, "''", "'")
        If Not rs2(2).Value & "" = "" Then sScreenName = Replace(rs2(2).Value, "''", "'")
        GetUserName = sLastName & ", " & sFirstName & " (" & sScreenName & ")"
    End IF
	rs2.Close
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE My History Account Logins</title>

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
		    <!--#include file = "my_hist_nav.asp" -->

		    <h4 class="h4">My History Account Logins</h4>

            <table style="font-size: 0.85em;">
                <tr>
                    <th style="text-align: right;">No.</th>
                    <th>Account User</th>
                    <th>When Visit</th>
                    <th>IPAddress</th>
                    <th>Browser</th>
                </tr>
                <%For i = 0 To UBound(Logins, 2) - 1%>
                    <tr>
                        <%If i mod 2 = 0 Then%>
                            <td class="alt"><%=i + 1%>)</td>
                            <td class="alt"><%=Logins(1, i)%></td>
                            <td class="alt"><%=Logins(2, i)%></td>
                            <td class="alt"><%=Logins(3, i)%></td>
                            <td class="alt"><%=Logins(4, i)%></td>
                        <%Else%>
                            <td><%=i + 1%>)</td>
                            <td><%=Logins(1, i)%></td>
                            <td><%=Logins(2, i)%></td>
                            <td><%=Logins(3, i)%></td>
                            <td><%=Logins(4, i)%></td>
                        <%End If%>
                    </tr>
                <%Next%>
            </table>
        </div>
	</div>
</div>
<%
conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%></body>
</html>
