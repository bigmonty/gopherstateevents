<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim Staff(), MyHours()
Dim i, j
Dim iThisYear

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iThisYear = Request.QueryString("this_year")
If CStr(iThisYear) = vbNullString Then iThisYear = Year(Date)

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(1, 0)
sql = "SELECT StaffID, FirstName, LastName FROM Staff ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	i = i + 1
	ReDim Preserve Staff(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub StaffHours(lStaffID)
    Dim x

    x = 0

    ReDim MyHours(4, 0)
'    Set rs = Server.CreateObject
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Staff Hours</title>

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
			<h4 class="h4">Staff Hours</h4>

            <%If Year(Date) > 2013 Then%>
                <div style="text-align: right;margin: 5px 0 0 0;padding: 0;font-size: 0.9em;">
                    <ul style="display: inline-block;">
                        <%For i = 2002 To Year(Date)%>
                            <li style="display: inline-block;"><a href="staff_hours.asp?this_year=<%=i%>"><%=i%></a>&nbsp;&nbsp;&nbsp;</li>
                        <%Next%>
                    </ul>
                </div>
		    <%End If%>

		    <table style="font-size:0.9em;margin:10px 0 0 0;">
               <%For i = 0 To UBound(Staff, 2) - 1%>
                    <tr>
                        <td style="white-space: nowrap;border: 1px solid #ececd8;padding:5px 10px 5px 10px;" valign="top">
                            <%Call StaffHours(Staff(0, i))%>
                            <h4 style="background-color: #ececd8"><%=Staff(1, i)%></h4>

                            <table>
                                <tr><th>No.</th><th>Date</th><th>Type</th><th>Hours</th><th>Comments</th></tr>
                                <%For j = 0 To UBound(MyHours, 2) - 1%>
                                    <%If j mod 2 = 0 Then%>
                                        <tr>
                                            <td class="alt" valign="top"><%=j + 1%>)</td>
                                            <td class="alt" valign="top"><%=MyHours(1, j)%></td>
                                            <td class="alt" valign="top"><%=MyHours(2, j)%></td>
                                            <td class="alt" valign="top"><%=MyHours(3, j)%></td>
                                            <td class="alt" valign="top"><%=MyHours(4, j)%></td>
                                        </tr>
                                    <%Else%>
                                        <tr>
                                            <td valign="top"><%=j + 1%>)</td>
                                            <td valign="top"><%=MyHours(1, j)%></td>
                                            <td valign="top"><%=MyHours(2, j)%></td>
                                            <td valign="top"><%=MyHours(3, j)%></td>
                                            <td valign="top"><%=MyHours(4, j)%></td>
                                        </tr>
                                    <%End If%>
                                <%Next%>
                            </table>
                        </td>
			        </tr>
                <%Next%>
		    </table>
        </div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%></body>
</html>
