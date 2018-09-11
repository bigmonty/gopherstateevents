<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim iYear
Dim sngMyAmt, sngStaffTtl, sngTtlEarned, sngTtlPd
Dim Staff(), Events(), StaffView(), SortArr(4)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = REquest.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

sngStaffTtl = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(4, 0)
sql = "SELECT StaffID, FirstName, LastName FROM Staff WHERE Active = 'y' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    Staff(2, i) = "0"
    Staff(3, i) = "0"
    Staff(4, i) = "0"
	i = i + 1
	ReDim Preserve Staff(4, i)
	rs.MoveNext
Loop
Set rs = Nothing

sngTtlEarned = 0
sngTtlPd = 0

For i = 0 To UBound(Staff, 2) - 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TransAmt, TransType FROM FinanceStaff WHERE StaffID = " & Staff(0, i) 
    '& " AND TransDate >= '1/1/"
    'sql = sql & iYear & "' AND TransDate <= '12/31/" & iYear & "'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If rs(1).Value = "Payment" Then
            Staff(3, i) = CSng(Staff(3, i)) + CSng(rs(0).Value)
            sngTtlPd = CSng(sngTtlPd)  + CSng(rs(0).Value)
        Else
            Staff(2, i) = CSng(Staff(2, i)) + CSng(rs(0).Value)
            sngTtlEarned = CSng(sngTtlEarned)  + CSng(rs(0).Value)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Staff(4, i) = CSng(Staff(2, i)) - CSng(Staff(3, i))
Next
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Staff Ledger</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <!--#include file = "staff_nav.asp" -->

		    <h3 class="h3">GSE Finances: Staff Ledger</h3>

            <ul class="nav">
                <%For i = 2015 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="staff_ledger.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
           </ul>

            <table class="table table-striped">
                <tr>
                    <th>No.</th>
                    <th>Name</th>
                    <th>Earned</th>
                    <th>Paid</th>
                    <th>Due</th>
                </tr>
                <%For i = 0 To UBound(Staff, 2) - 1%>
                    <tr>
                        <td><%=i + 1%>)</td>
                        <td><%=Staff(1, i)%></td>
                        <td>$<%=Staff(2, i)%></td>
                        <td>$<%=Staff(3, i)%></td>
                        <td>$<%=Round(Staff(4, i), 2)%></td>
                    </tr>
                <%Next%>
                <tr>
                    <th colspan="2">Summary:</th>
                    <th>$<%=Round(sngTtlEarned, 2)%></th>
                    <th>$<%=Round(sngTtlPd, 2)%></th>
                    <th>$<%=Round(CSng(sngTtlEarned) - CSng(sngTtlPd), 2)%></th>
                </tr>
            </table>
        </div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%	
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
