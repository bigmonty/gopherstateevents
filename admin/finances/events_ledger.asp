<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim iYear
Dim sngInvoice, sngInvoiceTotal, sngDeposit, sngDepositTotal, sngPayments, sngPaymentsTotal, sngBalance, sngBalanceTotal
Dim Events(), SortArr(3)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = REquest.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '1/1/" & iYear & "' AND EventDate <= '" & Date & "' ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'")
    Events(2, i) = rs(2).Value
    Events(3, i) = "Fitness Event"
    i = i + 1
    ReDim Preserve Events(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '" & Date & "' ORDER BY MeetDate"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'")
    Events(2, i) = rs(2).Value
    Events(3, i) = rs(3).Value
    i = i + 1
    ReDim Preserve Events(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'order by date
For i = 0 To UBound(Events, 2) - 2
    For j = i + 1 To UBound(Events, 2) - 1
        If CDate(Events(2, i)) > CDate(Events(2, j)) Then
            For k = 0 To 3
                SortArr(k) = Events(k, i)
                Events(k, i) = Events(k, j)
                Events(k, j) = SortArr(k)
            Next
        End If
    Next
Next

sngInvoiceTotal = 0
sngDepositTotal = 0
sngPaymentsTotal = 0
sngBalanceTotal = 0

Private Sub EventData(lThisEvent, sThisSport)
    sngInvoice = 0
    sngDeposit = 0
    sngPayments = 0
    sngBalance = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice FROM FinanceEvents WHERE EventID = " & lThisEvent & " AND Sport = '" & sThisSport & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sngInvoice = rs(0).Value
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtRcvd FROM FinanceIncome WHERE EventID = " & lThisEvent & " AND Sport = '" & sThisSport & "' AND IncomeType = 'Invoice Payment'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngPayments =  CSng(sngPayments) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If sThisSport = "Fitness Event" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AmtRcvd FROM FinanceIncome WHERE EventID = " & lThisEvent & " AND Sport = 'Fitness Event' AND IncomeType = 'Race Deposit'"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then sngDeposit = CSng(rs(0).Value)
        rs.Close
        Set rs = Nothing
    End If

    sngBalance = CSng(sngDeposit) + CSng(sngPayments) - CSng(sngInvoice)
    sngInvoiceTotal = CSng(sngInvoiceTotal) + CSng(sngInvoice)
    sngDepositTotal = CSng(sngDepositTotal) + CSng(sngDeposit)
    sngPaymentsTotal = CSng(sngPaymentsTotal) + CSng(sngPayments)
    sngBalanceTotal = CSng(sngBalanceTotal) + CSng(sngBalance)
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Events Ledger</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <!--#include file = "events_nav.asp" -->

		    <h3 class="h3">GSE Finances: Events Ledger</h3>

            <ul class="nav">
                <%For i = 2015 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="events_ledger.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
           </ul>

            <table class="table table-striped">
                <tr>
                    <th>No.</th>
                    <th>Event/Meet (Date)</th>
                    <th>Invoice</th>
                    <th>Deposit</th>
                    <th>Payments</th>
                    <th>Balance</th>
                </tr>
                <%For j = 0 To UBound(Events, 2) - 1%>
                    <%Call EventData(Events(0, j), Events(3, j))%>
                    <tr>
                        <td><%=j + 1%></td>
                        <td><%=Events(1, j)%> (<%=Events(2, j)%>)</td>
                        <td>$<%=sngInvoice%></td>
                        <td>$<%=sngDeposit%></td>
                        <td>$<%=sngPayments%></td>
                        <%If CSng(sngBalance) < 0 Then%>
                            <td style="color:red;">$<%=sngBalance%></td>
                        <%Else%>
                            <td>$<%=sngBalance%></td>
                        <%End If%>
                    </tr>
                <%Next%>
                <tr>
                    <th colspan="2">Column Totals</th>
                    <th>$<%=sngInvoiceTotal%></th>
                    <th>$<%=sngDepositTotal%></th>
                    <th>$<%=sngPaymentsTotal%></th>
                    <%If CSng(sngBalanceTotal) < 0 Then%>
                        <th style="color:red;">$<%=sngBalanceTotal%></th>
                    <%Else%>
                        <th>$<%=sngBalanceTotal%></th>
                    <%End If%>
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
