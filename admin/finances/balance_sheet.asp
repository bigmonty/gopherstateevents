<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim iYear, iNumEvents, sngMinEvents, sngEventDeficit
Dim sngInvoices, sngCurrDeposits, sngAdvDeposits, sngPayments, sngEventBalance, sngEventExpenses, sngPastDue, sngDeficit, sngEventProfit
Dim sngOtherExpenses, sngOtherIncome, sngTotalIncome, sngTotalExpenses, sngBottomLine, sngH51Balance, sngH51Expenses, sngH51Payments, sngCapitalOutlay
Dim sngProfitPerEvent, sngH51Payment, sngH51Income
Dim Events(), SortArr(3), Sports(2)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = REquest.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Sports(0) = "Fitness Event"
Sports(1) = "Nordic Ski"
Sports(2) = "Cross-Country"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE (EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "')"
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
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "')"
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

iNumEvents = UBound(Events, 2)

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

'get race balance (total invoice to date - deposits total - race payments
sngInvoices = 0
sngCurrDeposits = 0
sngAdvDeposits = 0
sngPayments = 0
sngEventExpenses = 0
sngOtherIncome = 0
sngOtherExpenses = 0
sngH51Balance = 0
sngH51Payments = 0
sngH51Expenses = 0
sngDeficit = 0
sngEventProfit = 0
sngCapitalOutlay = 0
sngProfitPerEvent = 0

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AmtRcvd FROM FinanceIncome WHERE IncomeType = 'Misc Income' AND WhenRcvd >= '1/1/" & iYear & "' AND WhenRcvd <= '12/31/" & iYear & "'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sngOtherIncome =  CSNg(sngOtherIncome) + CSng(rs(0).Value)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

For i = 0 To UBound(Events, 2) - 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice FROM FinanceEvents WHERE EventID = " & Events(0, i) & " AND Sport = '" & Events(3, i) & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sngInvoices =  CSNg(sngInvoices) + CSng(rs(0).Value)
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtRcvd FROM FinanceIncome WHERE EventID = " & Events(0, i) & " AND Sport = '" & Events(3, i) & "' AND IncomeType = 'Invoice Payment'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sngPayments =  CSNg(sngPayments) + CSng(rs(0).Value)
    rs.Close
    Set rs = Nothing

    'get expenses
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Staffing, MiscCost, PartCost, LaborCost FROM FinanceEvents WHERE EventID = " & Events(0, i)
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sngEventExpenses = CSng(sngEventExpenses) + CSng(rs(0).Value) + CSng(rs(1).Value) + CSng(rs(2).Value) + CSng(rs(3).Value)
    rs.Close
    Set rs = Nothing
Next

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AmtPaid FROM FinanceExpense WHERE (WhenPaid >= '1/1/" & iYear & "' AND WhenPaid <= '12/31/" & iYear & "') AND ExpenseType = 'H51 Software Expense'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sngH51Expenses = CSng(sngH51Expenses)  + CSng(rs(0).Value)
    sngH51Balance = CSng(sngH51Balance) + CSng(rs(0).Value)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AmtRcvd FROM FinanceIncome WHERE (WhenRcvd >= '1/1/" & iYear & "' AND WhenRcvd <= '12/31/" & iYear & "') AND IncomeType = 'H51 Payment'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sngH51Income = CSng(sngH51Income)  + CSng(rs(0).Value)
    sngH51Balance = CSng(sngH51Balance) - CSng(rs(0).Value)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sngH51Income = Round(sngH51Income, 2)
sngH51Balance = Round(sngH51Balance, 2)

'get all deposits for this year's races
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Deposit, EventDate FROM Events WHERE (EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "')"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If CDate(rs(1).Value) > Date Then
        sngAdvDeposits = CSng(sngAdvDeposits) + CSng(rs(0).Value)
    Else
        sngCurrDeposits = CSng(sngCurrDeposits) + CSng(rs(0).Value)
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sngTotalIncome = CSng(sngInvoices) + CSng(sngOtherIncome)
sngPastDue = CSng(sngInvoices) - CSng(sngPayments) - CSng(sngCurrDeposits)
sngEventBalance = CSng(sngCurrDeposits) + CSng(sngAdvDeposits) + CSng(sngPayments) - CSng(sngInvoices)

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AmtPaid FROM FinanceExpense WHERE (WhenPaid >= '1/1/" & iYear & "' AND WhenPaid <= '12/31/" & iYear & "') AND ExpenseType <>'H51 Software Expense'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sngOtherExpenses = CSng(sngOtherExpenses) + CSng(rs(0).Value)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AmtPaid FROM FinanceExpense WHERE (WhenPaid >= '1/1/" & iYear & "' AND WhenPaid <= '12/31/" & iYear & "') AND ExpenseType = 'Capital Outlay'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sngCapitalOutlay = CSng(sngCapitalOutlay) + CSng(rs(0).Value)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sngTotalExpenses = CSng(sngEventExpenses) + CSng(sngOtherExpenses)
sngBottomLine = CSng(sngTotalIncome) - CSng(sngTotalExpenses) + CSng(sngH51Balance)
sngEventProfit = CSng(sngInvoices) - CSng(sngEventExpenses)
If CInt(iNumEvents) > 0 Then sngProfitPerEvent = Round(CSng(sngEventProfit)/CInt(iNumEvents), 2)

sngDeficit = CSng(sngBottomLine) + CSng(sngCapitalOutlay)
sngDeficit = CSng(sngDeficit)

If CSng(sngProfitPerEvent) > 0 Then sngEventDeficit =  Round(-CSng(sngDeficit)/CSng(sngProfitPerEvent), 2)
sngMinEvents = CSng(sngEventDeficit) + CInt(iNumEvents)
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Balance Sheet</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td, th{
        padding: 5px;
        border: 1px solid #ececec;
    }
    
    ul{
        list-style-type: none;
    }
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">Balance Sheet To Date</h4>
            <div style="margin: 0;padding: 0 10px 0 0;text-align: right;">
                <%For i = 2015 To Year(Date) + 1%>
                    <a href="balance_sheet.asp?year=<%=i%>" style="font-size: 0.85em;"><%=i%></a>
                    <%If Not i = Year(Date) + 1 Then%>
                        &nbsp;|&nbsp;
                    <%End If%>
                <%Next%>
           </div>

            <div style="float: left;width: 300px;">
                <h4 style="background: none;margin-bottom: 0;padding-bottom: 0;">Bottom Line</h4>
                <ul style="margin-top: 0;padding-top: 0;">
                    <li>Total Income:&nbsp;$<%=sngTotalIncome%></li>
                    <li>Total Expenses:&nbsp;$<%=sngTotalExpenses%></li>
                    <li>H51 Balance (owed to GSE):&nbsp;$<%=sngH51Balance%></li>
                    <%If CSNg(sngBottomLine) < 0 Then%>
                        <li style="font-weight: bold;color: red;">Bottom Line:&nbsp;$<%=sngBottomLine%></li>
                    <%Else%>
                        <li style="font-weight: bold;">Bottom Line:&nbsp;$<%=sngBottomLine%></li>
                    <%End If%>
                </ul>

                <h4 style="background: none;margin-bottom: 0;padding-bottom: 0;">Income</h4>
                <ul style="margin-top: 0;padding-top: 0;">
                    <li>Invoice Total:&nbsp;$<%=sngInvoices%></li>
                    <li>Other Income:&nbsp;$<%=sngOtherIncome%></li>
                    <li style="font-weight: bold;">Total Income:&nbsp;$<%=sngTotalIncome%></li>
                </ul>

                <h4 style="background: none;margin-bottom: 0;padding-bottom: 0;">Expenses</h4>
                <ul style="margin-top: 0;padding-top: 0;">
                    <li>Event Expenses:&nbsp;$<%=sngEventExpenses%></li>
                    <li>Other Expenses:&nbsp;$<%=sngOtherExpenses%></li>
                    <li style="font-weight: bold;">Total Expenses:&nbsp;$<%=sngTotalExpenses%></li>
                </ul>

                <h4 style="background: none;margin-bottom: 0;padding-bottom: 0;">Events Balance</h4>
                <ul style="margin-top: 0;padding-top: 0;">
                    <li>Invoice Total:&nbsp;$<%=sngInvoices%></li>
                    <li>Current Deposits:&nbsp;$<%=sngCurrDeposits%></li>
                    <li>Advanced Deposits:&nbsp;$<%=sngAdvDeposits%></li>
                    <li>Payments Rcvd:&nbsp;$<%=sngPayments%></li>
                    <li style="color: red;">Payments Outstanding:&nbsp;$<%=sngPastDue%></li>
                    <%If CSNg(sngEventBalance) < 0 Then%>
                        <li style="font-weight: bold;color: red;">Events Balance:&nbsp;$<%=sngEventBalance%></li>
                    <%Else%>
                        <li style="font-weight: bold;">Events Balance:&nbsp;$<%=sngEventBalance%></li>
                    <%End If%>
                </ul>
            </div>
            <div style="margin-left:310px;">
                <h4 style="background: none;margin-bottom: 0;padding-bottom: 0;">Deficit Recovery-Events Needed</h4>
                <ul style="margin-top: 0;padding-top: 0;">
                    <%If CSNg(sngDeficit) < 0 Then%>
                        <li style="color: red;">Deficit (bottom line minus capital outlay):&nbsp;$<%=sngDeficit%></li>
                    <%Else%>
                        <li>Balance (bottom line minus capital outlay):&nbsp;$<%=sngDeficit%></li>
                    <%End If%>
                    <li>Num Events:&nbsp;<%=iNumEvents%></li>
                    <li>Event Profits:&nbsp;$<%=sngEventProfit%></li>
                    <li>Profit per Event:$<%=sngProfitPerEvent%></li>
                    <li>Minimum Events:&nbsp;<%=sngMinEvents%></li>
                    <li>Event Deficit:&nbsp;<%=sngEventDeficit%></li>
                </ul>
            </div>
        </div>
	</div>
</div>
<%	
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
