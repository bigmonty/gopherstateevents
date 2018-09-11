<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim lEventID
Dim i, j, k
Dim sWhich, sExpenseType, sComments, sFromTo, sIncomeType, sSport, sShowWhat
Dim iYear
Dim sngAmt, sngTotalIncome, sngTotalExpense, sngEventExpense, sngIncomeExpenseTotal, sngBalance, sngIncomeTypeTotal, sngExpenseTypeTotal
Dim Income(), Expenses(), IncomeTypes(10), Sports(2), Events(), SortArr(), ExpenseTypes(14), Chronology(), IncomeByType(), ExpenseByType()
Dim dWhen

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

sShowWhat = Request.QueryString("show_what")
If CStr(sShowWhat) = vbNullString Then sShowWhat = "both"

sWhich = Request.QueryString("which")
If CStr(sWhich) = vbNullString Then sWhich = "breakdown"

Sports(0) = "Fitness Event"
Sports(1) = "Nordic Ski"
Sports(2) = "Cross-Country"

ExpenseTypes(0) = "Bank Charges"
ExpenseTypes(1) = "Bob Bakken Draw"
ExpenseTypes(2) = "Bob Schneider Draw"
ExpenseTypes(3) = "Capital Outlay"
ExpenseTypes(4) = "Computer & Internet Expenses"
ExpenseTypes(5) = "H51 Software Expense"
ExpenseTypes(6) = "Insurance"
ExpenseTypes(7) = "Loan Repay"
ExpenseTypes(8) = "Materials & Equipment"
ExpenseTypes(9) = "Meals & Entertainment"
ExpenseTypes(10) = "Office Supplies"
ExpenseTypes(11) = "Professional Fees"
ExpenseTypes(12) = "Rent Expense"
ExpenseTypes(13) = "Travel & Lodging"
ExpenseTypes(14) = "Misc Expenses"

IncomeTypes(0) = "Advertising"
IncomeTypes(1) = "Credit for Return"
IncomeTypes(2) = "H51 Payment"
IncomeTypes(3) = "Investment Capital"
IncomeTypes(4) = "Invoice Payment"
IncomeTypes(5) = "Loan Proceeds"
IncomeTypes(6) = "Race Deposit"
IncomeTypes(7) = "Misc Income"
IncomeTypes(8) = "Tempo Events"
IncomeTypes(9) = "AdSense"
IncomeTypes(10) = "Crowd Torch"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE (EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "') "
sql = sql & "ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
    Events(2, i) = rs(2).Value
    Events(3, i) = "Fitness Event"
    i = i + 1
    ReDim Preserve Events(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "') "
sql = sql & "ORDER BY MeetDate"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
    Events(2, i) = rs(2).Value
    Events(3, i) = rs(3).Value
    i = i + 1
    ReDim Preserve Events(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'order by date
ReDim SortArr(2)
For i = 0 To UBound(Events, 2) - 2
    For j = i + 1 To UBound(Events, 2) - 1
        If CDate(Events(2, i)) > CDate(Events(2, j)) Then
            For k = 0 To 2
                SortArr(k) = Events(k, i)
                Events(k, i) = Events(k, j)
                Events(k, j) = SortArr(k)
            Next
        End If
    Next
Next

If Request.Form.Item("submit_expense") = "submit_expense" Then
    sngAmt = Request.Form.Item("amt_paid")
    dWhen = Request.Form.Item("when_paid")
    sFromTo = Request.Form.Item("paid_to")
    If Not Request.Form.Item("expense_type") = vbNullString Then sExpenseType = Replace(Request.Form.Item("expense_type"), "'", "''")
    If Not Request.Form.Item("comments") = vbNullString Then sComments = Replace(Request.Form.Item("comments"), "'", "''")
   
    sql = "INSERT INTO FinanceExpense (AmtPaid, WhenPaid, PaidTo, ExpenseType, Comments) VALUES (" & sngAmt & ", '" & dWhen& "', '" & sFromTo
    sql = sql & "', '" & sExpenseType & "', '" & sComments & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_income") = "submit_income" Then
    sngAmt = Request.Form.Item("amt_rcvd")
    dWhen = Request.Form.Item("when_rcvd")
    sFromTo = Request.Form.Item("rcvd_from")
    sIncomeType = Request.Form.Item("income_type")
    lEventID = Request.Form.Item("event_id")
    sSport = Request.Form.Item("sport")
    If Not Request.Form.Item("comments") = vbNullString Then sComments = Replace(Request.Form.Item("comments"), "'", "''")
   
    sql = "INSERT INTO FinanceIncome (AmtRcvd, WhenRcvd, RcvdFrom, IncomeType, EventID, Sport, Comments) VALUES (" & sngAmt & ", '" & dWhen & "', '" 
    sql = sql & sFromTo & "', '" & sIncomeType & "', '" & lEventID & "', '" & sSport & "', '" & sComments & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If

sngTotalIncome = 0
sngTotalExpense = 0

ReDim Income(7, 0)
ReDim Expenses(5, 0)
ReDim Chronology(9, 0)
ReDim IncomeByType(5, 0)
ReDim ExpenseByType(3, 0)

sngEventExpense = 0
For i = 0 To UBound(Events, 2) - 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Staffing, MiscCost, PartCost, LaborCost FROM FinanceEvents WHERE EventID = " & Events(0, i) & " AND Sport = '" & Events(3, i) & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sngEventExpense = CSng(sngEventExpense) + CSng(rs(0).Value) +  CSng(rs(1).Value) +  CSng(rs(2).Value) +  CSng(rs(3).Value) 
    End If
    rs.Close
    Set rs = Nothing
Next

Select Case sWhich 
    Case "breakdown"
        If sShowWhat = "income" Then
            Call IncomeHist()
        ElseIf sShowWhat = "expense" Then
            Call ExpenseHist()
        Else
            Call ExpenseHist()
            Call IncomeHist()
        End If
    Case "chronological"
        i = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT FinanceExpenseID, WhenPaid, AmtPaid, PaidTo, ExpenseType, Comments FROM FinanceExpense WHERE (WhenPaid >= '1/1/" & iYear 
        sql = sql & "' AND WhenPaid <= '12/31/" & iYear & "') ORDER BY WhenPaid"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            Chronology(0, i) = rs(0).Value  'id
            Chronology(1, i) = rs(1).Value  'date
            Chronology(2, i) = rs(2).Value  'amount
            Chronology(3, i) = "Expense"    'income/expense
            Chronology(4, i) = rs(3).Value  'to/from
            Chronology(5, i) = rs(4).Value  'type
            Chronology(6, i) = "n/a"        'event
            Chronology(7, i) = "n/a"        'sport
            If Not rs(5) & "" = "" Then Chronology(8, i) = Replace(rs(5).Value, "''", "'")
            Chronology(9, i) = "0"          'balance

            sngTotalExpense = CSng(sngTotalExpense) + CSng(rs(2).Value)

            i = i + 1
            ReDim Preserve Chronology(9, i)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT FinanceIncomeID, WhenRcvd, AmtRcvd, RcvdFrom, IncomeType, EventID, Sport, Comments FROM FinanceIncome WHERE (WhenRcvd >= '1/1/" & iYear 
        sql = sql & "' ANd WhenRcvd <= '12/31/" & iYear & "') ORDER BY WhenRcvd"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            Chronology(0, i) = rs(0).Value  'id
            Chronology(1, i) = rs(1).Value  'date
            Chronology(2, i) = rs(2).Value  'amount
            Chronology(3, i) = "Income"    'income/expense
            Chronology(4, i) = rs(3).Value  'to/from
            Chronology(5, i) = rs(4).Value  'type
            Chronology(6, i) = GetEventName(rs(5).Value, rs(6).Value)        'event
            Chronology(7, i) = rs(6).Value        'sport
            If Not rs(7) & "" = "" Then Chronology(8, i) = Replace(rs(7).Value, "''", "'")
            Chronology(9, i) = "0"          'balance

            sngTotalIncome = CSng(sngTotalIncome) + CSng(rs(2).Value)

            i = i + 1
            ReDim Preserve Chronology(9, i)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        'order by date
        ReDim SortArr(9)
        For i = 0 To UBound(Chronology, 2) - 2
            For j = i + 1 To UBound(Chronology, 2) - 1
                If CDate(Chronology(1, i)) > CDate(Chronology(1, j)) Then
                    For k = 0 To 9
                        SortArr(k) = Chronology(k, i)
                        Chronology(k, i) = Chronology(k, j)
                        Chronology(k, j) = SortArr(k)
                    Next
                End If
            Next
        Next

        For i = 0 To UBound(Chronology, 2) - 1
            If i = 0 Then
                If Chronology(3, i) = "Income" Then
                    Chronology(9, i) = Chronology(2, i)
                Else
                    Chronology(9, i) = - CSng(Chronology(2, i))
                End If
            Else
                If Chronology(3, i) = "Income" Then
                    Chronology(9, i) = CSng(Chronology(9, i - 1)) + CSng(Chronology(2, i))
                Else
                    Chronology(9, i) = CSng(Chronology(9, i - 1)) - CSng(Chronology(2, i))
                End If
            End If
        Next
    Case "income_category"
        'get total income
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AmtRcvd FROM FinanceIncome WHERE (WhenRcvd >= '1/1/" & iYear & "' AND WhenRcvd <= '12/31/" & iYear & "') ORDER BY WhenRcvd"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            sngTotalIncome = CSng(sngTotalIncome) + CSng(rs(0).Value)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Case "expense_category"
        'get total expenses
        If Not sShowWhat = "both" Then
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT AmtPaid FROM FinanceExpense WHERE (WhenPaid >= '1/1/" & iYear & "' AND WhenPaid <= '12/31/" & iYear & "') ORDER BY WhenPaid"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                sngTotalExpense = CSng(sngTotalExpense) + CSng(rs(0).Value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
End Select

Private Sub IncomeHist()
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceIncomeID, WhenRcvd, AmtRcvd, RcvdFrom, IncomeType, EventID, Sport, Comments FROM FinanceIncome WHERE (WhenRcvd >= '1/1/" & iYear 
    sql = sql & "' AND WhenRcvd <= '12/31/" & iYear & "') ORDER BY WhenRcvd DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Income(0, i) = rs(0).Value
        Income(1, i) = rs(1).Value
        Income(2, i) = rs(2).Value
        Income(3, i) = rs(3).Value
        Income(4, i) = rs(4).Value
        If Not rs(5).Value & "" = "" Then Income(5, i) = GetEventName(rs(5).Value, rs(6).Value)
        Income(6, i) = rs(6).Value
        If Not rs(7) & "" = "" Then Income(7, i) = Replace(rs(7).Value, "''", "'")

        If Not sShowWhat = "both" Then sngTotalIncome = CSng(sngTotalIncome) + CSng(rs(2).Value)

        i = i + 1
        ReDim Preserve Income(7, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If sWhich = "income_category" Then
        'order by date
        ReDim SortArr(7)
        For i = 0 To UBound(Income, 2) - 2
            For j = i + 1 To UBound(Income, 2) - 1
                If CDate(Income(1, i)) > CDate(Income(1, j)) Then
                    For k = 0 To 7
                        SortArr(k) = Income(k, i)
                        Income(k, i) = Income(k, j)
                        Income(k, j) = SortArr(k)
                    Next
                End If
            Next
        Next
    End If

    'get total expenses
    If Not sShowWhat = "both" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AmtPaid FROM FinanceExpense WHERE WhenPaid >= '1/1/" & iYear & "' AND WhenPaid <= '" & Date & "'"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            sngTotalExpense = CSng(sngTotalExpense) + CSng(rs(0).Value)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub ExpenseHist()
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceExpenseID, AmtPaid, WhenPaid, PaidTo, ExpenseType, Comments FROM FinanceExpense WHERE (WhenPaid >= '1/1/" & iYear 
    sql = sql & "' AND WhenPaid <= '12/31/" & iYear & "') ORDER BY WhenPaid DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Expenses(0, i) = rs(0).Value
        Expenses(1, i) = rs(1).Value
        Expenses(2, i) = rs(2).Value
        Expenses(3, i) = rs(3).Value
        Expenses(4, i) = Replace(rs(4).Value, "''", "'")
        If Not rs(5) & "" = "" Then Expenses(5, i) = Replace(rs(5).Value, "''", "'")

        sngTotalExpense = CSng(sngTotalExpense) + CSng(rs(1).Value)

        i = i + 1
        ReDim Preserve Expenses(5, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If sWhich = "expense_category" Then
        'order by date
        ReDim SortArr(5)
        For i = 0 To UBound(Expenses, 2) - 2
            For j = i + 1 To UBound(Expenses, 2) - 1
                If CDate(Expenses(2, i)) > CDate(Expenses(2, j)) Then
                    For k = 0 To 5
                        SortArr(k) = Expenses(k, i)
                        Expenses(k, i) = Expenses(k, j)
                        Expenses(k, j) = SortArr(k)
                    Next
                End If
            Next
        Next
    End If

    'get total income
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtRcvd FROM FinanceIncome WHERE WhenRcvd >= '1/1/" & iYear & "' AND WhenRcvd <= '12/31/" & iYear & "'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngTotalIncome = CSng(sngTotalIncome) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Function GetEventName(lThisEvent, sThisSport)
    GetEventName = "unknown"

    Set rs2 = Server.CreateObject("ADODB.Recordset")

    If sThisSport = "Fitness Event" Then
        sql2 = "SELECT EventName FROM Events WHERE EventID = " & lThisEvent
        rs2.Open sql2, conn, 1, 2
    Else
        sql2 = "SELECT MeetName FROM Meets WHERE MeetsID = " & lThisEvent
        rs2.Open sql2, conn2, 1, 2
    End If

    If rs2.RecordCount > 0 Then GetEventName = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

Private Sub IncomeThisType(sThisType)
    sngIncomeTypeTotal = 0

    i = 0
    ReDim IncomeByType(5, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceIncomeID, WhenRcvd, AmtRcvd, RcvdFrom, EventID, Sport FROM FinanceIncome WHERE (WhenRcvd >= '1/1/" & iYear 
    sql = sql & "' AND WhenRcvd <= '12/31/" & iYear & "') AND IncomeType = '" & sThisType & "' ORDER BY WhenRcvd DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        IncomeByType(0, i) = rs(0).Value
        IncomeByType(1, i) = rs(1).Value
        IncomeByType(2, i) = rs(2).Value
        IncomeByType(3, i) = rs(3).Value
        If Not rs(4).Value & "" = "" Then IncomeByType(4, i) = GetEventName(rs(4).Value, rs(5).Value)
        IncomeByType(5, i) = rs(5).Value

        sngIncomeTypeTotal = CSng(sngIncomeTypeTotal) + CSng(rs(2).Value)

        i = i + 1
        ReDim Preserve IncomeByType(5, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub ExpenseThisType(sThisType)
    sngExpenseTypeTotal = 0

    i = 0
    ReDim ExpenseByType(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceExpenseID, WhenPaid, AmtPaid, PaidTo FROM FinanceExpense WHERE (WhenPaid >= '1/1/" & iYear 
    sql = sql & "' AND WhenPaid <= '12/31/" & iYear & "') AND ExpenseType = '" & sThisType & "' ORDER BY WhenPaid DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        ExpenseByType(0, i) = rs(0).Value
        ExpenseByType(1, i) = rs(1).Value
        ExpenseByType(2, i) = rs(2).Value
        ExpenseByType(3, i) = rs(3).Value

        sngExpenseTypeTotal = CSng(sngExpenseTypeTotal) + CSng(rs(2).Value)

        i = i + 1
        ReDim Preserve ExpenseByType(3, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

sngIncomeExpenseTotal = Round(CSng(sngTotalIncome) - CSng(sngTotalExpense), 2)
sngBalance = Round(CSng(sngIncomeExpenseTotal) - CSng(sngEventExpense), 2)
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Income & Expenses</title>

<script>
$(function() {
    $( "#when_paid" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#when_rcvd" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <ul class="nav">
                <li class="nav-item"><a class="nav-link" href="income_expenses.asp?which=breakdown&amp;show_what=<%=sShowWhat%>&amp;year=<%=iYear%>">Income/Expense Breakdown</a></li>
                <li class="nav-item"><a class="nav-link" href="income_expenses.asp?which=income_category&amp;show_what=<%=sShowWhat%>&amp;year=<%=iYear%>">Income By Category</a></li>
                <li class="nav-item"><a class="nav-link" href="income_expenses.asp?which=expense_category&amp;show_what=<%=sShowWhat%>&amp;year=<%=iYear%>">Expenses By Category</a></li>
                <li class="nav-item"><a class="nav-link" href="income_expenses.asp?which=chronological&amp;show_what=<%=sShowWhat%>&amp;year=<%=iYear%>">Chronological List</a></li>
            </ul>
            <ul class="nav">
                <%For i = 2015 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="income_expenses.asp?year=<%=i%>&amp;which=<%=sWhich%>&amp;show_what=<%=sShowWhat%>"><%=i%></a></li>
                    <%If Not i = Year(Date) + 1 Then%>
                    <%End If%>
                <%Next%>
            </ul>

            <h4 class="h4">Current +/-</h4>
            <ul class="nav">
                <li class="nav-item">
                    Income & Expense Total:&nbsp;
                    <%If CSng(sngIncomeExpenseTotal) < 0 Then%>
                        <span style="color: red;">$<%=sngIncomeExpenseTotal%></span>
                    <%Else%>
                        $<%=sngIncomeExpenseTotal%>
                    <%End If%>
                    &nbsp;&nbsp;
                </li>
                <li class="nav-item">
                    Event Expense Total:&nbsp;$<%=sngEventExpense%>&nbsp;&nbsp
                </li>
                <li class="nav-item">
                    Current Balance:
                    <%If CSng(sngBalance) < 0 Then%>
                        <span style="color: red;">$<%=sngBalance%></span>
                    <%Else%>
                        $<%=sngBalance%>
                    <%End If%>
                    &nbsp;&nbsp
                </li>
            </ul>

            <%Select Case sWhich%>
                <%Case "breakdown"%>
		            <h4 class="h4">Payments & Expenses: Breakdown Format</h4>

                    <div class="bg-info">
                        <h5 class="h5">Enter Expense</h5>
                        <form role="form" class="form" name="new_expense" method="post" action="income_expenses.asp?show_what=<%=sShowWhat%>&amp;year=<%=iYear%>">
                        <div class="form-group row">
                            <label class="control-label col-sm-2" for="amt_paid">Amount:</label>
                            <div class="col-sm-2">
                                <input type="text" class="form-control" name="amt_paid" id="amt_paid">
                            </div>
                            <label class="control-label col-sm-2" for="when_paid">When Paid:</label>
                            <div class="col-sm-2">
                                <input type="text" class="form-control" name="when_paid" id="when_paid" autocomplete="off">
                            </div>
                            <label class="control-label col-sm-2" for="paid_to">Paid To:</label>
                            <div class="col-sm-2">
                                <input type="text" class="form-control" name="paid_to" id="paid_to">
                            </div>
                        </div>
                        <div class="form-group row">
                            <label class="control-label col-sm-2" for="expense_type">Type:</label>
                            <div class="col-sm-4">
                                <select class="form-control" name="expense_type" id="expense_type">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(ExpenseTypes)%>
                                        <option value="<%=ExpenseTypes(i)%>"><%=ExpenseTypes(i)%></option>
                                    <%Next%>
                                </select>
                            </div>
                            <label class="control-label col-sm-2" for="comments">Comments:</label>
                            <div class="col-sm-4">
                                <input type="text" class="form-control" name="comments" id="comments">
                            </div>
                        </div>
                        <input type="hidden" name="submit_expense" id="submit_expense" value="submit_expense">
                        <input type="submit" class="form-control" name="submit1" id="submit1" value="Submit Expense">
                        </form>
                    </div>

                    <div class="bg-warning">
                        <h5 class="h5">Enter Income</h5>
                        <form role="form" class="form" name="new_expense" method="post" action="income_expenses.asp?show_what=<%=sShowWhat%>&amp;year=<%=iYear%>">
                        <div class="form-group row">
                            <label class="control-label col-sm-2" for="amt_rcvd">Amount:</label>
                            <div class="col-sm-2">
                                <input type="text" class="form-control" name="amt_rcvd" id="amt_rcvd">
                            </div>
                            <label class="control-label col-sm-2" for="when_rcvd">When Rcvd:</label>
                            <div class="col-sm-2">
                                <input type="text" class="form-control" name="when_rcvd" id="when_rcvd" autocomplete="off">
                            </div>
                            <label class="control-label col-sm-2" for="rcvd_from">Rcvd From:</label>
                            <div class="col-sm-2">
                                <input type="text" class="form-control" name="rcvd_from" id="rcvd_from">
                            </div>
                        </div>
                        <div class="form-group row">
                            <label class="control-label col-sm-2" for="income_type">Type:</label>
                            <div class="col-sm-2">
                                <select class="form-control" name="income_type" id="income_type">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(IncomeTypes)%>
                                        <option value="<%=IncomeTypes(i)%>"><%=IncomeTypes(i)%></option>
                                    <%Next%>
                                </select>
                            </div>
                            <label class="control-label col-sm-2" for="event_id">Event:</label>
                            <div class="col-sm-2">
                                <select class="form-control" name="event_id" id="event_id">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(Events, 2) - 1%>
                                        <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                                    <%Next%>
                                </select>
                            </div>
                            <label class="control-label col-sm-2" for="sport">Sport:</label>
                            <div class="col-sm-2">
                                <select class="form-control" name="sport" id="sport">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(Sports)%>
                                        <option value="<%=Sports(i)%>"><%=Sports(i)%></option>
                                    <%Next%>
                                </select>
                            </div>
                        </div>
                        <div class="form-group row">
                            <label class="control-label col-sm-2" for="comments">Comments:</label>
                            <div class="col-sm-10">
                                <input type="text" class="form-control" name="comments" id="comments">
                            </div>
                        </div>
                        <input type="hidden" name="submit_income" id="submit_income" value="submit_income">
                        <input type="submit" class="form-control" name="submit2" id="submit2" value="Submit Income">
                        </form>
                    </div>

                    <ul class="nav">
                        <li class="nav-item"><a class="nav-link" href="income_expenses.asp?which=breakdown&amp;show_what=income&amp;year=<%=iYear%>">Show Income History</a></li>
                        <li class="nav-item"><a class="nav-link" href="income_expenses.asp?which=breakdown&amp;show_what=expense&amp;year=<%=iYear%>">Show Expense History</a></li>
                        <li class="nav-item"><a class="nav-link" href="income_expenses.asp?which=breakdown&amp;show_what=both&amp;year=<%=iYear%>">Show Both</a></li>
                    </ul>

                    <%If sShowWhat = "expense" Then%>
                        <h4 class="h4">Expense History</h4>

                        <table class="table table-striped table-condensed">
                            <tr>
                                <th>No.</th>
                                <th>When</th>
                                <th>Amount</th>
                                <th>To</th>
                                <th>For</th>
                                <th>Comments</th>
                            </tr>
                            <%For i = 0 To UBound(Expenses, 2) - 1%>
                                <tr>
                                    <td>
                                        <a href="javascript:pop('edit_expense.asp?finance_expense_id=<%=Expenses(0, i)%>', 900,200)"><%=i + 1%>)</a>
                                    </td>
                                    <td><%=Expenses(2, i)%></td>
                                    <td>$<%=Expenses(1, i)%></td>
                                    <td><%=Expenses(3, i)%></td>
                                    <td><%=Expenses(4, i)%></td>
                                    <td><%=Expenses(5, i)%></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <th colspan="2">Total Expense:</th>
                                <th>$<%=sngTotalExpense%></th>
                                <td colspan="3">&nbsp;</td>
                            </tr>
                        </table>
                    <%ElseIf sShowWhat = "income" Then%>
                        <h4 class="h4">Income History</h4>

                        <table class="table table-striped table-condensed">
                            <tr>
                                <th>No.</th>
                                <th>When</th>
                                <th>Amount</th>
                                <th>From</th>
                                <th>Type</th>
                                <th>Event</th>
                                <th>Sport</th>
                                <th>Comments</th>
                            </tr>
                            <%For i = 0 To UBound(Income, 2) - 1%>
                                <tr>
                                    <td>
                                        <a href="javascript:pop('edit_income.asp?finance_income_id=<%=Income(0, i)%>', 900,200)"><%=i + 1%>)</a>
                                    </td>
                                    <td><%=Income(1, i)%></td>
                                    <td>$<%=Income(2, i)%></td>
                                    <td><%=Income(3, i)%></td>
                                    <td><%=Income(4, i)%></td>
                                    <td><%=Income(5, i)%></td>
                                    <td><%=Income(6, i)%></td>
                                    <td><%=Income(7, i)%></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <th colspan="2">Total Income:</th>
                                <th>$<%=sngTotalIncome%></th>
                                <td colspan="5">&nbsp;</td>
                            </tr>
                        </table>
                    <%Else%>
                        <h4 class="h4">Expense History</h4>

                        <table class="table table-striped table-condensed">
                            <tr>
                                <th>No.</th>
                                <th>When</th>
                                <th>Amount</th>
                                <th>To</th>
                                <th>For</th>
                                <th>Comments</th>
                            </tr>
                            <%For i = 0 To UBound(Expenses, 2) - 1%>
                                <tr>
                                    <td>
                                        <a href="javascript:pop('edit_expense.asp?finance_expense_id=<%=Expenses(0, i)%>', 900,200)"><%=i + 1%>)</a>
                                    </td>
                                    <td><%=Expenses(2, i)%></td>
                                    <td>$<%=Expenses(1, i)%></td>
                                    <td><%=Expenses(3, i)%></td>
                                    <td><%=Expenses(4, i)%></td>
                                    <td><%=Expenses(5, i)%></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <th colspan="2">Total Expense:</th>
                                <th>$<%=sngTotalExpense%></th>
                                <td colspan="3">&nbsp;</td>
                            </tr>
                        </table>
        
                        <h4 class="h4">Income History</h4>

                        <table class="table table-striped table-condensed">
                            <tr>
                                <th>No.</th>
                                <th>When</th>
                                <th>Amount</th>
                                <th>From</th>
                                <th>Type</th>
                                <th>Event</th>
                                <th>Sport</th>
                                <th>Comments</th>
                            </tr>
                            <%For i = 0 To UBound(Income, 2) - 1%>
                                <tr>
                                    <td>
                                        <a href="javascript:pop('edit_income.asp?finance_income_id=<%=Income(0, i)%>', 900,200)"><%=i + 1%>)</a>
                                    </td>
                                    <td><%=Income(1, i)%></td>
                                    <td>$<%=Income(2, i)%></td>
                                    <td><%=Income(3, i)%></td>
                                    <td><%=Income(4, i)%></td>
                                    <td><%=Income(5, i)%></td>
                                    <td><%=Income(6, i)%></td>
                                    <td><%=Income(7, i)%></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <th colspan="2">Total Income:</th>
                                <th>$<%=sngTotalIncome%></th>
                                <td colspan="5">&nbsp;</td>
                            </tr>
                        </table>
                    <%End If%>
                <%Case "chronological"%>
		            <h4 class="h4">Payments & Expenses: Chronological List</h4>

                    <table class="table table-striped table-condensed">
                        <tr>
                            <th>No.</th>
                            <th>Date</th>
                            <th>Amount</th>
                            <th>Income/Expense</th>
                            <th>To/From</th>
                            <th>Type</th>
                            <th>Event</th>
                            <th>Sport</th>
                            <th>Comments</th>
                            <th>Balance</th>
                        </tr>
                        <%For i = 0 To UBound(Chronology, 2) - 1%>
                            <tr>
                                <td><%=i + 1%>)</td>
                                <td><%=Chronology(1, i)%></td>
                                <td>$<%=Chronology(2, i)%></td>
                                <td><%=Chronology(3, i)%></td>
                                <td><%=Chronology(4, i)%></td>
                                <td><%=Chronology(5, i)%></td>
                                <td><%=Chronology(6, i)%></td>
                                <td><%=Chronology(7, i)%></td>
                                <td><%=Chronology(8, i)%></td>
                                <th>$<%=Round(Chronology(9, i), 2)%></th>
                            </tr>
                        <%Next%>
                    </table>
                <%Case "income_category"%>
		            <h3>Income by Category</h3>

                    <%For j = 0 To UBound(IncomeTypes)%>
                        <%Call IncomeThisType(IncomeTypes(j))%>
                        <h4 class="h4"><%=IncomeTypes(j)%></h4>

                        <table class="table table-striped table-condensed">
                            <tr>
                                <th>No.</th>
                                <th>Date</th>
                                <th>Amount</th>
                                <th>To/From</th>
                                <th>Event</th>
                                <th>Sport</th>
                            </tr>
                            <%For i = 0 To UBound(IncomeByType, 2) - 1%>
                                <tr>
                                    <td><%=i + 1%>)</td>
                                    <td><%=IncomeByType(1, i)%></td>
                                    <td>$<%=IncomeByType(2, i)%></td>
                                    <td><%=IncomeByType(3, i)%></td>
                                    <td><%=IncomeByType(4, i)%></td>
                                    <td><%=IncomeByType(5, i)%></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <th colspan="5">Total:</th>
                                <th>$<%=sngIncomeTypeTotal%></th>
                            </tr>
                        </table>
                    <%Next%>
                <%Case "expense_category"%>
 		            <h3>Expenses by Category</h3>

                    <%For j = 0 To UBound(ExpenseTypes)%>
                        <%Call ExpenseThisType(ExpenseTypes(j))%>
                        <h4 class="h4"><%=ExpenseTypes(j)%></h4>

                        <table class="table table-striped table-condensed">
                            <tr>
                                <th>No.</th>
                                <th>Date</th>
                                <th>Amount</th>
                                <th>To/From</th>
                            </tr>
                            <%For i = 0 To UBound(ExpenseByType, 2) - 1%>
                                <tr>
                                    <td><%=i + 1%>)</td>
                                    <td><%=ExpenseByType(1, i)%></td>
                                    <td>$<%=ExpenseByType(2, i)%></td>
                                    <td><%=ExpenseByType(3, i)%></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <th colspan="3">Total:</th>
                                <th>$<%=sngExpenseTypeTotal%></th>
                            </tr>
                        </table>
                    <%Next%>
            <%End Select%>
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
