<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i
Dim iYear, iNumNordic, iNumFitness, iNumCC, iYTDEvnts, iNordicTotal, iFitnessTotal, iCCTotal
Dim sngEventExpense, sngInvoices, sngYTDInv, sngAvgInv, sngMthlyExpense, sngYTDExpenses, sngAvgExpenses, sngOverhead, sngAvgOverhead, sngMthlyOverhead
Dim sngMthlyIncome, sngYTDIncome, sngAvgIncome, sngAvgEvnts
Dim Income(), Expenses(), MthlyEvents()

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

iYTDEvnts = 0
sngYTDInv = 0
sngAvgInv = 0
sngYTDExpenses = 0
sngAvgExpenses = 0
sngYTDIncome = 0
sngAvgIncome = 0
sngAvgEvnts = 0
iNordicTotal = 0
iFitnessTotal = 0
iCCTotal = 0
sngOverhead = 0

Private Sub MthlyData(iThisMonth)
    Dim x

    iNumFitness = 0
    iNumNordic = 0
    iNumCC = 0
    sngEventExpense = 0
    sngInvoices = 0
    sngMthlyExpense = 0
    sngMthlyIncome = 0
    sngMthlyOverhead = 0

    'get fitness events this month
    x = 0
    ReDim MthlyEvents(1, x)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If CInt(iThisMonth) = 12 Then
       sql = "SELECT EventID FROM Events WHERE EventDate >= '12/1/" & iYear & "' AND EventDate < '1/1/" & CInt(iYear) + 1 & "'"
    Else
        sql = "SELECT EventID FROM Events WHERE EventDate >= '" & iThisMonth & "/1/" & iYear & "' AND EventDate < '" & CInt(iThisMonth) + 1 & "/1/" 
        sql = sql & iYear & "'"
    End If
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iNumFitness = rs.RecordCount
    Do While Not rs.EOF
        MthlyEvents(0, x) = rs(0).Value
        MthlyEvents(1, x) = "Fitness Event"
        x  = x + 1
        ReDim Preserve MthlyEvents(1, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    iFitnessTotal = CInt(iFitnessTotal) + CInt(iNumFitness)

    'get nordic meets this month
    Set rs = Server.CreateObject("ADODB.Recordset")
    If CInt(iThisMonth) = 12 Then
       sql = "SELECT MeetsID FROM Meets WHERE MeetDate >= '12/1/" & iYear & "' AND MeetDate < '1/1/" & CInt(iYear) + 1 & "' AND Sport = 'Nordic Ski'"
    Else
        sql = "SELECT MeetsID FROM Meets WHERE MeetDate >= '" & iThisMonth & "/1/" & iYear & "' AND MeetDate < '" & CInt(iThisMonth) + 1 & "/1/" 
        sql = sql & iYear & "' AND Sport = 'Nordic Ski'"
    End If
    rs.Open sql, conn2, 1, 2
    If rs.RecordCount > 0 Then iNumNordic = rs.RecordCount
    Do While Not rs.EOF
        MthlyEvents(0, x) = rs(0).Value
        MthlyEvents(1, x) = "Nordic Ski"
        x  = x + 1
        ReDim Preserve MthlyEvents(1, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    iNordicTotal = CInt(iNordicTotal) + CInt(iNumNordic)

    'get cc meets this month
    Set rs = Server.CreateObject("ADODB.Recordset")
    If CInt(iThisMonth) = 12 Then
       sql = "SELECT MeetsID FROM Meets WHERE MeetDate >= '12/1/" & iYear & "' AND MeetDate < '1/1/" & CInt(iYear) + 1 & "' AND Sport = 'Cross-Country'"
    Else
        sql = "SELECT MeetsID FROM Meets WHERE MeetDate >= '" & iThisMonth & "/1/" & iYear & "' AND MeetDate < '" & CInt(iThisMonth) + 1 & "/1/" 
        sql = sql & iYear & "' AND Sport = 'Cross-Country'"
    End If
    rs.Open sql, conn2, 1, 2
    If rs.RecordCount > 0 Then iNumCC = rs.RecordCount
    Do While Not rs.EOF
        MthlyEvents(0, x) = rs(0).Value
        MthlyEvents(1, x) = "Cross-Country"
        x  = x + 1
        ReDim Preserve MthlyEvents(1, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    iCCTotal = CInt(iCCTotal) + CInt(iNumCC)

    iYTDEvnts = CInt(iYTDEvnts) + CInt(iNumFitness) + CInt(iNumNordic) + CInt(iNumCC)
    sngAvgEvnts = Round(CInt(iYTDEvnts)/CSng(iThisMonth), 2)

    For x = 0 To UBound(MthlyEvents, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Staffing, MiscCost, PartCost, LaborCost, Invoice FROM FinanceEvents WHERE EventID = " & MthlyEvents(0, x) & " AND Sport = '" 
        sql = sql & MthlyEvents(1, x) & "'"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            sngEventExpense = CSng(sngEventExpense) + CSng(rs(0).Value) +  CSng(rs(1).Value) +  CSng(rs(2).Value) +  CSng(rs(3).Value) 
            sngInvoices = CSng(sngInvoices) + CSng(rs(4).Value)
        End If
        rs.Close
        Set rs = Nothing
    Next

    sngYTDInv = Csng(sngYTDInv) + CSng(sngInvoices)
    sngAvgInv = Round(CSng(sngYTDInv)/CInt(iThisMonth), 2)

    Set rs = Server.CreateObject("ADODB.Recordset")
    If CInt(iThisMonth) = 12 Then
       sql = "SELECT AmtPaid FROM FinanceExpense WHERE WhenPaid >= '12/1/" & iYear & "' AND WhenPaid < '1/1/" & CInt(iYear) + 1 & "'"
    Else
        sql = "SELECT AmtPaid FROM FinanceExpense WHERE WhenPaid >= '" & iThisMonth & "/1/" & iYear & "' AND WhenPaid < '" & CInt(iThisMonth) + 1 & "/1/" 
        sql = sql & iYear & "'"
    End If
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngMthlyExpense = CSng(sngMthlyExpense) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    sngYTDExpenses = CSng(sngYTDExpenses) + CSng(sngMthlyExpense) + CSng(sngEventExpense)
    sngAvgExpenses = Round(CSng(sngYTDExpenses)/CInt(iThisMonth), 2)

    Set rs = Server.CreateObject("ADODB.Recordset")
    If CInt(iThisMonth) = 12 Then
       sql = "SELECT AmtPaid FROM FinanceExpense WHERE (WhenPaid >= '12/1/" & iYear & "' AND WhenPaid < '1/1/" & CInt(iYear) + 1 & "') "
       sql = sql & " AND (ExpenseType NOT IN ('Capital Outlay', 'H51 Software Expense', 'Travel & Lodging', 'Bob Schneider Draw', 'Bob Bakken Draw'))"
    Else
        sql = "SELECT AmtPaid FROM FinanceExpense WHERE (WhenPaid >= '" & iThisMonth & "/1/" & iYear & "' AND WhenPaid < '" & CInt(iThisMonth) + 1 & "/1/" 
        sql = sql & iYear & "') AND (ExpenseType NOT IN ('Capital Outlay', 'H51 Software Expense', 'Travel & Lodging', 'Bob Schneider Draw', 'Bob Bakken Draw'))"
    End If
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngMthlyOverhead = CSng(sngMthlyOverhead) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    sngOverhead = CSng(sngOverhead) + CSng(sngMthlyOverhead)
    sngAvgOverhead = Round(CSng(sngOverhead)/CInt(iThisMonth), 2)

    Set rs = Server.CreateObject("ADODB.Recordset")
    If CInt(iThisMonth) = 12 Then
       sql = "SELECT AmtRcvd FROM FinanceIncome WHERE WhenRcvd >= '12/1/" & iYear & "' AND WhenRcvd < '1/1/" & CInt(iYear) + 1 & "'"
    Else
        sql = "SELECT AmtRcvd FROM FinanceIncome WHERE WhenRcvd >= '" & iThisMonth & "/1/" & iYear & "' AND WhenRcvd < '" & CInt(iThisMonth) + 1 & "/1/" 
        sql = sql & iYear & "'"
    End If
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngMthlyIncome = CSng(sngMthlyIncome) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    sngYTDIncome = CSng(sngYTDIncome) + CSng(sngMthlyIncome)
    sngAvgIncome = Round(CSng(sngYTDIncome)/CInt(iThisMonth), 2)
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Month-by-Month</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            

		    <h4 class="h4">Finances: Month-by-Month</h4>

            <div style="margin: 0;padding: 0 10px 0 0;text-align: right;">
                <%For i = 2015 To Year(Date) + 1%>
                    <a href="month-by-month.asp?year=<%=i%>" style="font-size: 0.85em;"><%=i%></a>
                    <%If Not i = Year(Date) + 1 Then%>
                        &nbsp;|&nbsp;
                    <%End If%>
                <%Next%>
           </div>

            <table class="table table-striped">
                <tr>
                    <th rowspan="2" valign="bottom">Month</th>
                    <th style="text-align: center;color:#ff6a00;" colspan="4">Events</th>
                    <th style="text-align: center;color: #963636;" colspan="3">Expenses</th>
                    <th style="text-align: center;color: #808080;" colspan="4">Monthly Summary</th>
                    <th style="text-align: center;color: #259133;" colspan="5">YTD</th>
                    <th style="text-align: center;color: #2b3b9e;" colspan="5">Averages</th>
                </tr>
                <tr>
                    <th style="color:#ff6a00;">Nord</th>
                    <th style="color:#ff6a00;">Fitn</th>
                    <th style="color:#ff6a00;">CC</th>
                    <th style="color:#ff6a00;">Invoice</th>
                    <th style="color: #963636;">Evnt</th>
                    <th style="color: #963636;">Other</th>
                    <th style="color: #963636;">CODB</th>
                    <th style="color: #808080;">Evnt</th>
                    <th style="color: #808080;">Income</th>
                    <th style="color: #808080;">Expense</th>
                    <th style="color: #808080;">Balance</th>
                    <th style="color: #259133;">Evnt</th>
                    <th style="color: #259133;">Income</th>
                    <th style="color: #259133;">Expense</th>
                    <th style="color: #259133;">Balance</th>
                    <th style="color: #259133;">Invoice</th>
                    <th style="color: #2b3b9e;">Evnt</th>
                    <th style="color: #2b3b9e;">CODB</th>
                    <th style="color: #2b3b9e;">Income</th>
                    <th style="color: #2b3b9e;">Expense</th>
                    <th style="color: #2b3b9e;">Invoice</th>
                </tr>
                <%For i = 1 To 12%>
                    <%Call MthlyData(i)%>
                    <tr>
                        <th style="text-align: left;"><%=MonthName(i, 3)%></th>
                        <td style="color:#ff6a00;text-align: center;"><%=iNumNordic%></td>
                        <td style="color: #ff6a00;text-align: center;"><%=iNumFitness%></td>
                        <td style="color: #ff6a00;text-align: center;"><%=iNumCC%></td>
                        <td style="color: #ff6a00;">$<%=sngInvoices%></td>
                        <td style="color: #963636;">$<%=sngEventExpense%></td>
                        <td style="color: #963636;">$<%=sngMthlyExpense%></td>
                        <td style="color: #963636;">$<%=sngMthlyOverhead%></td>
                        <td style="color: #808080;text-align: center;"><%=CInt(iNumNordic) + CInt(iNumFitness) + CInt(iNumCC)%></td>
                        <td style="color: #808080;">$<%=sngMthlyIncome%></td>
                        <td style="color: #808080;">$<%=CSng(sngEventExpense) + CSng(sngMthlyExpense)%></td>
                        <td style="color: #808080;">
                            <%If CSng(sngMthlyIncome) - CSng(sngEventExpense) -CSng(sngMthlyExpense) < 0 Then%>
                                $<span style="color:red;"><%=CSng(sngMthlyIncome) - CSng(sngEventExpense) -CSng(sngMthlyExpense)%></span>
                            <%Else%>
                                $<%=CSng(sngMthlyIncome) - CSng(sngEventExpense) -CSng(sngMthlyExpense)%>
                            <%End If%>
                        </td>
                        <td style="color: #259133;text-align: center;"><%=iYTDEvnts%></td>
                        <td style="color: #259133;">$<%=sngYTDIncome%></td>
                        <td style="color: #259133;">$<%=CSng(sngYTDExpenses)%></td>
                        <td style="color: #259133;"> 
                            <%If CSng(sngYTDIncome) - CSng(sngYTDExpenses) < 0 Then%>
                                $<span style="color:red;"><%=CSng(sngYTDIncome) - CSng(sngYTDExpenses)%></span>
                            <%Else%>
                                $<%=CSng(sngYTDIncome) - CSng(sngYTDExpenses)%>
                            <%End If%>
                        </td>
                        <td style="color: #259133;">$<%=sngYTDInv%></td>
                        <td style="color: #2b3b9e;text-align: center;"><%=sngAvgEvnts%></td>
                        <td style="color: #2b3b9e;">$<%=CSng(sngAvgOverhead)%></td>
                        <td style="color: #2b3b9e;">$<%=sngAvgIncome%></td>
                        <td style="color: #2b3b9e;">$<%=sngAvgExpenses%></td>
                        <td style="color: #2b3b9e;">$<%=sngAvgInv%></td>
                    </tr>
                <%Next%>
                <tr>
                    <th>Total</th>
                    <th style="color:#ff6a00;"><%=iNordicTotal%></th>
                    <th style="color:#ff6a00;"><%=iFitnessTotal%></th>
                    <th style="color:#ff6a00;"><%=iCCTotal%></th>
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
