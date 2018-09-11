<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim iYear, iNumEvents, iActiveStaff, iEvntsToDate, iNumSport, iEvntsTtl
Dim sngMyEarned, sngMyPaid, sngMyBalance, sngTotalEarned, sngTotalPaid, sngTotalDue, sngSportIncome, sngSportExpense, sngSportProfit, sngSportMargin
Dim sngIncomeTotal, sngExpenseTotal, sngProfitTotal, sngMarginTotal
Dim sWhich
Dim Events(), SortArr(3), Sports(2), Staff()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

sWhich = REquest.QueryString("which")
If CStr(sWhich) = vbNullString Then sWhich = "Sport"

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

iEvntsToDate = 0

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

iEvntsToDate = UBound(Events, 2)

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

iActiveStaff = 0

i = 0
ReDim Staff(2, 0)
sql = "SELECT StaffID, FirstName, LastName, Active FROM Staff ORDER BY Active DESC, LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    Staff(2, i) = rs(3).Value

    If rs(3).Value = "y" Then iActiveStaff = CInt(iActiveStaff) + 1

	i = i + 1
	ReDim Preserve Staff(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

sngTotalEarned = 0
sngTotalPaid = 0
sngTotalDue = 0

sngIncomeTotal = 0
sngExpenseTotal = 0
sngProfitTotal = 0
sngMarginTotal = 0

Private Sub SportSumm(sThisSport)
    Dim x

    sngSportIncome = 0
    sngSportExpense = 0
    sngSportMargin = 0
    iNumSport = 0

    For x = 0 To UBound(Events, 2) - 1
        If Events(3, x) = sThisSport Then
            If CDate(Events(2, x)) <= Date Then
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT Invoice, Staffing, MiscCost, PartCost, LaborCost, Mileage FROM FinanceEvents WHERE Sport = '" & sThisSport & "' "
                sql = sql & "AND EventID = " & Events(0, x)
                rs.Open sql, conn, 1, 2
                Do While Not rs.EOF
                    sngSportIncome = CSng(sngSportIncome) + CSng(rs(0).Value)
                    sngSportExpense = CSng(sngSportExpense) + CSng(rs(1).Value) + CSng(rs(2).Value) + CSng(rs(3).Value) + CSng(rs(4).Value) + CSng(rs(5).Value)

                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing

                iNumSport = CInt(iNumSport) + 1

                sngSportProfit = CSng(sngSportIncome) - CSng(sngSportExpense)
                If Not sngSportIncome = "0" Then sngSportMargin = Round(CSng(sngSportProfit)/CSng(sngSportIncome), 2)
            End If
        End If
    Next

    sngIncomeTotal = CSNg(sngIncomeTotal) + CSng(sngSportIncome)
    sngExpenseTotal = CSng(sngExpenseTotal) + CSng(sngSportExpense)
    sngProfitTotal = CSng(sngProfitTotal) + CSng(sngSportProfit)
    If Not sngExpenseTotal = "0" Then sngMarginTotal = Round(CSng(sngProfitTotal)/CSNg(sngIncomeTotal), 2)

    iEvntsTtl = CInt(iEvntsTtl) + CInt(iNumSport)
End Sub

Private Sub MySummary(lThisStaff)
    sngMyEarned = 0
    sngMyPaid = 0
    sngMyBalance = 0
    iNumEvents = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TransAmt, TransType FROM FinanceStaff WHERE StaffID = " & lThisStaff & " AND (TransDate >= '1/1/" & iYear 
    sql = sql & "' AND TransDate <= '12/31/" & iYear & "')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If rs(1).Value = "Payment" Then
            sngMyPaid = CSng(sngMyPaid) + CSng(rs(0).Value)
            sngTotalPaid = CSNg(sngTotalPaid) + CSng(rs(0).Value)
        Else
            sngMyEarned = CSng(sngMyEarned) + CSng(rs(0).Value)
            sngTotalEarned = CSng(sngTotalEarned) + CSng(rs(0).Value)
        End If

        If rs(1).Value = "Timing" Then iNumEvents = CInt(iNumEvents) + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    sngMyBalance = CSng(sngMyEarned) - CSng(sngMyPaid)
    sngTotalDue = CSng(sngTotalDue) + CSng(sngMyBalance)
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Yearly Summary</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h3 class="h3">GSE Finances: Yearly Summary</h3>

            <div class="row">
                <div class="col-md-6">
                    <ul class="nav">
                        <%For i = 2015 To Year(Date) + 1%>
                            <li class="nav-item"><a class="nav-link" href="summary.asp?year=<%=i%>&amp;which=<%=sWhich%>"><%=i%></a></li>
                        <%Next%>
                   </ul>
                </div>
                <div class="col-md-6" style="text-align: right;">
                    <ul class="nav">
                        <li class="nav-item"><a class="nav-link" href="summary.asp?which=Sport&amp;year=<%=iYear%>">Summary by Sport</a></li>
                        <li class="nav-item"><a class="nav-link" href="summary.asp?which=Staff&amp;year=<%=iYear%>">Summary by Staff</a></li>
                    </ul>
                </div>
            </div>

            <%Select Case sWhich%>
                <%Case "Sport"%>
                    <h4 class="h4">Summary By Sport</h4>
                    <table class="table table-striped">
                        <tr>
                            <th rowspan="2">Sport (Num)</th>
                            <th colspan="2">Income</th>
                            <th colspan="2">Expense</th>
                            <th colspan="2">Profit</th>
                            <th rowspan="2">Margin</th>
                        </tr>
                        <tr>
                            <th>Total</th>
                            <th>Avg</th>
                            <th>Total</th>
                            <th>Avg</th>
                            <th>Total</th>
                            <th>Avg</th>
                        </tr>
                        <%For i = 0 To UBound(Sports)%>
                            <%Call SportSumm(Sports(i))%>
                            <tr>
                                <td><%=Sports(i)%> (<%=iNumSport%>)</td>
                                <td>$<%=sngSportIncome%></td>
                                <td>$
                                    <%If CInt(iNumSport) > 0 Then%>
                                        <%=Round(CSng(sngSportIncome)/Cint(iNumSport), 2)%>
                                    <%Else%>    
                                        0
                                    <%End If%>
                                </td>
                                <td>$<%=sngSportExpense%></td>
                                <td>$
                                    <%If CInt(iNumSport) > 0 Then%>
                                        <%=Round(CSng(sngSportExpense)/Cint(iNumSport), 2)%>
                                    <%Else%>    
                                        0
                                    <%End If%>
                                </td>
                                <td>$<%=sngSportProfit%></td>
                                <td>$
                                    <%If CInt(iNumSport) > 0 Then%>
                                        <%=Round(CSng(sngSportProfit)/Cint(iNumSport), 2)%>
                                    <%Else%>    
                                        0
                                    <%End If%>
                                </td>
                                <td><%=sngSportMargin*100%>%</td>
                            </tr>
                        <%Next%>
                        <tr>
                            <th>Totals: (<%=iEvntsTtl%>)</th>
                            <th>$<%=sngIncomeTotal%></th>
                            <%If CInt(iEvntsToDate) > 0 Then%>
                                <th>$<%=Round(CSng(sngIncomeTotal)/Cint(iEvntsToDate), 2)%></th>
                            <%Else%>
                            <th>$0</th>
                            <%End If%>
                            <th>$<%=sngExpenseTotal%></th>
                            <%If CInt(iEvntsToDate) > 0 Then%>
                                <th>$<%=Round(CSng(sngExpenseTotal)/Cint(iEvntsToDate), 2)%></th>
                            <%Else%>
                            <th>$0</th>
                            <%End If%>
                            <th>$<%=sngProfitTotal%></th>
                            <%If CInt(iEvntsToDate) > 0 Then%>
                                <th>$<%=Round(CSng(sngProfitTotal)/Cint(iEvntsToDate), 2)%></th>
                            <%Else%>
                            <th>$0</th>
                            <%End If%>
                            <th><%=sngMarginTotal*100%>%</th>
                        </tr>
                    </table>
                <%Case "Staff"%>
                    <h4 class="h4">Summary By Staff</h4>
                    <ul class="list-inline">
                        <li>Total Staff:</span>&nbsp;<%=UBound(Staff,2)%></li>
                        <li>Active:</span>&nbsp;<%=iActiveStaff%></li>
                        <li>Inactive:</span>&nbsp;<%=UBound(Staff,2) - CInt(iActiveStaff)%></li>
                    </ul>
                    <table class="table table-striped">
                        <tr>
                            <th>No,</th>
                            <th>Name</th>
                            <th>Active</th>
                            <th>Events</th>
                            <th>Earned</th>
                            <th>Paid</th>
                            <th>Balance</th>
                        </tr>
                        <%For i = 0 To UBound(Staff, 2) - 1%>
                            <%Call MySummary(Staff(0, i))%>
                            <tr>
                                <td><%=i + 1%>)</td>
                                <td><%=Staff(1, i)%></td>
                                <td><%=Staff(2, i)%></td>
                                <td><%=iNumEvents%></td>
                                <td>$<%=sngMyEarned%></td>
                                <td>$<%=sngMyPaid%></td>
                                <td>$<%=sngMyBalance%></td>
                            </tr>
                        <%Next%>
                        <tr>
                            <th colspan="4">Totals</th>
                            <th>$<%=sngTotalEarned%></th>
                            <th>$<%=sngTotalPaid%></th>
                            <th>$<%=sngTotalDue%></th>
                        </tr>
                    </table>
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
