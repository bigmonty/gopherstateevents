<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim lStaffID, lEventID
Dim sStaffName, sTransType, sSport, sPmtMethod, sComments, sShowWhat
Dim iCheckNum, iYear
Dim sngMyEarned, sngMyPaid, sngMyBalance, sngTransAmt
Dim Events(), Staff(), SortArr(3), TransTypes(8), MyEarned(), MyPaid(), PmtMethods(3), MthlyEarned(11), MthlyPaid(11)
Dim dTransDate

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

lStaffID = Request.QueryString("staff_id")
If CStr(lStaffID) = vbNullString Then lStaffID = "0"

sShowWhat = Request.QueryString("show_what")
If sShowWhat = vbNullString Then sShowWhat = "both"

PmtMethods(0) = "Transfer"
PmtMethods(1) = "Check"
PmtMethods(2) = "Cash"
PmtMethods(3) = "Other"

TransTypes(0) = "Timing"
TransTypes(1) = "Mileage"
TransTypes(2) = "Expenses"
TransTypes(3) = "Race Prep"
TransTypes(4) = "Other Claim"
TransTypes(5) = "Payment"
TransTypes(6) = "Salary"
TransTypes(7) = "Draw"
TransTypes(8) = "Lease Pmt"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(1, 0)
sql = "SELECT StaffID, FirstName, LastName FROM Staff WHERE Active = 'y' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Staff(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim Events(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "' ORDER BY EventDate"
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
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "' ORDER BY MeetDate"
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

If Request.Form.Item("submit_new_data") = "submit_new_data" Then
    sTransType = Request.Form.Item("trans_type")
    sngTransAmt = Request.Form.Item("trans_amt")
    lEventID = Request.Form.Item("event_id")
    dTransDate = Request.Form.Item("trans_date")
    sSport = Request.Form.Item("sport")
    sPmtMethod = Request.Form.Item("pmt_method")
    iCheckNum = Request.Form.Item("check_num")
    If Not Request.Form.Item("comments") = vbNullString Then sComments = Replace(Request.Form.Item("comments"), "'", "''")

    sql = "INSERT INTO FinanceStaff (StaffID, TransType, TransAmt, TransDate, EventID, Sport, PmtMethod, CheckNum, Comments) VALUES ("
    sql = sql & lStaffID & ", '" & sTransType & "', " & sngTransAmt & ", '" & dTransDate & "', '" & lEventID & "', '" & sSport & "', '" & sPmtMethod
    sql = sql & "', '" & iCheckNum & "', '" & sComments & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_staff") = "submit_staff" Then
    lStaffID = Request.Form.Item("staff")
End If

If CStr(lStaffID) = vbNullString Then lStaffID = "0"
If Session("role") = "staff" Then lStaffID = Session("staff_id")

ReDim MyEarned(6, 0)
ReDim MyPaid(5, 0)

If Not CLng(lStaffID) = 0 Then
    'get staff name
    If Session("role") = "admin" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT FirstName, LastName FROM Staff WHERE StaffID = " & lStaffID
        rs.Open sql, conn, 1, 2
        sStaffName = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(0).Value, "''", "'")
        rs.Close
        Set rs = Nothing
    Else
        sStaffName = Session("my_name")
    End If

    sngMyEarned = 0
    sngMyPaid = 0
    sngMyBalance = 0

    'get staff earned
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceStaffID, TransType, TransAmt, TransDate, EventID, Sport, Comments FROM FinanceStaff "
    sql = sql & "WHERE StaffID = " & lStaffID & " AND TransType NOT IN('Payment','Lease Pmt') AND (TransDate >= '1/1/" & iYear & "' AND TransDate <= '12/31/" & iYear & "') ORDER BY TransDate DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngMyEarned = CSng(sngMyEarned) + CSng(rs(2).Value)

        MyEarned(0, i) = rs(0).Value
        MyEarned(1, i) = rs(1).Value
        MyEarned(2, i) = rs(2).Value
        MyEarned(3, i) = rs(3).Value
        MyEarned(4, i) = rs(4).Value
        MyEarned(5, i) = rs(5).Value
        If Not rs(6).Value & "" = "" Then MyEarned(6, i) = Replace(rs(6).Value, "''", "'")
        i = i + 1
        ReDim Preserve MyEarned(6, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get staff payments
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceStaffID, TransAmt, TransDate, PmtMethod, CheckNum, Comments FROM FinanceStaff WHERE StaffID = " & lStaffID
    sql = sql & " AND TransType IN ('Payment','Lease Pmt') AND (TransDate >= '1/1/" & iYear & "' AND TransDate <= '12/31/" & iYear & "') ORDER BY TransDate DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngMyPaid = CSng(sngMyPaid) + CSng(rs(1).Value)

        MyPaid(0, i) = rs(0).Value
        MyPaid(1, i) = rs(1).Value
        MyPaid(2, i) = rs(2).Value
        MyPaid(3, i) = rs(3).Value
        MyPaid(4, i) = rs(4).Value
        If Not rs(5).Value & "" = "" Then MyPaid(5, i) = Replace(rs(5).Value, "''", "'")
        i = i + 1
        ReDim Preserve MyPaid(5, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get staff balance 
    sngMyBalance = CSNg(sngMyEarned) - CSng(sngMyPaid)
End If

Private Function GetEventName(lThisEvent, sThisSport)
    Set rs = Server.CreateObject("ADODB.Recordset")

    If sThisSport = "Fitness Event" Then
        sql = "SELECT EventName FROM Events WHERE EventID = " & lThisEvent
        rs.Open sql, conn, 1, 2
    Else
        sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lThisEvent
        rs.Open sql, conn2, 1, 2
    End If

    If rs.RecordCount > 0 Then GetEventName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

Private Function TransEarned(sThisTransType)
    TransEarned = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TransAmt FROM FinanceStaff WHERE StaffID = " & lStaffID & " AND TransType = '" & sThisTransType & "' AND (TransDate >= '1/1/"
    sql = sql & iYear & "' AND TransDate <= '12/31/" & iYear & "')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        TransEarned = CSng(TransEarned) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    TransEarned = Round(CSng(TransEarned), 2)
End Function

For i = 0 To 11
    MthlyEarned(i) = 0
    MthlyPaid(i) = 0
Next

Private Function MthlyVal(sThisTransType, iThisMonth, iTransNum)
    Dim iNumDays

    MthlyVal = 0

    Select Case iThisMonth
        Case 2
            iNumDays = 28
        Case 4
            iNumDays = 30
        Case 6
            iNumDays = 30
        Case 9
            iNumDays = 30
        Case 11
            iNumDays = 30
        Case Else
            iNumDays = 31
    End Select

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TransAmt FROM FinanceStaff WHERE StaffID = " & lStaffID & " AND TransType = '" & sThisTransType & "' AND (TransDate >= '" & iThisMonth
    sql = sql & "/1/" & iYear & "' AND TransDate <= '" & iThisMonth & "/" & iNumDays & "/" & iYear & "')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MthlyVal = CSng(MthlyVal) + CSng(rs(0).Value)

        If CInt(iTransNum) <= 4 Then
            MthlyEarned(iThisMonth - 1) = CSng(MthlyEarned(iThisMonth - 1)) + CSng(rs(0).Value)
        Else
            MthlyPaid(iThisMonth - 1) = CSng(MthlyPaid(iThisMonth - 1)) + CSng(rs(0).Value)
        End If

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    MthlyVal = Round(CSng(MthlyVal), 2)
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Staff Finances</title>

<script>
$(function() {
    $( "#trans_date" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
        <%If Session("role") = "admin" Then%>   
            <!--#include file = "../../includes/admin_menu.asp" -->
        <%Else%>
            <!--#include file = "../../staff/staff_menu.asp" -->
        <%End If%>
		
		<div class="col-md-10">
            <%If Session("role") = "admin" Then%>   
                <!--#include file = "staff_nav.asp" -->
            <%End If%>
            

            <%If Session("role") = "admin" Then%>
		        <h3 class="h3">GSE Finances: Manage Staff</h3>
            <%Else%>
                <h3 class="h3">GSE Finances: Staff Summary for <%=Session("my_name")%></h3>
            <%End If%>

            <ul class="nav">
                <%For i = 2015 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="staff.asp?year=<%=i%>&amp;staff_id=<%=lStaffID%>"><%=i%></a></li>
                <%Next%>
           </ul>

            <%If Session("role") = "admin" Then%>
                <form class="form-inline" name="get_staff" method="post" action="staff.asp?year=<%=iYear%>">
                <label for="staff">Select Staff To Manage</label>
                <select class="form-control" name="staff" id="staff" onchange="this.form.submit1.click();">
                    <option value=""></option>
                    <%For i = 0 To UBound(Staff, 2) - 1%>
                        <%If CLng(Staff(0, i)) = CLng(lStaffID) Then%>
                            <option value="<%=Staff(0, i)%>" selected><%=Staff(1, i)%></option>
                        <%Else%>
                            <option value="<%=Staff(0, i)%>"><%=Staff(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
                <input type="hidden" name="submit_staff" id="submit_staff" value="submit_staff">
                <input type="submit" class="form-control" name="submit1" id="submit1" value="Manage This Staff Member">
                </form>
                <br>
            <%End If%>

            <%If Not CLng(lStaffID) = 0 Then%>
                <div class="row">
                    <div class="col-sm-6">
                        <h5 class="h5">Summary To Date</h5>
                        <ul class="list-group">
                            <li class="list-group-item">Earned:&nbsp;$<%=Round(CSng(sngMyEarned), 2)%></lEventID>
                            <li class="list-group-item">Paid:&nbsp;$<%=Round(CSng(sngMyPaid), 2)%></li>
                            <li class="list-group-item">Due:&nbsp;$<%=Round(CSng(sngMyBalance), 2)%></il>
                        </ul>

                        <h5 class="h5">Earnings By Type:</h5>
                        <ul class="list-group">
                            <%For i = 0 To UBound(TransTypes)%>
                                <%If Not TransTypes(i) = "Payment" Then%>
                                    <li class="list-group-item"><%=TransTypes(i)%>:&nbsp;$<%=TransEarned(TransTypes(i))%></li>
                                <%End If%>
                            <%Next%>
                        </ul>
                    </div>
                    <div class="col-sm-6 bg-danger">
                        <%If Session("role") = "admin" Then%>
                            <h5 class="h5">New Transaction:</h5>
                            <form class="form" name="enter_data" method="post" action="staff.asp?year=<%=iYear%>&amp;staff_id=<%=lStaffID%>">
                            <div class="form-group row">
                                <label for="trans_amt" class="control-label col-sm-4">Amount:</label>
				                <div class="col-sm-8">
                                    <input type="text" class="form-control" name="trans_amt" id="trans_amt">
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="trans_date"class="control-label col-sm-4">Date:</label>
				                <div class="col-sm-8">
                                    <input type="text" class="form-control" name="trans_date" id="trans_date">
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="trans_type"class="control-label col-sm-4">Type:</label>
				                <div class="col-sm-8">
                                    <select class="form-control" name="trans_type" id="trans_type">
                                        <option value=""></option>
                                        <%For i = 0 To UBound(TransTypes)%>
                                            <option value="<%=TransTypes(i)%>"><%=TransTypes(i)%></option>
                                        <%Next%>
                                    </select>
                                </div>
                            </div>
                             <div class="form-group row">
                                <label for="pmt_method"class="control-label col-sm-4">Pmt Method:</label>
				                <div class="col-sm-8">
                                    <select class="form-control" name="pmt_method" id="pmt_method">
                                        <option value=""></option>
                                        <%For i = 0 To UBound(PmtMethods)%>
                                            <option value="<%=PmtMethods(i)%>"><%=PmtMethods(i)%></option>
                                        <%Next%>
                                    </select>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="check_num"class="control-label col-sm-4">Check Num:</label>
				                <div class="col-sm-8">
                                    <input type="text" class="form-control" name="check_num" id="check_num">
                                </div>
                            </div>
                             <div class="form-group row">
                                <label for="event_id"class="control-label col-sm-4">Event:</label>
				                <div class="col-sm-8">
                                    <select class="form-control" name="event_id" id="event_id">
                                        <option value=""></option>
                                        <%For i = 0 To UBound(Events, 2) - 1%>
                                            <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                                        <%Next%>
                                    </select>
                                </div>
                            </div>
                             <div class="form-group row">
                                <label for="sport"class="control-label col-sm-4">Sport:</label>
				                <div class="col-sm-8">
                                    <select class="form-control" name="sport" id="sport">
                                        <option value=""></option>
                                        <option value="Fitness Event">Fitness Event</option>
                                        <option value="Nordic Ski">Nordic Ski</option>
                                        <option value="Cross-Country">Cross-Country</option>
                                    </select>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="comments"class="control-label col-sm-4">Comments:</label>
				                <div class="col-sm-8">
                                    <textarea class="form-control" name="comments" id="comments" rows="5"></textarea>
                                </div>
                            </div>
                            <div class="form-group">
                                <input type="hidden" name="submit_new_data" id="submit_new_data" value="submit_new_data">
                                <input type="submit" class="form-control" name="submit1" id="submit1" value="Submit New Data">
                            </div>
                            </form>
                        <%Else%>
                            &nbsp;
                        <%End If%>
                    </div>
                </div>
                
                <h5 class="h5">Month-by-Month</h5>
                <table class="table table-striped">
                    <tr>
                        <th>Type</th>
                        <%For i = 1 To 12%>
                            <th><%=MonthName(i, true)%></th>
                        <%Next%>
                    </tr>
                    <%For i = 0 To 4%>
                        <tr>
                            <td><%=TransTypes(i)%>:</td>
                            <%For j = 1 To 12%>
                                <td>$<%=MthlyVal(TransTypes(i), j, i)%></td>
                            <%Next%>
                        </tr>
                    <%Next%>
                    <tr>
                        <th>Total:</th>
                        <%For i = 1 To 12%>
                            <th>$<%=MthlyEarned(i - 1)%></th>
                        <%Next%>
                    </tr>
                    <%For i = 5 To UBound(TransTypes)%>
                        <tr>
                            <td><%=TransTypes(i)%>:</td>
                            <%For j = 1 To 12%>
                                <td>$<%=MthlyVal(TransTypes(i), j, i)%></td>
                            <%Next%>
                        </tr>
                    <%Next%>
                    <tr>
                        <th>Total:</th>
                        <%For i = 1 To 12%>
                            <th>$<%=MthlyPaid(i - 1)%></th>
                        <%Next%>
                    </tr>
                </table>

                <ul class="nav">
                    <li class="nav-item"><a class="nav-link" href="staff.asp?year=<%=iYear%>&amp;staff_id=<%=lStaffID%>&amp;show_what=both">Show Both</a></li>
                    <li class="nav-item"><a class="nav-link" href="staff.asp?year=<%=iYear%>&amp;staff_id=<%=lStaffID%>&amp;show_what=earnings">Show Earnings</a></li>
                    <li class="nav-item"><a class="nav-link" href="staff.asp?year=<%=iYear%>&amp;staff_id=<%=lStaffID%>&amp;show_what=payments">Show Payments</a></li>
               </ul>

                <%If sShowWhat = "earnings" or sShowWhat = "both" Then%>
                    <h5 class="h5">Earnings History</h5>
                    <table class="table table-striped">
                        <tr>
                            <th>No.</th>
                            <th>TransType</th>
                            <th>TransAmt</th>
                            <th>TransDate</th>
                            <th>Event</th>
                            <th>Sport</th>
                            <th>Comments</th>
                        </tr>
                        <%For i = 0 To UBound(MyEarned, 2) - 1%>
                            <tr>
                                <td>
                                    <%If Session("role") = "admin" Then%>   
                                        <a href="javascript:pop('edit_staff.asp?finance_staff_id=<%=MyEarned(0, i)%>',800,400)"><%=i + 1%></a>
                                    <%Else%>
                                        <%=i + 1%>
                                    <%End If%>
                                </td>
                                <td><%=MyEarned(1, i)%></td>
                                <td>$<%=MyEarned(2, i)%></td>
                                <td><%=MyEarned(3, i)%></td>
                                <td><%=GetEventName(MyEarned(4, i), MyEarned(5, i))%></td>
                                <td><%=MyEarned(5, i)%></td>
                                <td><%=MyEarned(6, i)%></td>
                            </tr>
                        <%Next%>
                    </table>
                <%End If%>

                <%If sShowWhat = "payments" or sShowWhat = "both" Then%>
                    <h5 class="h5">Payment History</h5>
                    <table class="table table-striped">
                        <tr>
                            <th>No.</th>
                            <th>TransAmt</th>
                            <th>TransDate</th>
                            <th>Payment Method</th>
                            <th>Check Num</th>
                            <th>Comments</th>
                        </tr>
                        <%For i = 0 To UBound(MyPaid, 2) - 1%>
                            <tr>
                                <td>
                                    <%If Session("role") = "admin" Then%>   
                                        <a href="javascript:pop('edit_staff.asp?finance_staff_id=<%=MyPaid(0, i)%>',800,400)"><%=i + 1%></a>
                                    <%Else%>
                                        <%=i + 1%>
                                    <%End If%>
                                </td>
                                <td>$<%=MyPaid(1, i)%></td>
                                <td><%=MyPaid(2, i)%></td>
                                <td><%=MyPaid(3, i)%></td>
                                <td><%=MyPaid(4, i)%></td>
                                <td><%=MyPaid(5, i)%></td>
                            </tr>
                        <%Next%>
                    </table>
                <%End If%>
            <%End If%>
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
