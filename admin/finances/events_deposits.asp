<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim iYear
Dim sComments
Dim sngAmtRcvd, sngAmtRcvdTotal
Dim Events(), SortArr(3)
Dim dWhenRcvd

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

If Request.Form.Item("submit_this") = "submit_this" Then
    For i = 0 To UBound(Events, 2) - 1
        bFound = False
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AmtRcvd, WhenRcvd, Comments FROM FinanceIncome WHERE IncomeType = 'Race Deposit' AND EventID = " & Events(0, i) & " AND Sport = '"
        sql = sql & Events(3, i) & "'"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then 
            sngAmtRcvd = CSng(rs(0).Value)
            dWhenRcvd = rs(1).Value
            sComments = vbNullString
            bFound = True
        End If
        rs.Close
        Set rs = Nothing

        If bFound = False Then
        End If
    Next
End If

sngAmtRcvdTotal = "0"

Private Sub EventData(lThisEvent, sThisSport)
    sngAmtRcvd = 0
    dWhenRcvd = vbNullString
    sComments = vbNullString

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtRcvd, WhenRcvd, Comments FROM FinanceIncome WHERE IncomeType = 'Race Deposit' AND EventID = " & lThisEvent & " AND Sport = '"
    sql = sql & sThisSport & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        sngAmtRcvd = CSng(rs(0).Value)
        dWhenRcvd = rs(1).Value
        sComments = vbNullString
    End If
    rs.Close
    Set rs = Nothing

    sngAmtRcvdTotal = CSng(sngAmtRcvdTotal) + CSng(sngAmtRcvd)
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Events Ledger</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <!--#include file = "events_nav.asp" -->

		    <h3 class="h3">GSE Finances: Events Deposits</h3>

            <ul class="list-inline bg-warning">
                <%For i = 2015 To Year(Date) + 1%>
                    <li><a href="events_deposits.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
           </ul>

            <form class="form" name="deposits" method="post" action="events_deposits.asp?year=<%=iYear%>">
            <table class="table table-striped">
                <tr>
                    <td colspan="5">
                        <input type="hidden" name="submt_deposit" id="submt_deposit" value="submt_deposit">
                        <input type="submit" class="form-control" name="submit1" id="submit1" value="Save Changes">
                    </td>
                </tr>
                <tr>
                    <th>No.</th>
                    <th>Event/Meet (Date)</th>
                    <th>Deposit</th>
                    <th>When Received</th>
                    <th>Comments</th>
                </tr>
                <%For j = 0 To UBound(Events, 2) - 1%>
                    <%Call EventData(Events(0, j), Events(3, j))%>
                    <tr>
                        <td><%=j + 1%></td>
                        <td><%=Events(1, j)%> (<%=Events(2, j)%>)</td>
                        <td><input type="text" class="form-control" name="amt_rcvd_<%=Events(0, j)%>" id="amt_rcvd_<%=Events(0, j)%>" value="<%=sngAmtRcvd%>"></td>
                        <td><input type="text" class="form-control" name="when_rcvd_<%=Events(0, j)%>" id="when_rcvd_<%=Events(0, j)%>" value="<%=dWhenRcvd%>"></td>
                        <td><input type="text" class="form-control" name="comments_<%=Events(0, j)%>" id="comments_<%=Events(0, j)%>" value="<%=sComments%>"></td>
                    </tr>
                <%Next%>
                <tr>
                    <th colspan="2">Column Totals</th>
                    <th colspan="3">$<%=sngAmtRcvdTotal%></th>
                </tr>
            </table>
            </form>
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
