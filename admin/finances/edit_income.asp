<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim lFinanceIncomeID, lEventID
Dim sIncomeType, sComments, sFromTo, sDelete, sSport
Dim iYear
Dim sngAmt
Dim IncomeTypes(10), Sports(2), Events(), SortArr(2)
Dim dWhen

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lFinanceIncomeID = Request.QueryString("finance_income_id")

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Sports(0) = "Fitness Event"
Sports(1) = "Nordic Ski"
Sports(2) = "Cross-Country"

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
ReDim Events(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '1/1/" & CInt(iYear) & "' AND EventDate <= '12/31/" & Year(Date) & "'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
    Events(2, i) = rs(2).Value
    i = i + 1
    ReDim Preserve Events(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate >= '1/1/" & CInt(iYear) & "' AND MeetDate <= '12/31/" & Year(Date) & "'"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
    Events(2, i) = rs(2).Value
    i = i + 1
    ReDim Preserve Events(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'order by date
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

If Request.Form.Item("submit_changes") = "submit_changes" Then
    sDelete = Request.Form.Item("delete")
    sngAmt = Request.Form.Item("amt_rcvd")
    dWhen = Request.Form.Item("when_rcvd")
    sFromTo = Request.Form.Item("rcvd_from")
    sIncomeType = Request.Form.Item("income_type")
    lEventID = Request.Form.Item("event_id")
    sSport = Request.Form.Item("sport")
    If Not Request.Form.Item("comments") = vbNullString Then sComments = Replace(Request.Form.Item("comments"), "'", "''")

    If CStr(lEventID) = vbNullString Then lEventID = 0

    If sDelete = "on" Then
        sql = "DELETE FROM FinanceIncome WHERE FinanceIncomeID = " & lFinanceIncomeID
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AmtRcvd, WhenRcvd, RcvdFrom, IncomeType, EventID, Sport, Comments FROM FinanceIncome WHERE FinanceIncomeID = " & lFinanceIncomeID
        rs.Open sql, conn, 1, 2
        rs(0).Value = sngAmt
        rs(1).Value = dWhen
        rs(2).Value = sFromTo
        rs(3).Value = sIncomeType
        rs(4).Value = lEventID
        rs(5).Value = sSport
        rs(6).Value = sComments
        rs.Update
        rs.Close
        Set rs = Nothing
    End If

    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AmtRcvd, WhenRcvd, RcvdFrom, IncomeType, EventID, Sport, Comments FROM FinanceIncome WHERE FinanceIncomeID = " & lFinanceIncomeID
rs.Open sql, conn, 1, 2
sngAmt = rs(0).Value
dWhen = rs(1).Value
sFromTo = rs(2).Value
sIncomeType = rs(3).Value
lEventID = rs(4).Value
sSport = rs(5).Value
If Not rs(6).Value & "" = "" Then sComments = Replace(rs(6).Value, "''", "'")
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Edit Expenses</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    th{
        padding: 2px 2px 2px 5px;
        text-align: right;
    }
    
    td{
        padding: 2px 2px 2px 5px;
        text-align: left;
    }
</style>

</head>

<body>
<div>
    <h3>Edit Income</h3>

    <form name="edit_income" method="post" action="edit_income.asp?finance_income_id=<%=lFinanceIncomeID%>">
    <table style="margin: 0;">
        <tr>
            <th style="text-align: right;">Amount:</th>
            <td>$<input type="text" name="amt_rcvd" id="amt_rcvd" size="4" value="<%=sngAmt%>"></td>
            <th style="text-align: right;">When Rcvd:</th>
            <td><input type="text" name="when_rcvd" id="when_rcvd" size="4" value="<%=dWhen%>"></td>
            <th style="text-align: right;">Rcvd From:</th>
            <td><input type="text" name="rcvd_from" id="rcvd_from" value="<%=sFromTo%>"></td>
        </tr>
        <tr>
            <th style="text-align: right;">Type:</th>
            <td>
                <select name="income_type" id="income_type">
                    <%For i = 0 To UBound(IncomeTypes)%>
                        <%If CStr(sIncomeType) = CStr(IncomeTypes(i)) Then%>
                            <option value="<%=IncomeTypes(i)%>" selected><%=IncomeTypes(i)%></option>
                        <%Else%>
                            <option value="<%=IncomeTypes(i)%>"><%=IncomeTypes(i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </td>
            <th style="text-align: right;">Event:</th>
            <td>
                <select name="event_id" id="event_id">
                    <option value=""></option>
                    <%For i = 0 To UBound(Events, 2) - 1%>
                        <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
                            <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
                        <%Else%>
                            <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </td>
            <th style="text-align: right;">Sport:</th>
            <td>
                <select name="sport" id="sport">
                    <option value=""></option>
                    <%For i = 0 To UBound(Sports)%>
                        <%If CStr(sSport) = CStr(Sports(i)) Then%>
                            <option value="<%=Sports(i)%>" selected><%=Sports(i)%></option>
                        <%Else%>
                            <option value="<%=Sports(i)%>"><%=Sports(i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </td>
        </tr>
        <tr>
            <th style="text-align: right;">Comments:</th>
            <td colspan="5"><input type="text" name="comments" id="comments" size="75" value="<%=sComments%>"></td>
        </tr>
         <tr>    
            <td class="alt"style="text-align: center;color: red;" colspan="6">
                <input type="checkbox" name="delete" id="delete">&nbsp;Delete Record (There is no Undo for this action!)
            </td>
        </tr>
       <tr>
            <td style="text-align:center;padding-left: 10px;" colspan="6">
                <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                <input type="submit" name="submit2" id="submit2" value="Submit Changes">
            </td>
        </tr>
    </table>
    </form>
</div>
<%	
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
