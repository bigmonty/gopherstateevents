<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lFinanceExpenseID
Dim sExpenseType, sComments, sFromTo, sDelete
Dim sngAmt
Dim ExpenseTypes(13)
Dim dWhen

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lFinanceExpenseID = Request.QueryString("finance_expense_id")

ExpenseTypes(0) = "Bank Charges"
ExpenseTypes(1) = "Bob Bakken Draw"
ExpenseTypes(2) = "Bob Schneider Draw"
ExpenseTypes(3) = "Computer & Internet Expenses"
ExpenseTypes(4) = "H51 Software Expense"
ExpenseTypes(5) = "Insurance"
ExpenseTypes(6) = "Loan Repay"
ExpenseTypes(7) = "Materials & Equipment"
ExpenseTypes(8) = "Meals & Entertainment"
ExpenseTypes(9) = "Office Supplies"
ExpenseTypes(10) = "Professional Fees"
ExpenseTypes(11) = "Rent Expense"
ExpenseTypes(12) = "Travel & Lodging"
ExpenseTypes(13) = "Misc Expenses"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_changes") = "submit_changes" Then
    sDelete = Request.Form.Item("delete")
    sngAmt = Request.Form.Item("amt_paid")
    dWhen = Request.Form.Item("when_paid")
    sFromTo = Request.Form.Item("paid_to")
    If Not Request.Form.Item("expense_type") = vbNullString Then sExpenseType = Replace(Request.Form.Item("expense_type"), "'", "''")
    If Not Request.Form.Item("comments") = vbNullString Then sComments = Replace(Request.Form.Item("comments"), "'", "''")

    If sDelete = "on" Then
        sql = "DELETE FROM FinanceExpense WHERE FinanceExpenseID = " & lFinanceExpenseID
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AmtPaid, WhenPaid, PaidTo, ExpenseType, Comments FROM FinanceExpense WHERE FinanceExpenseID = " & lFinanceExpenseID
        rs.Open sql, conn, 1, 2
        rs(0).Value = sngAmt
        rs(1).Value = dWhen
        rs(2).Value = sFromTo
        rs(3).Value = sExpenseType
        rs(4).Value = sComments
        rs.Update
        rs.Close
        Set rs = Nothing
    End If

    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AmtPaid, WhenPaid, PaidTo, ExpenseType, Comments FROM FinanceExpense WHERE FinanceExpenseID = " & lFinanceExpenseID
rs.Open sql, conn, 1, 2
sngAmt = rs(0).Value
dWhen = rs(1).Value
sFromTo = rs(2).Value
sExpenseType = rs(3).Value
If Not rs(4).Value & "" = "" Then sComments = Replace(rs(4).Value, "''", "'")
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
    <h3>Edit Expense</h3>

    <form name="edit_expense" method="post" action="edit_expense.asp?finance_expense_id=<%=lFinanceExpenseID%>">
    <table style="margin: 0;">
        <tr>
            <th style="text-align: right;">Amount:</th>
            <td>$<input type="text" name="amt_paid" id="amt_paid" size="4" value="<%=sngAmt%>"></td>
            <th style="text-align: right;">When Paid:</th>
            <td><input type="text" name="when_paid" id="when_paid" size="4" value="<%=dWhen%>"></td>
            <th style="text-align: right;">Paid To:</th>
            <td><input type="text" name="paid_to" id="paid_to" value="<%=sFromTo%>"></td>
            <th style="text-align: right;">Type:</th>
            <td>
                <select name="expense_type" id="expense_type">
                    <%For i = 0 To UBound(ExpenseTypes)%>
                        <%If CStr(sExpenseType) = CStr(ExpenseTypes(i)) Then%>
                            <option value="<%=ExpenseTypes(i)%>" selected><%=ExpenseTypes(i)%></option>
                        <%Else%>
                            <option value="<%=ExpenseTypes(i)%>"><%=ExpenseTypes(i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </td>
        </tr>
        <tr>
            <th style="text-align: right;">Comments:</th>
            <td colspan="7"><input type="text" name="comments" id="comments" size="110" value="<%=sComments%>"></td>
        </tr>
        <tr>    
            <td class="alt"style="text-align: center;color: red;" colspan="8">
                <input type="checkbox" name="delete" id="delete">&nbsp;Delete Record (There is no Undo for this action!)
            </td>
        </tr>
       <tr>
            <td style="text-align:center;padding-left: 10px;" colspan="8">
                <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                <input type="submit" name="submit1" id="submit1" value="Submit Changes">
            </td>
        </tr>
    </table>
    </form>
</div>
<%	
conn.Close
Set conn = Nothing
%>
</body>
</html>
