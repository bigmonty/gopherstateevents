<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lDebtID
Dim sSource, sDebtType, sPurpose, sTerms, sComments, sItemCmnts, sActivityClass, sActivityType, sActSource
Dim dIncurred, dActivityDate
Dim sngInitAmt, sngBalance, sngAmount
Dim DebtType(4), Activity()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lDebtID = Request.QueryString("debt_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

DebtType(0) = "Loan-Bank"
DebtType(1) = "Loan-Private"
DebtType(2) = "Recurring"
DebtType(3) = "Installment"
DebtType(4) = "Other"

If Request.Form.Item("submit_new") = "submit_new" Then
    sActSource = Replace(Request.Form.Item("act_source"), "''", "'")
    sActivityType = Replace(Request.Form.Item("activity_type"), "''", "'")
    If Not Request.Form.Item("item_cmnts") & "" = "" Then sItemCmnts = Replace(Request.Form.Item("item_cmnts"), "''", "'")
    dActivityDate = Request.Form.Item("activity_date")
    sngAmount = Request.Form.Item("amount")
    sActivityClass = Replace(Request.Form.Item("activity_class"), "''", "'")

    sql = "INSERT INTO DebtActivity(DebtID, Amount, ActivityDate, ActivityClass, Type, Source, Comments) VALUES (" & lDebtID & ", " & sngAmount & ", '" 
    sql = sql & dActivityDate & "', '" & sActivityClass & "', '" & sActivityType & "', '" & sActSource & "', '" & sItemCmnts & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sActSource = vbNullString
    sActivityType = vbNullString
    sItemCmnts = vbNullString
    dActivityDate = vbNullString
    sngAmount = vbNullString
    sActivityClass = vbNullString
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
    sSource = Replace(Request.Form.Item("source"), "''", "'")
    sDebtType = Request.Form.Item("debt_type")
    sPurpose = Replace(Request.Form.Item("purpose"), "''", "'")
    sTerms = Replace(Request.Form.Item("terms"), "''", "'")
    If Not Request.Form.Item("comments") & "" = "" Then sComments = Replace(Request.Form.Item("comments"), "''", "'")
    dIncurred = Request.Form.Item("incurred")
    sngInitAmt = Request.Form.Item("init_amt")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Source, DebtType, InitAmt, Incurred, Purpose, Terms, Comments FROM Debt WHERE DebtID = " & lDebtID
    rs.Open sql, conn, 1, 2
    rs(0).Value = sSource
    rs(1).Value = sDebtType
    rs(2).Value = sngInitAmt
    rs(3).Value = dIncurred
    rs(4).Value = sPurpose
    rs(5).Value = sTerms
    rs(6).Value = sComments
    rs.Update
    rs.Close
    Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Source, DebtType, InitAmt, Incurred, Purpose, Terms, Comments FROM Debt WHERE DebtID = " & lDebtID
rs.Open sql, conn, 1, 2
sSource = rs(0).Value
sDebtType = rs(1).Value
sngInitAmt = rs(2).Value
dIncurred = rs(3).Value
sPurpose = rs(4).Value
sTerms = rs(5).Value
sComments = rs(6).Value
rs.Close
Set rs = Nothing

'get activity
sngBalance = 0

i = 0
ReDim Activity(6, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DebtActivityID, Amount, ActivityDate, ActivityClass, Type, Source, Comments FROM DebtActivity WHERE DebtID = " & lDebtID 
sql = sql & " ORDER BY ActivityDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    For j = 0 To 6
        If j = 1 Then
            Activity(1, i) = "$" & rs(1).Value
        Else
            If Not rs(j).Value & "" = "" Then Activity(j, i) = Replace(rs(j).Value, "''", "'")
        End If
    Next

    If rs(3).Value = "increase" Then
        sngBalance = CSng(sngBalance) + CSng(rs(0).Value)
    Else
        sngBalance = CSng(sngBalance) - CSng(rs(0).Value)
    End If

    i = i + 1
    ReDim Preserve Activity(6, i)

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy;Finances: Debt Item</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
	<h3 class="h3">GSE Finances: Edit Debt Item</h3>

    <div class="bg-success">
        <h4 class="h4">Edit Item</h4>

        <form class="form-horizontal" name="edit_item" method="post" action="debt_item.asp?debt_id=<%=lDebtID%>">
        <div class="form-group">
            <label for="source" class="control-label col-xs-1">Source:</label>
            <div class="col-xs-3">
                <input type="text" class="form-control input-sm" name="source" id="source" maxlength="20" value="<%=sSource%>">
            </div>
            <label for="debt_type" class="control-label col-xs-1">Type:</label>
            <div class="col-xs-3">
                <select class="form-control input-sm" name="debt_type" id="debt_type">
                    <%For i = 0 To UBound(DebtType)%>
                        <%If sDebtType = DebtType(i) Then%>
                            <option value="<%=DebtType(i)%>" selected><%=DebtType(i)%></option>
                        <%Else%>
                            <option value="<%=DebtType(i)%>"><%=DebtType(i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </div>
            <label for="init_amt" class="control-label col-xs-1">Amount:</label>
            <div class="col-xs-3">
                <input type="text" class="form-control input-sm" name="init_amt" id="init_amt" value="<%=sngInitAmt%>">
            </div>
        </div>
        <div class="form-group">
            <label for="incurred" class="control-label col-xs-1">Incurred:</label>
            <div class="col-xs-3">
                <input type="text" class="form-control input-sm" name="incurred" id="incurred" value="<%=dIncurred%>">
            </div>
            <label for="purpose" class="control-label col-xs-1">Purpose:</label>
            <div class="col-xs-3">
                <input type="text" class="form-control input-sm" name="purpose" id="purpose" maxlength="50" value="<%=sPurpose%>">
            </div>
            <label for="terms" class="control-label col-xs-1">Terms:</label>
            <div class="col-xs-3">
                <input type="text" class="form-control input-sm" name="terms" id="terms" maxlength="50" value="<%=sTerms%>">
            </div>
        </div>
        <div class="form-group">
            <label for="comments" class="control-label col-xs-1">Cmnts:</label>
            <div class="col-xs-11">
                <textarea class="form-control input-sm" name="comments" id="comments" cols="100" rows="2"><%=sComments%></textarea>
            </div>
        </div>
        <div class="form-group">
            <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
            <input type="submit" class="form-control input-sm" name="submit1" id="submit1" value="Save Changes">
        </div>
        </form>
    </div>

    <hr>
    <h4 class="h4">Add Item Activity</h4>

    <form class="form-horizontal" name="add_activity" method="post" action="debt_item.asp?debt_id=<%=lDebtID%>">
    <div class="form-group">
        <label for="amount" class="control-label col-xs-1">Amount:</label>
        <div class="col-xs-3">
            <input type="text" class="form-control input-sm" name="amount" id="amount">
        </div>
        <label for="act_source" class="control-label col-xs-1">Source:</label>
        <div class="col-xs-3">
            <input type="text" class="form-control input-sm" name="act_source" id="act_source" maxlength="20">
        </div>
        <label for="activity_class" class="control-label col-xs-1">Class:</label>
        <div class="col-xs-3">
            <select class="form-control input-sm" name="activity_class" id="activity_class">
                <option value="">&nbsp;</option>
                <option value="decrease">decrease</option>
                <option value="increase">increase</option>
            </select>
        </div>
    </div>
    <div class="form-group">
        <label for="activity_date" class="control-label col-xs-1">Date:</label>
        <div class="col-xs-3">
            <input type="text" class="form-control input-sm" name="activity_date" id="activity_date">
        </div>
        <label for="activity_type" class="control-label col-xs-1">Type:</label>
        <div class="col-xs-7">
            <input type="text" class="form-control input-sm" name="activity_type" id="activity_type" maxlength="50">
        </div>
    </div>
    <div class="form-group">
        <label for="item_cmnts" class="control-label col-xs-1">Cmnts:</label>
        <div class="col-xs-11">
            <textarea class="form-control input-sm" name="item_cmnts" id="item_cmnts" rows="2"></textarea>
        </div>
    </div>
    <div class="form-group">
        <input type="hidden" name="submit_new" id="submit_new" value="submit_new">
        <input type="submit" class="form-control input-sm" name="submit2" id="submit2" value="Enter Activity">
    </div>
    </form>

    <hr>

    <h4 class="h4">Activity Summary</h4>

    <table class="table table-striped">
        <tr>
            <th>Amount</th>
            <th>Date</th>
            <th>Class</th>
            <th>Type</th>
            <th>Source</th>
            <th>Comments</th>
        </tr>

        <%For i = 0 To UBound(Activity, 2) - 1%>
            <tr>
                <%For j = 1 To 6%>
                    <td><%=Activity(j, i)%></td>
                <%Next%>
            </tr>
        <%Next%>
    </table>
</div>	
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
