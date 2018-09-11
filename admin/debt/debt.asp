<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lDebtID
Dim sSource, sDebtType, sPurpose, sTerms, sComments
Dim dIncurred
Dim sngInitAmt, sngTtlIncurred, sngTtlBalance
Dim Debt(), DebtType(4)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

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
    sSource = Replace(Request.Form.Item("source"), "''", "'")
    sDebtType = Request.Form.Item("debt_type")
    sPurpose = Replace(Request.Form.Item("purpose"), "''", "'")
    sTerms = Replace(Request.Form.Item("terms"), "''", "'")
    If Not Request.Form.Item("comments") & "" = "" Then sComments = Replace(Request.Form.Item("comments"), "''", "'")
    dIncurred = Request.Form.Item("incurred")
    sngInitAmt = Request.Form.Item("init_amt")

    sql = "INSERT INTO Debt(Source, DebtType, InitAmt, Incurred, Purpose, Terms, Comments) VALUES ('" & sSource & "', '" & sDebtType & "', "
    sql = sql & sngInitAmt & ", '" & dIncurred & "', '" & sPurpose & "', '" & sTerms & "', '" & sComments & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sSource = vbNullString
    sDebtType = vbNullString
    sPurpose = vbNullString
    sTerms = vbNullString
    sComments = vbNullString
    dIncurred = vbNullString
    sngInitAmt = vbNullString
End If

sngTtlIncurred = 0
sngTtlBalance = 0

i = 0
ReDim Debt(7, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DebtID, Source, DebtType, InitAmt, Incurred, Purpose, Terms FROM Debt ORDER BY Incurred DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    For j = 0 To 6
        If j = 3 Then
            Debt(j, i) = "$" & rs(3).Value
        Else
            If Not rs(j) & "" = "" Then Debt(j, i) = Replace(rs(j).Value, "''", "'")
        End If
    Next
    Debt(7, i) = GetBalance(rs(0).Value, rs(3).Value)

    sngTtlIncurred = CSng(sngTtlIncurred) + CSng(rs(3).Value)

    i = i + 1
    ReDim Preserve Debt(7, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function GetBalance(lThisDebt, sngThisInitialAmt)
    GetBalance = sngThisInitialAmt

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Amount FROM DebtActivity WHERE DebtID = " & lThisDebt & " AND ActivityClass = 'increase'"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        GetBalance = CSng(GetBalance) + CSng(rs2(0).Value)
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Amount FROM DebtActivity WHERE DebtID = " & lThisDebt & " AND ActivityClass = 'decrease'"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        GetBalance = CSng(GetBalance) - CSng(rs2(0).Value)
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    sngTtlBalance = CSng(sngTtlBalance) + CSng(GetBalance)
    GetBalance = "$" & GetBalance
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Debt Tracker</title>
<!--#include file = "../../includes/js.asp" -->

</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">GSE Finances: Debt Tracker</h3>

            <div class="bg-warning"><a href="debt.asp">Refresh</a></div>

            <h4 class="h4">Debt Summary</h4>
            <ul class="list-inline">
                <li class="list-group-item">Total Debt Incurred:&nbsp;$<%=sngTtlIncurred%></li>
                <li class="list-group-item">Outstanding Balance:&nbsp;$<%=sngTtlBalance%></li>
                <li class="list-group-item">Overall Change:&nbsp;$<%=CSng(sngTtlIncurred) - CSng(sngTtlBalance)%></li>
            </ul>

            <div class="bg-success">
                <h4 class="h4">Add Debt Account</h4>

                <form class="form-horizontal" name="add_debt" method="post" action="debt.asp">
                <div class="form-group">
                    <label for="source" class="control-label col-xs-1">Source:</label>
                    <div class="col-xs-3">
                        <input type="text" class="form-control input-sm" name="source" id="source" maxlength="20">
                    </div>
                    <label for="debt_type" class="control-label col-xs-1">Type:</label>
                    <div class="col-xs-3">
                        <select class="form-control input-sm" name="debt_type" id="debt_type">
                            <option value="">&nbsp;</option>
                            <%For i = 0 To UBound(DebtType)%>
                                <option value="<%=DebtType(i)%>"><%=DebtType(i)%></option>
                            <%Next%>
                        </select>
                    </div>
                    <label for="init_amt" class="control-label col-xs-1">Amount:</label>
                    <div class="col-xs-3">
                        <input type="text" class="form-control input-sm" name="init_amt" id="init_amt">
                    </div>
                </div>
                <div class="form-group">
                    <label for="incurred" class="control-label col-xs-1">Incurred:</label>
                    <div class="col-xs-3">
                        <input type="text" class="form-control input-sm" name="incurred" id="incurred">
                    </div>
                    <label for="purpose" class="control-label col-xs-1">Purpose:</label>
                    <div class="col-xs-3">
                        <input type="text" class="form-control input-sm" name="purpose" id="purpose" maxlength="50">
                    </div>
                    <label for="terms" class="control-label col-xs-1">Terms:</label>
                    <div class="col-xs-3">
                        <input type="text" class="form-control input-sm" name="terms" id="terms" maxlength="50">
                    </div>
                </div>
                <div class="form-group">
                    <label for="comments" class="control-label col-xs-1">Comments:</label>
                    <div class="col-xs-11">
                        <textarea class="form-control input-sm" name="comments" id="comments" rows="2"></textarea>
                    </div>
                </div>
                <div class="form-group">
                    <input type="hidden" name="submit_new" id="submit_new" value="submit_new">
                    <input type="submit" class="form-control input-sm" name="submit1" id="submit1" value="Add Debt">
                </div>
                </form>
            </div>

            <hr>

            <h4 class="h4">Existing Debt</h4>

            <table class="table table-striped">
                <tr>
                    <th>Source</th>
                    <th>Type</th>
                    <th>Init Amt</th>
                    <th>Incurred</th>
                    <th>Purpose</th>
                    <th>Terms</th>
                    <th>Balance</th>
                </tr>

                <%For i = 0 To UBound(Debt, 2) - 1%>
                    <tr>
                        <td>
                            <a href="javascript:pop('debt_item.asp?debt_id=<%=Debt(0, i)%>',800,600)"><%=Debt(1, i)%></a>
                        </td>
                        <%For j = 2 To 7%>
                            <td><%=Debt(j, i)%></td>
                        <%Next%>
                    </tr>
                <%Next%>
            </table>
		</div>
	</div>
</div>	
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
