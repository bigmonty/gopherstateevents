<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim lStaffID, lEventID, lFinanceStaffID
Dim sStaffName, sTransType, sSport, sPmtMethod, sComments, sDelete
Dim iCheckNum
Dim iYear
Dim sngTransAmt
Dim Events(), SortArr(3), TransTypes(5), PmtMethods(3), Sports(2)
Dim dTransDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lFinanceStaffID = Request.QueryString("finance_staff_id")

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Sports(0) = "Fitness Event"
Sports(1) = "Nordic Ski"
Sports(2) = "Cross-Country"

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

If Request.Form.Item("submit_changes") = "submit_changes" Then
    sTransType = Request.Form.Item("trans_type")
    sngTransAmt = Request.Form.Item("trans_amt")
    lEventID = Request.Form.Item("event_id")
    dTransDate = Request.Form.Item("trans_date")
    sSport = Request.Form.Item("sport")
    sPmtMethod = Request.Form.Item("pmt_method")
    iCheckNum = Request.Form.Item("check_num")
    If Not Request.Form.Item("comments") = vbNullString Then sComments = Request.Form.Item("comments")
    sDelete = Request.Form.Item("delete")

    If sDelete = "on" Then
        sql = "DELETE FROM FinanceStaff WHERE FinanceStaffID = " & lFinanceStaffID
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT TransType, TransAmt, TransDate, EventID, Sport, PmtMethod, CheckNum, Comments FROM FinanceStaff WHERE FinanceStaffID = " 
        sql = sql & lFinanceStaffID
        rs.Open sql, conn, 1, 2
        rs(0).Value = sTransType
        rs(1).Value = sngTransAmt
        rs(2).Value = dTransDate
        If Not CStr(lEventID) & "" = "" Then rs(3).Value = lEventID
        rs(4).Value = sSport
        rs(5).Value = sPmtMethod
        rs(6).Value = iCheckNum
        rs(7).Value = sComments
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
End If

If CStr(lStaffID) = vbNullString Then lStaffID = "0"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT StaffID, TransType, TransAmt, TransDate, EventID, Sport, PmtMethod, CheckNum, Comments FROM FinanceStaff WHERE FinanceStaffID = " 
sql = sql & lFinanceStaffID
rs.Open sql, conn, 1, 2
lStaffID = rs(0).Value
sTransType = rs(1).Value
sngTransAmt = rs(2).Value
dTransDate = rs(3).Value
lEventID = rs(4).Value
sSport = rs(5).Value
sPmtMethod = rs(6).Value
iCheckNum = rs(7).Value
If Not rs(8).Value & "" = "" Then sComments = Replace(rs(8).Value, "''", "'")
rs.Close
Set rs = Nothing

'get staff name
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName FROM Staff WHERE StaffID = " & lStaffID
rs.Open sql, conn, 1, 2
sStaffName = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing

Private Function GetEventName(lThisEvent, sThisSport)
    Set rs = Server.CreateObject("ADODB.Recordset")

    If sThisSport = "Fitness Event" Then
        sql = "SELECT EventName FROM Events WHERE EventID = " & lThisEvent
        rs.Open sql, conn, 1, 2
    Else
        sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lThisEvent
        rs.Open sql, conn2, 1, 2
    End If

    GetEventName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Edit Staff Finance</title>
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
    <h3>Edit Staff Finance Record for <%=sStaffName%></h3>

    <div>
        <form name="enter_data" method="post" action="edit_staff.asp?year=<%=iYear%>&amp;finance_staff_id=<%=lFinanceStaffID%>">
        <table>
            <tr>
                <th>Amount:</th>
                <td>$<input type="text" name="trans_amt" id="trans_amt" size="3" value="<%=sngTransAmt%>"></td>
                <th>Date:</th>
                <td><input type="text" name="trans_date" id="trans_date" size="6" value="<%=dTransDate%>"></td>                            
                <th>Type:</th>
                <td>
                    <select name="trans_type" id="trans_type">
                        <%For i = 0 To UBound(TransTypes)%>
                            <%If CStr(TransTypes(i)) = CStr(sTransType) Then%>
                                <option value="<%=TransTypes(i)%>" selected><%=TransTypes(i)%></option>
                            <%Else%>
                                <option value="<%=TransTypes(i)%>"><%=TransTypes(i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </td>
                <th>Pmt Method:</th>
                <td>
                    <select name="pmt_method" id="pmt_method">
                        <%For i = 0 To UBound(PmtMethods)%>
                            <%If CStr(PmtMethods(i)) = CStr(sPmtMethod) Then%>
                                <option value="<%=PmtMethods(i)%>" selected><%=PmtMethods(i)%></option>
                            <%Else%>
                                <option value="<%=PmtMethods(i)%>"><%=PmtMethods(i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </td>
            </tr>
            <tr>
                <th>Check Num:</th>
                <td><input type="text" name="check_num" id="check_num" size="4" value="<%=iCheckNum%>"></td>                            
                <th>Event:</th>
                <td>
                    <select name="event_id" id="event_id">
                        <option value=""></option>
                        <%For i = 0 To UBound(Events, 2) - 1%>
                            <%If CLng(Events(0, i)) = CLng(lEVentID) Then%>
                                <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
                            <%Else%>
                                <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </td>
                <th>Sport:</th>
                <td colspan="3">
                    <select name="sport" id="sport">
                        <option value=""></option>
                        <%For i = 0 To UBound(Sports)%>
                            <%If CStr(Sports(i)) = CStr(sSport) Then%>
                                <option value="<%=Sports(i)%>" selected><%=Sports(i)%></option>
                            <%Else%>
                                <option value="<%=Sports(i)%>"><%=Sports(i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </td>
            </tr>
            <tr>  
                <th valign="top">Comments:</th>  
                <td colspan="7">
                    <textarea name="comments" id="comments" rows="2" cols="80" style="font-size: 1.1em;"><%=sComments%></textarea>
                </td>
            </tr>
            <tr>    
                <td class="alt"style="text-align: center;color: red;" colspan="8">
                    <input type="checkbox" name="delete" id="delete">&nbsp;Delete Record (There is no Undo for this action!)
                </td>
            </tr>
            <tr>    
                <td style="text-align: center;" colspan="8">
                    <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                    <input type="submit" name="submit1" id="submit1" value="Submit New Data">
                </td>
            </tr>
        </table>
        </form>
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
