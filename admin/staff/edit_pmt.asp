<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i, j
Dim lStaffID, lEventID
Dim sStaffName, sEventName, sEventType, sComments
Dim sngPmtAmt
Dim dDatePaid, dEventDate
Dim Events()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lStaffID = Request.QueryString("staff_id")
lEventID = Request.QueryString("event_id")
sEventType = Request.QueryString("event_type")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PmtAmt, DatePaid, Comments FROM StaffPmt WHERE EventID = " & lEventID & " AND StaffID = " & lStaffID
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("pmt_amt")
    If Request.Form.Item("date_paid") & "" = "" Then
        rs(1).Value = "1/1/1900"
    Else
        rs(1).Value = Request.Form.Item("date_paid")
    End If
    If Request.Form.Item("comments") & "" = "" Then
        rs(2).Value = NULL
    Else
        rs(2).Value = Replace(Request.Form.Item("comments"), "'", "''")
    End If
    rs.Update
    rs.Close
    Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset") 
sql = "SELECT PmtAmt, DatePaid, Comments FROM StaffPmt WHERE StaffID = " & lStaffID & " AND EventID = " & lEventID
rs.Open sql, conn, 1, 2
sngPmtAmt = rs(0).Value
If Not rs(1).Value = "1/1/1900" Then dDatePaid = rs(1).Value
If Not rs(2).Value & "" = "" Then sComments = Replace(rs(2).Value, "''", "'")
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName FROM Staff WHERE StaffID = " & lStaffID
rs.Open sql, conn, 1, 2
sStaffName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
If sEventType = "fitness" Then
    sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
Else
    sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lEventID
    rs.Open sql, conn2, 1, 2
End If
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing
%>
<html lang="en">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>GSE Edit Staff Payment</title>
<!--#include file = "../../includes/meta2.asp" -->



<style type="text/css">
    th {
        text-align: right;
        padding-left: 10px;
    }
</style>
</head>

<body>
<div style="margin: 10px;padding: 10px;background-color: #fff;">
    <h3>GSE&copy;Edit Staff Payments for <%=sStaffName%></h3>

    <form name="staff_payments" method="post" action="edit_pmt.asp?event_id=<%=lEventID%>&amp;staff_id=<%=lStaffID%>&amp;event_type=<%=sEventType%>">
    <table>
        <tr>
            <td style="text-align: center;" colspan="8">
                <input type="hidden" name="submit_this" id="submit_this" value="submit_this">
                <input type="submit" name="submit1" id="submit1" value="Edit Payment">
            </td>
        </tr>
        <tr>
            <th>Date:</th>
            <td><%=dEventDate%></td>
            <th>Event:</th>
            <td><%=sEventName%></td>
            <th>Payment:</th>
            <td><input type="text" name="pmt_amt" id="pmt_amt" value="<%=sngPmtAmt%>" size="5"></td>
            <th>Date Paid:</th>
            <td><input type="text" name="date_paid" id="date_paid" value="<%=dDatePaid%>" size="5"></td>
        </tr>
        <tr>
            <th>Comments:</th>            
            <td colspan="7"><textarea name="comments" id="comments" rows="2" cols="75" style="font-size:1.1em;"><%=sComments%></textarea></td>
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
