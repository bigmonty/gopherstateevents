<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql, rs2, sql2
Dim i, j
Dim lEventID
Dim iNumRqd
Dim sTech, sSupport, sEventType, sEventName, sLocation, sTimingMethod
Dim Events(), Support(), Roles(3), AsgdStaff(), AvailStaff(), Delete(), OtherStaff()
Dim dEventDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Roles(0) = "None"
Roles(1) = "Tech"
Roles(2) = "Support"
Roles(3) = "Other"

lEventID = Request.QueryString("event_id")
sEventType = Request.QueryString("event_type")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
Select Case sEventType
    Case "Fitness Event"
        sql = "SELECT EventName, EventDate, Location, TimingMethod FROM Events WHERE EventID = " & lEventID
        rs.Open sql, conn, 1, 2
    Case Else
        sql = "SELECT MeetsID, MeetName, MeetDate, MeetSite, TimingMethod FROM Meets WHERE MeetsID = "  & lEventID
        rs.Open sql, conn2, 1, 2
End Select
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sLocation = rs(2).Value
sTimingMethod = rs(3).Value
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_other") = "submit_other" Then
    Call GetOtherStaff()

    For i = 0 To UBound(OtherStaff, 2) - 1
        If Not Request.Form.Item("role_" & OtherStaff(0, i)) = vbNullString Then
            'add to avail
            sql = "INSERT INTO StaffAvail (EventID, EventType, Availability, StaffID, Comments) VALUES (" & lEventID & ", '" & sEventType
            sql = sql & "', 'Will Do', " & OtherStaff(0, i) & ", 'Entered by Admin')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing

           'assign staff
            sql = "INSERT INTO StaffAsgmt (EventID, EventType, StaffID, Role) VALUES (" & lEventID & ", '" & sEventType & "', " & OtherStaff(0, i) & ", '"
            sql = sql & Request.Form.Item("role_" & OtherStaff(0, i)) & "')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    Next

    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
ElseIf Request.Form.Item("submit_num_rqd") = "submit_num_rqd" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT NumRqd FROM StaffRqd WHERE EventID = " & lEventID & " AND EventType = '" & sEventType & "'"
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("num_rqd")
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_edit") = "submit_edit" Then
    Call GetAsgdStaff()

    j = 0
    ReDim Delete(0)
    For i = 0 To UBound(AsgdStaff, 2) - 1
        If Request.Form.Item("role_" & AsgdStaff(0, i)) = vbNullString Then
            Delete(j) = AsgdStaff(0, i)
            j = j + 1
            ReDim Preserve Delete(j)
        Else
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT Role, Amount, DatePaid, Comments FROM StaffAsgmt WHERE EventID = " & lEventID & " AND EventType = '" & sEventType 
            sql = sql & "' AND StaffID = " & AsgdStaff(0, i)
            rs.Open sql, conn, 1, 2
            rs(0).Value = Request.Form.Item("role_" & AsgdStaff(0, i))
            rs(1).Value = Request.Form.Item("amount_" & AsgdStaff(0, i))
            If Not Request.Form.Item("date_paid_" & AsgdStaff(0, i)) & "" = "" Then rs(2).Value = Request.Form.Item("date_paid_" & AsgdStaff(0, i))
            If Not Request.Form.Item("comments_" & AsgdStaff(0, i)) & "" = "" Then rs(3).Value = Request.Form.Item("comments_" & AsgdStaff(0, i))
            rs.Update
            rs.Close
            Set rs = Nothing
       End If
    Next

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM StaffAsgmt WHERE EventID = " & lEventID & " AND EventType = '" & sEventType & "' AND StaffID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next

    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
ElseIf Request.Form.Item("submit_asgmt") = "submit_asgmt" Then
    Call GetAvailStaff()

    For i = 0 To UBound(AvailStaff, 2) - 1
        If Not Request.Form.Item("role_" & AvailStaff(0, i)) = vbNullString Then
            sql = "INSERT INTO StaffAsgmt (EventID, EventType, StaffID, Role) VALUES (" & lEventID & ", '" & sEventType & "', " & AvailStaff(0, i) & ", '"
            sql = sql & Request.Form.Item("role_" & AvailStaff(0, i)) & "')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    Next

    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
End If

Call GetAsgdStaff()
Call GetAvailStaff()
Call GetOtherStaff()
iNumRqd = GetNumRqd()

Private Sub GetOtherStaff()
    i = 0
    ReDim OtherStaff(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT StaffID, LastName, FirstName FROM Staff WHERE Active = 'y' ORDER BY LastName, FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Claimed(rs(0).Value) = "n" Then
            OtherStaff(0, i) = rs(0).Value
            OtherStaff(1, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
            i = i + 1
            ReDim Preserve OtherStaff(1, i)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetAsgdStaff()
    i = 0
    ReDim AsgdStaff(6, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sa.StaffID, s.LastName, s.FirstName, sa.Role, sa.Amount, sa.DatePaid, sa.Comments FROM StaffAsgmt sa INNER JOIN Staff s "
    sql = sql & "ON sa.StaffID = s.StaffID WHERE sa.EventID = " & lEventID & " AND sa.EventType = '" & sEventType & "' ORDER BY LastName, FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        AsgdStaff(0, i) = rs(0).Value
        AsgdStaff(1, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
        AsgdStaff(2, i) = rs(3).Value
        AsgdStaff(3, i) = GetAvail(rs(0).Value)
        AsgdStaff(4, i) = rs(4).Value
        If Not rs(5).Value = "1/1/1900" Then AsgdStaff(5, i) = rs(5).Value
        If Not rs(6).Value & "" = "" Then AsgdStaff(6, i) = Replace(rs(6).Value, "''", "'")
        i = i + 1
        ReDim Preserve AsgdStaff(6, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetAvailStaff()
    i = 0
    ReDim AvailStaff(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sa.StaffID, s.LastName, s.FirstName, sa.Availability, sa.Comments FROM StaffAvail sa INNER JOIN Staff s "
    sql = sql & "ON sa.StaffID = s.StaffID WHERE sa.EventID = " & lEventID & " AND sa.EventType = '" & sEventType & "' ORDER BY LastName, FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If IsAsgd(rs(0).Value) = "n" Then
            AvailStaff(0, i) = rs(0).Value
            AvailStaff(1, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
            AvailStaff(2, i) = rs(3).Value
            If Not rs(4).Value & "" = "" Then AvailStaff(3, i) = Replace(rs(4).Value, "''", "'")
            i = i + 1
            ReDim Preserve AvailStaff(3, i)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Function GetNumRqd()
    Dim bFound

    bFound = "y"
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT NumRqd FROM StaffRqd WHERE EventID = " & lEventID & " AND EventType = '" & sEventType & "'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then 
        GetNumRqd = rs2(0).Value
    Else
        GetNumRqd = "2"
        bFound = "n"
    End If
    rs2.Close
    Set rs2 = Nothing

    If bFound = "n" Then
        sql2 = "INSERT INTO StaffRqd (EventID, EventType, NumRqd) VALUES (" & lEventID & ", '" & sEventType & "', 2)"
        Set rs2 = conn.Execute(sql2)
        Set rs2 = Nothing
    End If
End Function

Private Function GetAvail(lThisStaff)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Availability FROM StaffAvail WHERE StaffID = " & lThisStaff & " AND EventID = " & lEventID & " AND EventType = '" & sEventType & "'"
    rs2.Open sql2, conn, 1, 2
    GetAvail = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function IsAsgd(lThisStaff)
    IsAsgd = "n"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Role FROM StaffAsgmt WHERE StaffID = " & lThisStaff & " AND EventID = " & lEventID & " AND EventType = '" & sEventType & "'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then isAsgd = "y"
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function Claimed(lThisStaff)
    Claimed = "n"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT StaffID FROM StaffAsgmt WHERE StaffID = " & lThisStaff & " AND EventID = " & lEventID & " AND EventType = '" & sEventType & "'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then Claimed = "y"
    rs2.Close
    Set rs2 = Nothing

    If Claimed = "n" Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT StaffID FROM StaffAvail WHERE StaffID = " & lThisStaff & " AND EventID = " & lEventID & " AND EventType = '" & sEventType & "'"
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then Claimed = "y"
        rs2.Close
        Set rs2 = Nothing
    End If
End Function
%>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Staff Event Assignment</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
     th, td {
        padding-left: 5px;
    }
    
    th {
        text-align: left;
    }
</style>
</head>

<body>
<div style="margin: 10px;padding: 10px;background-color: #fff;">
    <h3>GSE&copy;&nbsp;Edit Staff Assignments for <%=sEventName%> (<%=dEventDate%>)</h3>

    <h4 style="background: none;background-color: #fff;text-align: left;color: #000;">Location: <%=sLocation%></h4>
    <h4 style="background: none;background-color: #fff;text-align: left;;color: #000;">Timing Method: <%=sTimingMethod%></h4>

    <div style="background-color:#ececd8;">
        <form name="set_num" method="post" action = "edit_asgmts.asp?event_id=<%=lEventID%>&amp;event_type=<%=sEventType%>">
        Num Needed:
        <select name="num_rqd" id="num_rqd">
            <%For i = 1 To 10%>
                <%If CInt(iNumRqd) = CInt(i) Then%>
                    <option value="<%=i%>" selected><%=i%></option>
                <%Else%>
                    <option value="<%=i%>"><%=i%></option>
                <%End If%>
            <%Next%>
        </select>
        <input type="hidden" name="submit_num_rqd" id="submit_num_rqd" value="submit_num_rqd">
        <input type="submit" name="submit3" id="submit3" value="Set Number">
        </form>
    </div>

    <h4 class="h4">Current Assignments</h4>
    <form name="assign_staff" method="post" action="edit_asgmts.asp?event_id=<%=lEventID%>&amp;event_type=<%=sEventType%>">
    <table>
        <tr>
            <th>Staff Name</th>
            <th>Availability</th>
            <th>Role</th>
            <th>Amount</th>
            <th>Date Paid</th>
            <th>Admin Comments</th>
        </tr>
        <%For i = 0 To UBound(AsgdStaff, 2) - 1%>
            <tr>
                <td valign="top"><%=AsgdStaff(1, i)%></td>
                <td valign="top"><%=AsgdStaff(2, i)%></td>
                <td valign="top">
                    <select name="role_<%=AsgdStaff(0, i)%>" id="role_<%=AsgdStaff(0, i)%>">
                        <option value="">&nbsp;</option>
                        <%For j = 0 To UBound(Roles)%>
                            <%If CStr(AsgdStaff(2, i)) = Roles(j) Then%>
                                <option value="<%=Roles(j)%>" selected><%=Roles(j)%></option>
                            <%Else%>
                                <option value="<%=Roles(j)%>"><%=Roles(j)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </td>
                <td valign="top"><input type="text" name="amount_<%=AsgdStaff(0, i)%>" id="amount_<%=AsgdStaff(0, i)%>" value="<%=AsgdStaff(4, i)%>"></td>
                <td valign="top"><input type="text" name="date_paid_<%=AsgdStaff(0, i)%>" id="date_paid_<%=AsgdStaff(0, i)%>" value="<%=AsgdStaff(5, i)%>"></td>
                <td><textarea name="comments_<%=AsgdStaff(0, i)%>" id="comments_<%=AsgdStaff(0, i)%>" cols="40" rows="2"><%=AsgdStaff(6, i)%></textarea></td>
            </tr>
        <%Next%>
        <tr>
            <td style="text-align: center;" colspan="6">
                <input type="hidden" name="submit_edit" id="submit_edit" value="submit_edit">
                <input type="submit" name="submit1" id="submit1" value="Edit Staff">
            </td>
        </tr>
    </table>
    </form>

    <h4 class="h4">Available Staff</h4>
    <form name="assign_staff" method="post" action="edit_asgmts.asp?event_id=<%=lEventID%>&amp;event_type=<%=sEventType%>">
    <table>
        <tr>
            <th>Staff Name</th>
            <th>Availability</th>
            <th>Role</th>
            <th>User Comments</th>
        </tr>
        <%For i = 0 To UBound(AvailStaff, 2) - 1%>
            <tr>
                <td valign="top"><%=AvailStaff(1, i)%></td>
                <td valign="top"><%=AvailStaff(2, i)%></td>
                <td valign="top">
                    <select name="role_<%=AvailStaff(0, i)%>" id="role_<%=AvailStaff(0, i)%>">
                        <option value="">&nbsp;</option>
                        <%For j = 0 To UBound(Roles)%>
                            <option value="<%=Roles(j)%>"><%=Roles(j)%></option>
                        <%Next%>
                    </select>
                </td>
                <td><%=AvailStaff(3, i)%></td>
            </tr>
        <%Next%>
        <tr>
            <td style="text-align: center;" colspan="4">
                <input type="hidden" name="submit_asgmt" id="submit_asgmt" value="submit_asgmt">
                <input type="submit" name="submit2" id="submit2" value="Assign Staff">
            </td>
        </tr>
    </table>
    </form>

    <h4 class="h4">Other Staff</h4>
    <form name="other_staff" method="post" action="edit_asgmts.asp?event_id=<%=lEventID%>&amp;event_type=<%=sEventType%>">
    <table>
        <tr>
            <th>Staff Name</th>
            <th><th>Role</th></th>
        </tr>
        <%For i = 0 To UBound(OtherStaff, 2) - 1%>
            <tr>
                <td valign="top"><%=OtherStaff(1, i)%></td>
                <td valign="top">
                    <select name="role_<%=OtherStaff(0, i)%>" id="role_<%=OtherStaff(0, i)%>">
                        <option value="">&nbsp;</option>
                        <%For j = 0 To UBound(Roles)%>
                            <option value="<%=Roles(j)%>"><%=Roles(j)%></option>
                        <%Next%>
                    </select>
                </td>
            </tr>
        <%Next%>
        <tr>
            <td style="text-align: center;" colspan="2">
                <input type="hidden" name="submit_other" id="submit_other" value="submit_other">
                <input type="submit" name="submit4" id="submit4" value="Assign Other Staff">
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
