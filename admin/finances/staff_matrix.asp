<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim iYear, iMyEvnts
Dim sngMyAmt, sngStaffTtl
Dim Staff(), Events(), StaffView(), SortArr(4)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = REquest.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

sngStaffTtl = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(2, 0)
sql = "SELECT StaffID, FirstName, LastName, Active FROM Staff WHERE Active = 'y' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    Staff(2, i) = rs(3).Value
	i = i + 1
	ReDim Preserve Staff(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim Events(4, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '1/1/" & iYear & "' AND EventDate <= '" & Date & "' ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'")
    Events(2, i) = rs(2).Value
    Events(3, i) = "Fitness Event"
    Events(4, i) = "0"
    i = i + 1
    ReDim Preserve Events(4, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '" & Date & "' ORDER BY MeetDate"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'")
    Events(2, i) = rs(2).Value
    Events(3, i) = rs(3).Value
    Events(4, i) = "0"
    i = i + 1
    ReDim Preserve Events(4, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'order by date
For i = 0 To UBound(Events, 2) - 2
    For j = i + 1 To UBound(Events, 2) - 1
        If CDate(Events(2, i)) > CDate(Events(2, j)) Then
            For k = 0 To 4
                SortArr(k) = Events(k, i)
                Events(k, i) = Events(k, j)
                Events(k, j) = SortArr(k)
            Next
        End If
    Next
Next

ReDim StaffView(2, 0)

If Request.Form.Item("staff_view") = "staff_view" Then
    j = 0
    For i = 0 To UBound(Staff, 2) - 1
        If Request.Form.Item("view_" & Staff(0, i)) = "on" Then
            StaffView(0, j) = Staff(0, i)
            StaffView(1, j) = Staff(1, i)
            StaffView(2, j) = "0"
            j = j + 1
            ReDim Preserve StaffView(2, j)
        End If
    Next
End If

If UBound(StaffView, 2) = 0 Then
    j = 0
    For i = 0 To UBound(Staff, 2) - 1
        If Staff(2, i) = "y" Then
            StaffView(0, j) = Staff(0, i)
            StaffView(1, j) = Staff(1, i)
            StaffView(2, j) = "0"
            j = j + 1
            ReDim Preserve StaffView(2, j)
        End If
    Next
End If

Private Sub EventStaffAmt(lThisEvent, lThisStaff, sThisSport)
    Dim x

    sngMyAmt = 0
    iMyEvnts = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TransAmt FROM FinanceStaff WHERE EventID = " & lThisEvent & " AND StaffID = " & lThisStaff & " AND Sport = '" & sThisSport & "' "
    sql = sql & "AND TransType IN ('Timing', 'Race Prep', 'Mileage')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngMyAmt = CSng(sngMyAmt) + CSng(rs(0).Value)
        iMyEvnts = CInt(iMyEvnts) + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    For x = 0 To UBound(Events, 2) - 1
        If CLng(Events(0, x)) = CLng(lThisEvent) Then
            Events(4, x) = CSng(Events(4, x)) + CSng(sngMyAmt)
            Exit For
        End If
    Next
    
    For x = 0 To UBound(StaffView, 2) - 1
        If CLng(StaffView(0, x)) = CLng(lThisStaff) Then
            StaffView(2, x) = CSng(StaffView(2, x)) + CSng(sngMyAmt)
            Exit For
        End If
    Next
End Sub

Private Function MyEvnts(lThisStaff)
    MyEvnts = "0"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TransAmt FROM FinanceStaff WHERE StaffID = " & lThisStaff & " AND TransType = 'Timing' AND TransDate >= '1/1/" & iYear 
    sql = sql & "' AND TransDate <= '" & Date & "' AND TransAmt > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then MyEvnts = rs.RecordCount
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Staff Matrix</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <!--#include file = "staff_nav.asp" -->

		    <h3 class="h3">GSE Finances: Staff Matrix</h3>

            <ul class="nav">
                <%For i = 2015 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="staff_matrix.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
           </ul>

            <h4 class="h4">Select Staff To View</h4>
            <form class="form" name="get_staff" method="post" action="staff_matrix.asp">
            <div class="bg-danger">
                <%For i = 0 To UBound(Staff, 2) - 1%>
                    <%If Staff(2, i) = "y" Then%>
                        <input type="checkbox" name="staff_view_<%=Staff(0, i)%>" id="staff_view_<%=Staff(0, i)%>" checked>&nbsp;<%=Staff(1, i)%>&nbsp;&nbsp;&nbsp;
                    <%Else%>
                        <input type="checkbox" name="staff_view_<%=Staff(0, i)%>" id="staff_view_<%=Staff(0, i)%>">&nbsp;<%=Staff(1, i)%>&nbsp;&nbsp;&nbsp;
                    <%End If%>

                    <%If i = 6 Then%>
                        <br>
                    <%End If%>
                <%Next%>
                <br>
                <input type="hidden" name="staff_view" id="staff_view" value="staff_view">
                <input class="form-control" type="submit" name="submit1" id="submit1" value="View Staff">
            </div>
            </form>

            <table class="table table-striped">
                <tr>
                    <th>No.</th>
                    <th>Event/Meet (Date)</th>
                    <%For i = 0 To UBound(StaffView, 2) - 1%>
                        <th><%=StaffView(1, i)%></th>
                    <%Next%>
                    <th>Event Pmts</th>
                </tr>
                <%For j = 0 To UBound(Events, 2) - 1%>
                    <tr>
                        <td><%=j + 1%></td>
                        <td><%=Events(1, j)%> (<%=Events(2, j)%>)</td>
                        <%For k = 0 To UBound(StaffView, 2) - 1%>
                            <%Call EventStaffAmt(Events(0, j), StaffView(0, k), Events(3, j))%>
                            <td style="text-align: right;">$<%=sngMyAmt%></td>
                        <%Next%>
                        <th style="text-align: right;">$<%=Events(4, j)%></th>
                    </tr>
                <%Next%>
                <tr>
                    <th style="text-align: right;" colspan="2">Staff Pmts</th>
                    <%For k = 0 To UBound(StaffView, 2) - 1%>
                        <th style="text-align: right;">$<%=StaffView(2, k)%></th>
                        <%sngStaffTtl = CSNg(sngStaffTtl) + CSng(StaffView(2, k))%>
                    <%Next%>
                    <th style="text-align: right;color: red;">$<%=sngStaffTtl%></th>
                </tr>
                <tr>
                    <th style="text-align: right;" colspan="2">Evnts</th>
                    <%For k = 0 To UBound(StaffView, 2) - 1%>
                        <th style="text-align: right;"><%=MyEvnts(StaffView(0, k))%></th>
                    <%Next%>
                    <th style="text-align: right;color: red;">&nbsp;</th>
                </tr>
            </table>
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
