<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j
Dim iYear
Dim sShowWhat
Dim Staff(), Events()
Dim dEndDate, dBegDate

If Session("role") & "" = "" Then Response.Redirect("http://www.google.com")

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

dEndDate = Date
dBegDate = "1/1/" & Year(Date)

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_dates") = "submit_dates" Then
    dEndDate = Request.Form.Item("end_date")
    dBegDate = REquest.Form.Item("beg_date")
    sShowWhat = Request.Form.Item("show_what")
End If

If sShowWhat = vbNullString Then sShowWhat = "All"

Call GetEvents()

i = 0
ReDim Staff(3, 0)
sql = "SELECT StaffID, FirstName, LastName FROM Staff WHERE Active = 'y' AND StaffID NOT IN (1, 2) ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Left(Replace(rs(1).Value, "''","'"), 5) & " " & Left(Replace(rs(2).Value, "''", "'"), 1)
    Staff(2, i) = 0 'availability total
    Staff(3, i) = 0 'assigned total
	i = i + 1
	ReDim Preserve Staff(3, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub GetEvents()
    Dim x, y, z
    Dim sEvntType
    Dim TempArr(3)

    x = 0
    ReDim Events(3, 0)
    For y = 0 To 2
        Set rs = Server.CreateObject("ADODB.Recordset")

        sql = vbNullString

        Select Case CInt(y)
            Case 0
                If sShowWhat = "All" or sShowWhat = "Fitness Event" Then
                    sEvntType = "Fitness Event"
                    sql = "SELECT EventID, EventName, EventDate FROM Events ORDER BY EventDate, EventName"
                    rs.Open sql, conn, 1, 2
                End If
            Case 1
                If sShowWhat = "All" or sShowWhat = "Cross-Country" Then
                    sEvntType = "Cross-Country"
                    sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE Sport = 'Cross-Country' ORDER BY MeetDate, MeetName"
                    rs.Open sql, conn2, 1, 2
                End If
            Case 2
                If sShowWhat = "All" or sShowWhat = "Nordic Ski" Then
                    sEvntType = "Nordic Ski"
                    sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE Sport = 'Nordic Ski' ORDER BY MeetDate, MeetName"
                    rs.Open sql, conn2, 1, 2
                End If
        End Select

        If Not sql = vbNullString Then
            Do While Not rs.EOF
                If CDate(rs(2).Value) >= CDate(dBegDate) Then
                    If CDate(rs(2).Value) <= CDate(dEndDate) Then
	                    Events(0, x) = rs(0).Value
	                    Events(1, x) = Replace(rs(1).Value, "''","'")
	                    Events(2, x) = rs(2).Value
                        Events(3, x) = sEvntType
	                    x = x + 1
	                    ReDim Preserve Events(3, x)
                    End If
                End If
	            rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
    Next

    For x = 0 To UBound(Events, 2) - 2
        For y = x + 1 To UBound(Events, 2) - 1
            If CDate(Events(2, x)) > CDate(Events(2, y)) THen
                For z = 0 To 3
                    TempArr(z) = Events(z, x)
                    Events(z, x) = Events(z, y)
                    Events(z, y) = TempArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Function GetAvail(lEventID, lStaffID)
    Dim x

    GetAvail = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Availability FROM StaffAvail WHERE EventID = " & lEventID & " AND StaffID = " & lStaffID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        Select Case rs(0).Value
            Case "Want To Do"
                GetAvail = "3"
            Case "Can Do"
                GetAvail = "2"
            Case "Will Do"
                GetAvail = "1"
        End Select
    Else
        GetAvail = "-"
    End If
    rs.Close
    Set rs = Nothing

    For x = 0 To UBound(Staff, 2) - 1
        If CLng(lStaffID) = CLng(Staff(0, x)) AND GetAvail <> "-" Then
            If CInt(GetAvail) = "3" Then Staff(2, x) = CInt(Staff(2, x)) + 3
            Exit For
        End If
    Next
End Function

Private Function GetActual(lEventID, lStaffID)
    Dim x
    Dim iMyAvail

    iMyAvail = 0
    GetActual = 0

    'first get the points based on their status
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Availability FROM StaffAvail WHERE EventID = " & lEventID & " AND StaffID = " & lStaffID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        Select Case rs(0).Value
            Case "Want To Do"
                iMyAvail = "3"
            Case "Can Do"
                iMyAvail = "2"
            Case "Will Do"
                iMyAvail = "1"
        End Select
    End If
    rs.Close
    Set rs = Nothing

    'now get their assignment value if they are assigned
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Role FROM StaffAsgmt WHERE EventID = " & lEventID & " AND StaffID = " & lStaffID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetActual = "3"
    rs.Close
    Set rs = Nothing

    For x = 0 To UBound(Staff, 2) - 1
        If CLng(lStaffID) = CLng(Staff(0, x)) Then
            Staff(3, x) = CInt(Staff(3, x)) + CInt(GetActual)
            Exit For
        End If
    Next
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Staff Event Assignment</title>
<!--#include file = "../../includes/js.asp" -->

<script>
$(function() {
    $( "#beg_date" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#end_date" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
        <!--#include file = "../../includes/admin_menu.asp" -->

		<div class="col-md-10">
			<h4 class="h4">GSE&copy; Event Assignments for <%=iYear%></h4>

            <form class="form-inline" name="date_range" method="post" action="event_matrix.asp">
            <div class="form-group">
                <label for="beg_date">From:</label>
                <input type="text" class="form-control" name="beg_date" id="beg_date" value="<%=dBegDate%>">
                <label for="end_date">To:</label>
                <input type="text" class="form-control" name="end_date" id="end_date" value="<%=dEndDate%>">
                <label for="show_what">Show What:</label>
                <select class="form-control" name="show_what" id="show_what" onchange="this.form.submit1.click();">
                    <option value="All">All</option>
                    <%Select Case sShowWhat%>
                        <%Case "Fitness Event"%>
                            <option value="Fitness Event" selected>Fitness Event</option>
                            <option value="Cross-Country">Cross-Country</option>
                            <option value="Nordic Ski">Nordic Ski</option>
                        <%Case "Cross-Country"%>
                            <option value="Fitness Event">Fitness Event</option>
                            <option value="Cross-Country" selected>Cross-Country</option>
                            <option value="Nordic Ski">Nordic Ski</option>
                        <%Case "Nordic Ski"%>
                            <option value="Fitness Event">Fitness Event</option>
                            <option value="Cross-Country">Cross-Country</option>
                            <option value="Nordic Ski" selected>Nordic Ski</option>
                        <%Case Else%>
                            <option value="Fitness Event">Fitness Event</option>
                            <option value="Cross-Country">Cross-Country</option>
                            <option value="Nordic Ski">Nordic Ski</option>
                    <%End Select%>
                </select>
                <input type="hidden" name="submit_dates" id="submit_dates" value="submit_dates">
                <input type="submit" class="form-control" name="submit1" id="submit1" value="Set Date Range">
            </div>
            </form>

            <table class="table table-striped table-condensed table-bordered table-responsive">
                <tr>
                    <th rowspan="2">No.</th>
                    <th rowspan="2">Event</th>
                    <th rowspan="2">Date</th>
                    <%For i = 0 To UBound(Staff, 2) - 1%>
                        <th colspan="2"><%=Staff(1, i)%></th>
                    <%Next%>
                </tr>
                <tr>
                    <%For i = 0 To UBound(Staff, 2) - 1%>
                        <th class="text-warning">Avail</th>
                        <th class="text-danger">Asgnd</th>
                    <%Next%>
                </tr>
                <%For i = 0 To UBound(Events, 2) - 1%>
                    <tr>
                        <td><%=i + 1%>)</td>
                        <td style="white-space: nowrap;" valign="top">
                            <a href="javascript:pop('edit_asgmts.asp?event_id=<%=Events(0, i)%>&amp;event_type=<%=Events(3, i)%>',1000,700)"><%=Events(1, i)%></a>
                        </td>
                        <td><%=Events(2, i)%></td>
                        <%For j = 0 To UBound(Staff, 2) - 1%>
                            <td class="text-warning"><%=GetAvail(Events(0, i), Staff(0, j))%></td>
                            <td class="text-danger"><%=GetActual(Events(0, i), Staff(0, j))%></td>
                        <%Next%>
                    </tr>
                <%Next%>
                <tr>
                    <th colspan="3">Totals:</th>
                    <%For j = 0 To UBound(Staff, 2) - 1%>
                        <th class="text-warning"><%=Staff(2, j)%></th>
                        <th class="text-danger"><%=Staff(3, j)%></th>
                    <%Next%>
                </tr>
                <tr>
                    <th colspan="3">Pct:</th>
                    <%For j = 0 To UBound(Staff, 2) - 1%>
                        <th class="text-danger" colspan="2">
                            <%If CInt(Staff(2, j)) ="0" Then%>
                                --
                            <%Else%>
                                <%=Round(CInt(Staff(3, j))/CInt(Staff(2, j)), 2)*100%>%
                            <%End If%>
                        </th>
                    <%Next%>
                </tr>
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
