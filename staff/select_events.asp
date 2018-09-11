<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j, k
Dim sStaffName, sEvntType, sEventName, sLocation, sTimingMethod
Dim AvailEvnts(), TempArr(), Status(3), MyAsgmts(), MyAvail(), Delete(), AllEvnts()
Dim dEventDate

If Not Session("role") = "staff" Then Response.Redirect "/default.asp?sign_out=y"

Status(0) = "Want To Do"
Status(1) = "Will Do"
Status(2) = "Can Do"
Status(3) = "Can Not Do"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName FROM Staff WHERE StaffID = " & Session("staff_id")
rs.Open sql, conn, 1, 2
sStaffName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

ReDim AllEvnts(6, 0)
j = 0
For i = 0 To 2
    Set rs = Server.CreateObject("ADODB.Recordset")

    Select Case CInt(i)
        Case 0
            sEvntType = "Fitness Event"
            sql = "SELECT EventID, EventName, EventDate, Location, TimingMethod FROM Events WHERE EventDate > '" & Date & "' "
            sql = sql & "ORDER BY EventDate, EventName"
            rs.Open sql, conn, 1, 2
        Case 1
            sEvntType = "Cross-Country"
            sql = "SELECT MeetsID, MeetName, MeetDate, MeetSite, TimingMethod FROM Meets WHERE MeetDate > '" & Date & "' "
            sql = sql & "AND Sport = 'Cross-Country' ORDER BY MeetDate, MeetName"
            rs.Open sql, conn2, 1, 2
        Case 2
            sEvntType = "Nordic Ski"
            sql = "SELECT MeetsID, MeetName, MeetDate, MeetSite, TimingMethod FROM Meets WHERE MeetDate > '" & Date & "'  "
            sql = sql & "AND Sport = 'Nordic Ski' ORDER BY MeetDate, MeetName"
            rs.Open sql, conn2, 1, 2
    End Select

    Do While Not rs.EOF
        AllEvnts(0, j) = rs(0).Value
        AllEvnts(1, j) = Left(Replace(rs(1).Value, "''","'"), 35)
        AllEvnts(2, j) = rs(2).Value
        AllEvnts(3, j) = rs(3).Value
        AllEvnts(4, j) = sEvntType
        AllEvnts(5, j) = rs(4).Value
        AllEvnts(6, j) = GetFirstRace(rs(0).Value, sEvntType)
        j = j + 1
        ReDim Preserve AllEvnts(6, j)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
Next

ReDim TempArr(6)
For i = 0 To UBound(AllEvnts, 2) - 2
    For j = i + 1 To UBound(AllEvnts, 2) - 1
        If CDate(AllEvnts(2, i)) > CDate(AllEvnts(2, j)) THen
            For k = 0 To 6
                TempArr(k) = AllEvnts(k, i)
                AllEvnts(k, i) = AllEvnts(k, j)
                AllEvnts(k, j) = TempArr(k)
            Next
        End If
    Next
Next

i = 0
ReDim MyAsgmts(10, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventType, Role, Amount, DatePaid, Comments FROM StaffAsgmt WHERE StaffID = " & Session("staff_id")
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Call GetEventInfo(rs(0).Value, rs(1).Value)

    If Not CStr(dEventDate) = vbNullString Then
        If CDAte(dEventDate) >= Date Then
	        MyAsgmts(0, i) = rs(0).Value
	        MyAsgmts(1, i) = Left(sEventName, 35)
	        MyAsgmts(2, i) = dEventDate
            MyAsgmts(3, i) = sLocation
            MyAsgmts(4, i) = rs(1).Value
            MyAsgmts(5, i) = sTimingMethod
            MyAsgmts(6, i) = rs(2).Value
            MyAsgmts(7, i) = rs(3).Value
            If Not rs(4).Value = "1/1/1900" Then MyAsgmts(8, i) = rs(4).Value
            MyAsgmts(9, i) = GetAvail(rs(0).Value, rs(1).Value)
            If Not rs(5).Value & "" = "" Then  MyAsgmts(10, i) = Replace(rs(5).Value, "''", "'")
	        i = i + 1
	        ReDim Preserve MyAsgmts(10, i)
        End If
    End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

ReDim TempArr(10)
For i = 0 To UBound(MyAsgmts, 2) - 2
    For j = i + 1 To UBound(MyAsgmts, 2) - 1
        If CDate(MyAsgmts(2, i)) > CDate(MyAsgmts(2, j)) THen
            For k = 0 To 10
                TempArr(k) = MyAsgmts(k, i)
                MyAsgmts(k, i) = MyAsgmts(k, j)
                MyAsgmts(k, j) = TempArr(k)
            Next
        End If
    Next
Next

If Request.Form.Item("submit_edit") = "submit_edit" Then 
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT StaffAvailID, Availability, Comments FROM StaffAvail WHERE StaffID = " & Session("staff_id")
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("edit_status_" & rs(0).Value) & "" = "" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
            rs(1).Value = Request.Form.Item("edit_status_" & rs(0).Value)
            If Request.Form.Item("comments_" & rs(0).Value) & "" = "" Then
                rs(2).Value = NULL
            Else
                rs(2).Value = Replace(Request.Form.Item("comments_" & rs(0).Value), "'", "''")
            End If
            rs.Update
        End If

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_avail") = "submit_avail" Then 
    Call GetMyAvail()
    Call GetAvailList()

    For i = 0 To UBound(AvailEvnts, 2) - 1
        If Not Request.Form.Item("status_" & AvailEvnts(0, i)) & "" = "" Then
            sql = "INSERT INTO StaffAvail (EventID, EventType, Availability, StaffID) VALUES (" & AvailEvnts(0, i) & ", '" & AvailEvnts(4, i)
            sql = sql & "', '" & Request.Form.Item("status_" & AvailEvnts(0, i)) & "', " & Session("staff_id") & ")"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    Next
End If

Call GetMyAvail()
Call GetAvailList()

Private Sub GetMyAvail()
    i = 0
    ReDim MyAvail(8, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventType, Availability, StaffAvailID, Comments FROM StaffAvail WHERE StaffID = "
    sql = sql & Session("staff_id")
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Call GetEventInfo(rs(0).Value, rs(1).Value)

        If Not CStr(dEventDate) = vbNullString Then
            If CDAte(dEventDate) >= Date Then
                If Assigned(rs(0).Value, rs(1).Value) = "n" Then
                    MyAvail(0, i) = rs(0).Value
                    MyAvail(1, i) = Left(sEventName, 35)
                    MyAvail(2, i) = dEventDate
                    MyAvail(3, i) = sLocation
                    MyAvail(4, i) = rs(1).Value
                    MyAvail(5, i) = sTimingMethod
                    MyAvail(6, i) = rs(2).Value
                    MyAvail(7, i) = rs(3).Value
                    If Not rs(4).Value & "" = "" Then  MyAvail(8, i) = Replace(rs(4).Value, "''", "'")
                    i = i + 1
                    ReDim Preserve MyAvail(8, i)
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ReDim TempArr(8)
    For i = 0 To UBound(MyAvail, 2) - 2
        For j = i + 1 To UBound(MyAvail, 2) - 1
            If CDate(MyAvail(2, i)) > CDate(MyAvail(2, j)) THen
                For k = 0 To 8
                    TempArr(k) = MyAvail(k, i)
                    MyAvail(k, i) = MyAvail(k, j)
                    MyAvail(k, j) = TempArr(k)
                Next
            End If
        Next
    Next
End Sub

Private Function Assigned(lEventID, sEventType)
    Assigned = "n"
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT EventID FROM StaffAsgmt WHERE StaffID = " & Session("staff_id") & " AND EventID = " & lEventID
    sql2 = sql2 & " AND EventType = '" & sEventType & "'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then Assigned = "y"
    rs2.Close
    set rs2 = Nothing
End Function

Private Sub GetEventInfo(lEventID, sThisEventType)
    sEventName = vbNullString
    dEventDate = vbNullString
    sLocation = vbNullString
    sTimingMethod = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    Select Case sThisEventType
        Case "Fitness Event"
            sql2 = "SELECT EventName, EventDate, Location, TimingMethod FROM Events WHERE EventID = " & lEventID
            rs2.Open sql2, conn, 1, 2
        Case Else
            sql2 = "SELECT MeetName, MeetDate, MeetSite, TimingMethod FROM Meets WHERE MeetsID = " & lEventID 
            rs2.Open sql2, conn2, 1, 2
    End Select
    If rs2.RecordCount > 0 Then
        sEventName = Replace(rs2(0).Value, "''", "'")
        dEventDate = rs2(1).Value
        sLocation = rs2(2).Value
        sTimingMethod = rs2(3).Value
    End If
    rs2.Close
    Set rs2 = Nothing
End Sub

Private Sub GetAvailList()
    Dim x, y

    x = 0
    ReDim AvailEvnts(6, 0)
    For y = 0 to UBound(AllEvnts, 2) - 1
        If AvailSet(AllEvnts(0, y), AllEvnts(4, y)) = "n" Then
            AvailEvnts(0, x) = AllEvnts(0, y)
            AvailEvnts(1, x) = Left(Replace(AllEvnts(1, y), "''","'"), 35)
            AvailEvnts(2, x) = AllEvnts(2, y)
            AvailEvnts(3, x) = AllEvnts(3, y)
            AvailEvnts(4, x) = AllEvnts(4, y)
            AvailEvnts(5, x) = AllEvnts(5, y)
            AvailEvnts(6, x) = AllEvnts(6, y)
            x = x + 1
            ReDim Preserve AvailEvnts(6, x)
        End If
    Next

    ReDim TempArr(6)
    For x = 0 To UBound(AvailEvnts, 2) - 2
        For y = x + 1 To UBound(AvailEvnts, 2) - 1
            If CDate(AvailEvnts(2, x)) > CDate(AvailEvnts(2, y)) THen
                For z = 0 To 6
                    TempArr(z) = AvailEvnts(z, x)
                    AvailEvnts(z, x) = AvailEvnts(z, y)
                    AvailEvnts(z, y) = TempArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Function AvailSet(lThisEventID, sThisEventType)
    Dim x

    AvailSet = "n"

    For x = 0 To UBound(MyAvail, 2) - 1
        If (CLng(MyAvail(0, x)) = CLng(lThisEventID)) AND (CStr(MyAvail(4, x)) = CStr(sThisEventType)) Then
            AvailSet = "y"
            Exit For
        End If
    Next

    If AvailSet = "n" Then
        For x = 0 To UBound(MyAsgmts, 2) - 1
            If (CLng(MyAsgmts(0, x)) = CLng(lThisEventID)) AND (CStr(MyAsgmts(4, x)) = CStr(sThisEventType)) Then
                AvailSet = "y"
                Exit For
            End If
        Next
    End If
End Function

Private Function IsAvail(lEventID, sEventType)
    Dim iNumStaff

    IsAvail = "y"

    'see if this person has already assigned for this one
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT StaffID FROM StaffAsgmt WHERE EventID = " & lEventID & " AND EventType = '" & sEventType & "' AND StaffID = " & Session("staff_id")
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then IsAvail = "n"
    rs2.Close
    Set rs2 = Nothing

    If IsAvail = "y" Then
        'see if this person has already claimed availability for this one
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT StaffID FROM StaffAvail WHERE EventID = " & lEventID & " AND EventType = '" & sEventType & "' AND StaffID = " & Session("staff_id")
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then IsAvail = "n"
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

Private Function GetAvail(lThisEvent, sThisType)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Availability FROM StaffAvail WHERE StaffID = " & Session("staff_id") & " AND EventID = " & lThisEvent & " AND EventType = '" 
    sql2 = sql2 & sThisType & "'"
    rs2.Open sql2, conn, 1, 2
    GetAvail = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetFirstRace(lThisEvent, sThisType)
    GetFirstRace = "tbd"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    Select Case sThisType
        Case "Fitness Event"
            sql2 = "SELECT StartTime FROM RaceData WHERE EventID = " & lThisEvent & " ORDER BY StartTime"
            rs2.Open sql2, conn, 1, 2
        Case Else
            sql2 = "SELECT RaceTime FROM Races WHERE MeetsID = " & lThisEvent & " ORDER BY RaceTime"
            rs2.Open sql2, conn2, 1, 2
    End Select
    If rs2.RecordCount > 0 Then GetFirstRace = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Staff History</title>
<meta name="description" content="Gopher State Events staff profile page.">
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
  	<div class="row">
		<!--#include file = "staff_menu.asp" -->
		<div class="col-sm-10">
			<h3 class="h3">GSE Staff History For <%=sStaffName%></h3>

            <div>
                <p>Please look at all events on our schedule and make yourself available, or not, to varying degrees.  The degrees are:</p>
                <ol class="list-group">
                    <li class="list-group-item"><span style="font-weight: bold;">"Want To Do":</span> This is a prefered event for you.</li>
                    <li class="list-group-item"><span style="font-weight: bold;">"Can Do":</span> You are available and are willing to take this event.</li>
                    <li class="list-group-item"><span style="font-weight: bold;">"Will Do":</span> You really would rather not but if we are in a jam you will help out.</li>
                    <li class="list-group-item"><span style="font-weight: bold;">"Can Not Do":</span> You are unavailable for that event.</li>
                </ol>
            </div>

            <div>
                <h4 class="h4">My Assignments</h4>
                <table class="table table-striped">
                    <tr>
                        <th>No.</th>
                        <th>Event</th>
                        <th>Date</th>
                        <th>Location</th>
                        <th>Type</th>
                        <th>Timing</th>
                        <th>Role</th>
                        <th>Admin Comments</th>
                    </tr>
                    <%For i = 0 To UBound(MyAsgmts, 2) - 1%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td><a href="javascript:pop('/events/raceware_events.asp?event_id=<%=MyAsgmts(0, i)%>',800,600)"><%=MyAsgmts(1, i)%></a></td>
                            <td><%=MyAsgmts(2, i)%></td>
                            <td><%=MyAsgmts(3, i)%></td>
                            <td><%=MyAsgmts(4, i)%></td>
                            <td><%=MyAsgmts(5, i)%></td>
                            <td><%=MyAsgmts(6, i)%></td>
                            <td><%=MyAsgmts(10, i)%></td>
                        </tr>
                    <%Next%>
                </table>
            </div>

            <div class="bg-danger" style="padding: 5px;">
                <h4 class="h4">My Availability</h4>
                <form class="form" name="edit_avail" method="post" action="select_events.asp">
                <input type="hidden" name="submit_edit" id="submit_edit" value="submit_edit">
                <input class="form-control" type="submit" name="submit1" id="submit1" value="Submit Changes">
                <table class="table">
                    <tr>
                        <th>No.</th>
                        <th>Event</th>
                        <th>Date</th>
                        <th>Location</th>
                        <th>Type</th>
                        <th>My Status</th>
                    </tr>
                    <%For i = 0 To UBound(MyAvail, 2) - 1%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td><%=MyAvail(1, i)%></td>
                            <td><%=MyAvail(2, i)%></td>
                            <td><%=MyAvail(3, i)%></td>
                            <td><%=MyAvail(4, i)%></td>
                            <td>
                                <select class="form-control" name="edit_status_<%=MyAvail(7, i)%>" id="edit_status_<%=MyAvail(7, i)%>">
                                    <option value="">&nbsp;</option>
                                    <%For j = 0 To UBound(Status)%>
                                        <%If CStr(MyAvail(6, i)) = CStr(Status(j)) Then%>
                                            <option value="<%=Status(j)%>" selected><%=Status(j)%></option>
                                        <%Else%>
                                            <option value="<%=Status(j)%>"><%=Status(j)%></option>
                                        <%End If%>
                                    <%Next%>
                                </select>
                            </td>
                        </tr>
                    <%Next%>
                </table>
                </form>
            </div>

            <div>
            <h4 class="h4">Available Events</h4>
                <form class="form" name="claim_events" method="post" action="select_events.asp">
                <input type="hidden" name="submit_avail" id="submit_avail" value="submit_avail">
                <input class="form-control" type="submit" name="submit2" id="submit2" value="Submit Availability">
                <table class="table table-striped">
                    <tr>
                        <th>No.</th>
                        <th>Event</th>
                        <th>Date</th>
                        <th>First Race</th>
                        <th>Location</th>
                        <th>Type</th>
                        <th>My Status</th>
                    </tr>
                    <%For i = 0 To UBound(AvailEvnts, 2) - 1%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td><%=AvailEvnts(1, i)%></td>
                            <td><%=AvailEvnts(2, i)%></td>
                            <td><%=AvailEvnts(6, i)%></td>
                            <td><%=AvailEvnts(3, i)%></td>
                            <td><%=AvailEvnts(4, i)%></td>
                            <td>
                                <select class="form-control" name="status_<%=AvailEvnts(0, i)%>" id="status_<%=AvailEvnts(0, i)%>">
                                    <option value="">&nbsp;</option>
                                    <%For j = 0 To UBound(Status)%>
                                        <option value="<%=Status(j)%>"><%=Status(j)%></option>
                                    <%Next%>
                                </select>
                            </td>
                        </tr>
                    <%Next%>
                </table>
                </form>
            </div>
		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>