<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j
Dim iYear
Dim Staff(), Events()

If Session("role") & "" = "" Then Response.Redirect("http://www.google.com")

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Call GetEvents()

i = 0
ReDim Staff(1, 0)
sql = "SELECT StaffID, FirstName, LastName FROM Staff ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(1).Value, "''","'") & " " & Replace(rs(2).Value, "''", "'")
	i = i + 1
	ReDim Preserve Staff(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub GetEvents()
    Dim x, y, z
    Dim sEvntType, sThisColor
    Dim dLastDate
    Dim TempArr(8)

    x = 0
    ReDim Events(8, 0)
    For y = 0 To 2
        Set rs = Server.CreateObject("ADODB.Recordset")

        Select Case CInt(y)
            Case 0
                sEvntType = "Fitness Event"
                sql = "SELECT EventID, EventName, EventDate, Location, TimingMethod FROM Events ORDER BY EventDate, EventName"
                rs.Open sql, conn, 1, 2
            Case 1
                sEvntType = "Cross-Country"
                sql = "SELECT MeetsID, MeetName, MeetDate, MeetSite, TimingMethod FROM Meets WHERE Sport = 'Cross-Country' ORDER BY MeetDate, MeetName"
                rs.Open sql, conn2, 1, 2
            Case 2
                sEvntType = "Nordic Ski"
                sql = "SELECT MeetsID, MeetName, MeetDate, MeetSite, TimingMethod FROM Meets WHERE Sport = 'Nordic Ski' ORDER BY MeetDate, MeetName"
                rs.Open sql, conn2, 1, 2
        End Select

        Do While Not rs.EOF
            If Year(rs(2).Value) = CInt(iYear) Then
                If x = 0 Then
                    dLastDate = rs(2).Value
                    sThisColor = "#039"
                Else
                    If Not CDate(dLastDate) = CDate(rs(2).Value) Then
                        dLastDate = rs(2).Value

                        If sThisColor = "#039" Then
                            sThisColor = "#093"
                        Else
                            sThisColor = "#039"
                        End If
                    End If
                End If
	            Events(0, x) = rs(0).Value
	            Events(1, x) = "<span style='color:" & sThisColor & "'>" & Replace(rs(1).Value, "''","'") & "</span>"
	            Events(2, x) = "<span style='color:" & sThisColor & "'>" & rs(2).Value & "</span>"
                Events(3, x) = "<span style='color:" & sThisColor & "'>" & rs(3).Value & "</span>"
                Events(4, x) = "<span style='color:" & sThisColor & "'>" & sEvntType & "</span>"
                Events(5, x) = "<span style='color:" & sThisColor & "'>" & rs(4).Value & "</span>"
                Events(6, x) = rs(2).Value
                Events(7, x) = sEvntType
                Events(8, x) = "<span style='color:" & sThisColor & "'>" & GetFirstRace(rs(0).Value, sEvntType) & "</span>"
	            x = x + 1
	            ReDim Preserve Events(8, x)
            End If
	        rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next

    For x = 0 To UBound(Events, 2) - 2
        For y = x + 1 To UBound(Events, 2) - 1
            If CDate(Events(6, x)) > CDate(Events(6, y)) THen
                For z = 0 To 8
                    TempArr(z) = Events(z, x)
                    Events(z, x) = Events(z, y)
                    Events(z, y) = TempArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Function GetTech(lThisEvent, sThisType)
    GetTech = vbNullString

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT s.FirstName, s.LastName FROM Staff s INNER JOIN StaffAsgmt sa ON s.StaffID = sa.StaffID WHERE sa.EventID = " & lThisEvent 
    sql = sql & " AND sa.EventType = '" & sThisType & "' AND sa.Role = 'Tech'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        GetTech = GetTech & Replace(rs(0).Value, "''", "'") & " " & Left(rs(1).Value, 1) & ", "
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Not GetTech = vbNullString Then GetTech = Left(GetTech, Len(GetTech) - 2)
End Function

Private Function GetSupport(lThisEvent, sThisType)
    GetSupport = vbNullString

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT s.FirstName, s.LastName FROM Staff s INNER JOIN StaffAsgmt sa ON s.StaffID = sa.StaffID WHERE sa.EventID = " & lThisEvent 
    sql = sql & " AND sa.EventType = '" & sThisType & "' AND (sa.Role = 'Support' OR sa.Role = 'Other')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        GetSupport = GetSupport & Replace(rs(0).Value, "''", "'") & " " & Left(rs(1).Value, 1) & ", "
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Not GetSupport = vbNullString Then GetSupport = Left(GetSupport, Len(GetSupport) - 2)
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
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Staff Event Assignment</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
        <%If Session("role") = "admin" Then%>
            <!--#include file = "../../includes/admin_menu.asp" -->
        <%End If%>

		<div class="col-sm-10">
			<h4 class="h4">GSE&copy; Event Assignments for <%=iYear%></h4>

            <div style="background-color: #ececd8;text-align: right;margin: 0 10px 0 10px;padding: 2px;font-size: 0.85em;">
                <%If Session("role") = "staff" Then%>
                    <a href="/staff/my_history.asp?year=<%=iYear%>">My History</a>
                    &nbsp;|&nbsp;
                    <a href="/staff/calendar.asp?year=<%=iYear%>">Calendar</a>
                    &nbsp;|&nbsp;
                <%Else%>
                    <a href="staff_assign.asp?year=<%=iYear%>">Assignments By Staff</a>
                    &nbsp;|&nbsp;
                <%End If%>
                <%For i = 2013 To Year(Date) + 1%>
                    <a href="event_assign.asp?year=<%=i%>"><%=i%></a>
                    <%If Not i = Year(Date) + 1 Then%>
                        &nbsp;|&nbsp;
                    <%End If%>
                <%Next%>
            </div>

            <table class="table table-striped">
                <tr>
                    <th rowspan="2" valign="bottom">No.</th>
                    <th rowspan="2" valign="bottom">Event</th>
                    <th rowspan="2" valign="bottom">Date</th>
                    <th rowspan="2" valign="bottom">First Race</th>
                    <th rowspan="2" valign="bottom">Location</th>
                    <th rowspan="2" valign="bottom">Event Type</th>
                    <th rowspan="2" valign="bottom">Timing</th>
                    <th colspan="2" valign="bottom">Staff</th>
                </tr>
                <tr>
                    <th>Tech</th>
                    <th>Support</th>
                </tr>
                <%For i = 0 To UBound(Events, 2) - 1%>
                    <tr>
                        <%If Session("role") = "staff" Then%>
                            <%If i mod 2 = 0 Then%>
                                <td class="alt" valign="top"><%=i + 1%>)</td>
                                <td class="alt" valign="top"><%=Events(1, i)%></td>
                                <td class="alt" valign="top"><%=Events(2, i)%></td>
                                <td class="alt" valign="top"><%=Events(8, i)%></td>
                                <td class="alt" valign="top"><%=Events(3, i)%></td>
                                <td class="alt" valign="top"><%=Events(4, i)%></td>
                                <td class="alt" valign="top"><%=Events(5, i)%></td>
                                <td class="alt" valign="top"><%=GetTech(Events(0, i), Events(7, i))%></td>
                                <td class="alt" valign="top"><%=GetSupport(Events(0, i), Events(7, i))%></td>
                            <%Else%>
                                <td valign="top"><%=i + 1%>)</td>
                                <td valign="top"><%=Events(1, i)%></td>
                                <td valign="top"><%=Events(2, i)%></td>
                                <td valign="top"><%=Events(8, i)%></td>
                                <td valign="top"><%=Events(3, i)%></td>
                                <td valign="top"><%=Events(4, i)%></td>
                                <td valign="top"><%=Events(5, i)%></td>
                                <td valign="top"><%=GetTech(Events(0, i), Events(7, i))%></td>
                                <td valign="top"><%=GetSupport(Events(0, i), Events(7, i))%></td>
                            <%End If%>
                        <%ElseIf Session("role") = "admin" Then%>
                            <%If i mod 2 = 0 Then%>
                                <td class="alt" valign="top"><%=i + 1%>)</td>
                                <td class="alt" valign="top">
                                    <a href="javascript:pop('edit_asgmts.asp?event_id=<%=Events(0, i)%>&amp;event_type=<%=Events(7, i)%>',1000,700)"><%=Events(1, i)%></a>
                                </td>
                                <td class="alt" valign="top"><%=Events(2, i)%></td>
                                <td class="alt" valign="top"><%=Events(8, i)%></td>
                                <td class="alt" valign="top"><%=Events(3, i)%></td>
                                <td class="alt" valign="top"><%=Events(4, i)%></td>
                                <td class="alt" valign="top"><%=Events(5, i)%></td>
                                <td class="alt" valign="top"><%=GetTech(Events(0, i), Events(7, i))%></td>
                                <td class="alt" valign="top"><%=GetSupport(Events(0, i), Events(7, i))%></td>
                            <%Else%>
                                <td valign="top"><%=i + 1%>)</td>
                                <td valign="top">
                                    <a href="javascript:pop('edit_asgmts.asp?event_id=<%=Events(0, i)%>&amp;event_type=<%=Events(7, i)%>',1000,700)"><%=Events(1, i)%></a>
                                </td>
                                <td valign="top"><%=Events(2, i)%></td>
                                <td valign="top"><%=Events(8, i)%></td>
                                <td valign="top"><%=Events(3, i)%></td>
                                <td valign="top"><%=Events(4, i)%></td>
                                <td valign="top"><%=Events(5, i)%></td>
                                <td valign="top"><%=GetTech(Events(0, i), Events(7, i))%></td>
                                <td valign="top"><%=GetSupport(Events(0, i), Events(7, i))%></td>
                            <%End If%>
                        <%End If%>
                    </tr>
                <%Next%>
            </table>
   		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
