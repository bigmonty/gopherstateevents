<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql, sql2, rs2
Dim i, j, k, x, m, n
Dim iNumDays, dSelDate, iFirstDay, iNumWks, iMonth, iYear, sPrev, sNext
Dim EventTypes(), MonthArr(1, 12), EventArray(), StaffArr()

If Session("role") & "" = "" Then Response.Redirect("http://www.google.com")

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

iMonth = Request.QueryString("month")
If CStr(iMonth) = vbNullString Then iMonth = Month(Date)
If IsNumeric(iMonth) = False Then Response.Redirect("http://www.google.com")

sPrev = Request.QueryString("prev")
If sPrev = vbNullString Then sPrev = "n"

sNext = Request.QueryString("next")
If sNext = vbNullString Then sNext = "n"

If CStr(iYear) = vbNullString Then iYear = Year(Date)
If CStr(iMonth) = vbNullString Then iMonth = Month(Date)

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
			
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
			
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then 
	iMonth = Request.Form.Item("month")
	iYear = Request.Form.Item("year")
End If

If sPrev = "y" Then
	If iMonth = 1 Then'
		iMonth = 12
		iYear = iYear - 1
	Else
		iMonth = iMonth - 1
	End If
ElseIf sNext = "y" Then
	If iMonth = 12 Then'
		iMonth = 1
		iYear = iYear + 1
	Else
		iMonth = iMonth + 1
	End If
End If

dSelDate = iMonth & "/1/" & iYear

'get event types
i = 0
ReDim EventTypes(1, 0)
sql = "SELECT EvntRaceTypesID, EvntRaceType FROM EvntRaceTypes ORDER BY EvntRaceType"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventTypes(0, i) = rs(0).Value
	EventTypes(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve EventTypes(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

MonthArr(0, 0) = "1"
MonthArr(1, 0) = "January"
MonthArr(0, 1) = "2"
MonthArr(1, 1) = "February"
MonthArr(0, 2) = "3"
MonthArr(1, 2) = "March"
MonthArr(0, 3) = "4"
MonthArr(1, 3) = "April"
MonthArr(0, 4) = "5"
MonthArr(1, 4) = "May"
MonthArr(0, 5) = "6"
MonthArr(1, 5) = "June"
MonthArr(0, 6) = "7"
MonthArr(1, 6) = "July"
MonthArr(0, 7) = "8"
MonthArr(1, 7) = "August"
MonthArr(0, 8) = "9"
MonthArr(1, 8) = "September"
MonthArr(0, 9) = "10"
MonthArr(1, 9) = "October"
MonthArr(0, 10) = "11"
MonthArr(1, 10) = "November"
MonthArr(0, 11) = "12"
MonthArr(1, 11) = "December"

'get num days in month
Select Case iMonth
	Case 2
		If iYear = "2004" Then
			iNumDays = 29
		Else
			iNumDays = 28
		End If
	Case 4
		iNumDays = 30
	Case 6
		iNumDays = 30
	Case 9
		iNumDays = 30
	Case 11
		iNumDays = 30
	Case Else
		iNumDays = 31
End Select

i = 0
ReDim EventArray(4, 0)
If Session("role") = "admin" Then
    sql = "SELECT EventID, EventDate, EventName, EventType, ShowOnline, Location FROM Events WHERE (EventDate >= '" & iMonth & "/1/" & iYear 
    sql = sql & "' AND EventDate <= '" & iMonth & "/" & iNumDays & "/" & iYear & "') ORDER BY EventDate"
Else
    sql = "SELECT EventID, EventDate, EventName, EventType, ShowOnline, Location FROM Events WHERE ShowOnline = 'y' AND (EventDate >= '" & iMonth & "/1/" 
    sql = sql & iYear & "' AND EventDate <= '" & iMonth & "/" & iNumDays & "/" & iYear & "') ORDER BY EventDate"
End If
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventArray(0, i) = rs(0).Value
	EventArray(1, i) = rs(1).Value
	EventArray(2, i) = Replace(rs(2).Value, "''", "'") 
    If Not rs(5).Value & "" = "" Then
        EventArray(2, i) = EventArray(2, i) & " <span style='font-weight:normal;'>(" & Replace(rs(5).Value, "''", "'") & ")</span>"
    End If
	EventArray(3, i) = "Fitness Event"
    EventArray(4, i) = rs(4).Value
	i = i + 1
	ReDim Preserve EventArray(4, i)
	rs.MoveNext
Loop
Set rs = Nothing
	
'now get cc/nordic meets			
If Session("role") = "admin" Then
    sql = "SELECT MeetsID, MeetDate, MeetName, Sport, ShowOnline FROM Meets WHERE (MeetDate >= '" & iMonth & "/1/" & iYear 
    sql = sql & "' AND MeetDate <= '" & iMonth & "/" & iNumDays & "/" & iYear & "') ORDER BY MeetDate"
Else
    sql = "SELECT MeetsID, MeetDate, MeetName, Sport, ShowOnline FROM Meets WHERE ShowOnline = 'y' AND (MeetDate >= '" & iMonth & "/1/" & iYear 
    sql = sql & "' AND MeetDate <= '" & iMonth & "/" & iNumDays & "/" & iYear & "') ORDER BY MeetDate"
End If
Set rs = conn2.Execute(sql)
Do While Not rs.EOF
	EventArray(0, i) = rs(0).Value
	EventArray(1, i) = rs(1).Value
	EventArray(2, i) = rs(2).Value
	EventArray(3, i) = rs(3).Value
    EventArray(4, i) = rs(4).Value
	i = i + 1
	ReDim Preserve EventArray(4, i)
	rs.MoveNext
Loop
Set rs = Nothing

'determine what day to start the calendar
iFirstDay = Weekday(dSelDate)

Private Function GetThisType(lEventType)
	sql2 = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs2 = conn.Execute(sql2)
	GetThisType = rs2(0).Value
	Set rs2 = Nothing
End Function

Private Sub GetStaff(lThisEvent, sEventType)
    Dim y

    y = 0
    ReDim StaffArr(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT StaffID FROM StaffAsgmt WHERE EventID = " & lThisEvent & " AND EventType = '" & sEventType & "'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        StaffArr(y) = GetStaffName(rs(0).Value)
        y = y + 1
        ReDim Preserve StaffArr(y)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Function GetStaffName(lThisStaff)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName FROM Staff WHERE StaffID = " & lThisStaff
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetStaffName = Replace(rs2(0).Value, "''", "'") & " " & Replace(rs2(1).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; Staff Calendar</title>
<meta name="description" content="GSE calendar of events for timing road races, nordic ski, showshoe events, mountain bike, duathlon, and cross-country meets.">

<style type="text/css">
<!--
td.calendar{
	border:none;
	width:115px;
	height:75px;
	padding:5px;
	color:#039;
    background-color: #ececd8;
	}

th.head{
	height:20px;
	font-weight:bold;
	text-align:center;
	border:thin solid #fff;
    background-color: #F3C63E;
    color: #fff;
    padding: 5px;
    white-space: nowrap;
	}
-->
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

    <div class="row">
        <%If Session("role") = "admin" Then%>
            <!--#include file = "../includes/admin_menu.asp" -->
        <%Else%>
		    <!--#include file = "staff_menu.asp" -->
        <%End If%>

		<div class="col-sm-10">
   			<h3 class="h3">GSE Staff Calendar</h3>

            <div style="background-color: #fff;border:none;text-align: right;margin: 0 10px 0 0;padding: 2px;font-size: 0.85em;">
                <%If Not Session("role") = "admin" Then%>
                    <a href="my_history.asp?year=<%=iYear%>">My History</a>
                    &nbsp;|&nbsp;
                    <a href="/admin/staff/event_assign.asp?year=<%=iYear%>">View As List</a>
                    &nbsp;|&nbsp;
                <%End If%>
            </div>

	        <form class="form-inline" role="form" name="get_month" method="post" action="calendar.asp">
			<a href="calendar.asp?view_list=n&amp;year=<%=iYear%>&amp;month=<%=iMonth%>&amp;prev=y"><img src="/graphics/previous.png" alt="<"></a>
			<label for="month">Month:</label>
			<select class="form-control" name="month" id="month" onchange="this.form.submit2.click()">
				<%For i = 0 to 11%>
					<%If CInt(iMonth) =CInt(MonthArr(0, i)) Then%>
						<option value="<%=MonthArr(0, i)%>" selected><%=MonthArr(1, i)%></option>
					<%Else%>
						<option value="<%=MonthArr(0, i)%>"><%=MonthArr(1, i)%></option>
					<%End If%>
				<%Next%>
			</select>
			&nbsp;&nbsp;<label for="year">Year:</label>&nbsp;
			<select class="form-control" name="year" id="year" onchange="this.form.submit2.click()">
				<%For i = 2003 to Year(Date) + 1%>
					<%If CInt(iYear) = CInt(i) Then%>
						<option value="<%=i%>" selected><%=i%></option>
					<%Else%>
						<option value="<%=i%>"><%=i%></option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
			<input class="form-control" type="submit" name="submit2" id="submit2" value="Get Calendar">
			<a href="calendar.asp?view_list=n&amp;year=<%=iYear%>&amp;month=<%=iMonth%>&amp;next=y"><img src="/graphics/next.png" alt=">"></a>
	        </form>

	        <table style="width:830px;background-color:#fff;margin:0;">
		        <tr>
			        <th class="head">Sun</th>
			        <th class="head">Mon</th>
			        <th class="head">Tue</th>
			        <th class="head">Wed</th>
			        <th class="head">Thu</th>
			        <th class="head">Fri</th>
			        <th class="head">Sat</th>
		        </tr>
		        <%
		        If CInt(iNumDays) = 28 And CInt(iFirstDay) = 1 Then
			        iNumWks = 4
		        ElseIf CInt(iNumDays) = 31 And CInt(iFirstDay) > 5 Then
			        iNumWks = 6
		        ElseIf CInt(iNumDays) = 30 And CInt(iFirstDay) = 7 Then
			        iNumWks = 6
		        Else
			        iNumWks = 5
		        End If
		        %>
		        <%x = 1%>
		        <%For i = 1 to CInt(iNumWks)%>
			        <tr>
				        <%For j = 1 to 7%>
						        <%If i = 1 Then%>
							        <%If j >= iFirstDay Then%>
								        <td class="calendar" valign="top">
									        <%=x%>
									        <br>
                                            <%n = 0%>
									        <%For m = 0 to UBound(EventArray, 2) - 1%>
										        <%If CInt(Day(CDate(EventArray(1, m)))) = CInt(x) Then%>
                                                    <%If n > 0 Then%>
                                                        <hr style="margin: 5px 0 5px 0;">
                                                    <%End If%>

											        <%If EventArray(3, m) = "Fitness Event" Then%>
												        <%If EventArray(4, m) = "n" Then%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
												            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>"
													            onclick="openThis(this.href,1024,768);return false;" 
                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
											            <%Else%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
 												            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" 
                                                                style="font-size: 0.85em;font-weight: bold;"
													            onclick="openThis(this.href,1024,768);return false;"><%=EventArray(2, m)%></a>
                                                        <%End If%>
											        <%Else%>
												        <%If EventArray(4, m) = "n" Then%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
                                                            <a href="javascript:pop('/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>',800,600)" 
                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
											            <%Else%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
                                                            <a href="javascript:pop('/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>',800,600)"
                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                        <%End If%>
											        <%End If%>

											        <br>

                                                    <%Call GetStaff(EventArray(0, m), EventArray(3, m))%>
                                                    
                                                    <ul style="font-size:0.85em;">
                                                        <%For k = 0 To UBound(StaffArr) - 1%>
                                                            <li><%=StaffArr(k)%></li>
                                                        <%Next%>
                                                    </ul>
                                                    <%n = n + 1%>
										        <%End If%>
									        <%Next%>
									        <%x =x + 1%>
								        </td>
							        <%Else%>
								        <td class="calendar" style="background-color:#fff;border:  1px solid #fff;" valign="top">
									        &nbsp;
								        </td>
							        <%End If%>
						        <%Else%>
							        <%If x <= iNumDays Then%>
								        <td class="calendar" valign="top">
									        <%=x%>
									        <br>
                                            <%n = 0%>
									        <%For m = 0 to UBound(EventArray, 2) - 1%>
										        <%If CInt(Day(CDate(EventArray(1, m)))) = CInt(x) Then%>
                                                    <%If n > 0 Then%>
                                                        <hr style="margin: 5px 0 5px 0;">
                                                    <%End If%>

											        <%If EventArray(3, m) = "Fitness Event" Then%>
												        <%If EventArray(4, m) = "n" Then%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
												            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>"
													            onclick="openThis(this.href,1024,768);return false;" 
                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
											            <%Else%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
 												            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" 
                                                                style="font-size: 0.85em;font-weight: bold;"
													            onclick="openThis(this.href,1024,768);return false;"><%=EventArray(2, m)%></a>
                                                        <%End If%>
                                                    <%Else%>
												        <%If EventArray(4, m) = "n" Then%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
                                                            <a href="javascript:pop('/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>',800,600)" 
                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
											            <%Else%>
                                                            <img src="/graphics/stopwatch.png" style="width: 15px;">
                                                            <a href="javascript:pop('/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>',800,600)"
                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                        <%End If%>
											        <%End If%>

											        <br>

                                                    <%Call GetStaff(EventArray(0, m), EventArray(3, m))%>
                                                    
                                                    <ul style="font-size:0.85em;">
                                                        <%For k = 0 To UBound(StaffArr) - 1%>
                                                            <li><%=StaffArr(k)%></li>
                                                        <%Next%>
                                                    </ul>
                                                    <%n = n + 1%>
										        <%End If%>
									        <%Next%>
									        <%x =x + 1%>
								        </td>
							        <%Else%>
								        <td class="calendar" style="background-color:#fff;border:  1px solid #fff;" valign="top">
									        &nbsp;
								        </td>
							        <%End If%>
						        <%End If%>
				        <%Next%>
			        </tr>
		        <%Next%>
	        </table>
        </div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
