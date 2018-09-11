<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql, sql2, rs2
Dim i, j, x, m, n
Dim lEventType
Dim iNumDays, iFirstDay, iNumWks, iMonth, iYear, iNumFitness, iNumCC,  iNumNordic, iNumTotal
Dim sViewList, sPrev, sNext
Dim EventTypes(), MonthArr(1, 12), EventArray()
Dim dSelDate

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

iMonth = Request.QueryString("month")
If CStr(iMonth) = vbNullString Then iMonth = Month(Date)
If IsNumeric(iMonth) = False Then Response.Redirect("http://www.google.com")

sViewList = Request.QueryString("view_list")
If sViewList = vbNullString Then sViewList = "n"

sPrev = Request.QueryString("prev")
If sPrev = vbNullString Then sPrev = "n"

sNext = Request.QueryString("next")
If sNext = vbNullString Then sNext = "n"
	
Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
			
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
			
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If CStr(iYear) = vbNullSTring Then iYear = Year(Date)
If CStr(iMonth) = vbNullSTring Then iMonth = Month(Date)
If CStr(lEventType) = vbNullString Then lEventType= 0

If Request.Form.Item("submit_year") = "submit_year" Then
	iYear = Request.Form.Item("year")
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	iMonth = Request.Form.Item("month")
	iYear = Request.Form.Item("year")
	lEventType = Request.Form.Item("event_type")
End If

If sViewList = "n" Then
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

If sViewList = "n" Then
    i = 0
    ReDim EventArray(7, 0)
    If lEventType = "0" Then
	    If Session("role") = "admin" Then
            sql = "SELECT EventID, EventDate, EventName, EventType, ShowOnline, Location, Logo, Website FROM Events ORDER BY EventDate"
        Else
            sql = "SELECT EventID, EventDate, EventName, EventType, ShowOnline, Location, Logo, Website FROM Events WHERE ShowOnline = 'y' ORDER BY EventDate"
        End If
	    Set rs = conn.Execute(sql)
	    Do While Not rs.EOF
		    If CInt(Year(CDate(rs(1).Value))) = CInt(iYear) Then
			    If CInt(Month(CDate(rs(1).Value))) = CInt(iMonth) Then
				    EventArray(0, i) = rs(0).Value
				    EventArray(1, i) = rs(1).Value
				    EventArray(2, i) = Replace(rs(2).Value, "''", "'") 
                    If Not rs(5).Value & "" = "" Then
                        EventArray(2, i) = EventArray(2, i) & " <span style='font-weight:normal;'><br>(" & Replace(rs(5).Value, "''", "'") & ")</span>"
                    End If
				    EventArray(3, i) = GetThisType(rs(3).Value)
                    EventArray(4, i) = rs(4).Value
                    EventArray(5, i) = rs(6).Value
                    EventArray(6, i) = rs(3).Value
                    EventArray(7, i) = rs(7).Value
				    i = i + 1
				    ReDim Preserve EventArray(7, i)
			    End If
		    End If
		    rs.MoveNext
	    Loop
	    Set rs = Nothing
	
	    'now get cc meets			
	    If Session("role") = "admin" Then
            sql = "SELECT MeetsID, MeetDate, MeetName, ShowOnline, Logo, Website FROM Meets ORDER BY MeetDate"
        Else
            sql = "SELECT MeetsID, MeetDate, MeetName, ShowOnline, Logo, Website FROM Meets WHERE ShowOnline = 'y' ORDER BY MeetDate"
        End If
	    Set rs = conn2.Execute(sql)
	    Do While Not rs.EOF
		    If CInt(Year(CDate(rs(1).Value))) = CInt(iYear) Then
			    If CInt(Month(CDate(rs(1).Value))) = CInt(iMonth) Then
				    EventArray(0, i) = rs(0).Value
				    EventArray(1, i) = rs(1).Value
				    EventArray(2, i) = rs(2).Value
				    EventArray(3, i) = "1"
                    EventArray(4, i) = rs(3).Value
                    EventArray(5, i) = rs(4).Value
                    EventArray(7, i) = rs(5).Value
				    i = i + 1
				    ReDim Preserve EventArray(7, i)
			    End If
		    End If
		    rs.MoveNext
	    Loop
	    Set rs = Nothing
    ElseIf lEventType = "1" Then
	    If Session("role") = "admin" Then
            sql = "SELECT MeetsID, MeetDate, MeetName, ShowOnline, Logo, Website FROM Meets ORDER BY MeetDate"
        Else
            sql = "SELECT MeetsID, MeetDate, MeetName, ShowOnline, Logo, Website FROM Meets WHERE ShowOnline = 'y' ORDER BY MeetDate"
        End If
	    Set rs = conn2.Execute(sql)
	    Do While Not rs.EOF
		    If CInt(Year(CDate(rs(1).Value))) = CInt(iYear) Then
			    If CInt(Month(CDate(rs(1).Value))) = CInt(iMonth) Then
				    EventArray(0, i) = rs(0).Value
				    EventArray(1, i) = rs(1).Value
				    EventArray(2, i) = rs(2).Value
				    EventArray(3, i) = "cross_ctry"
                    EventArray(4, i) = rs(3).Value
                    EventArray(5, i) = rs(4).Value
                    EventArray(7, i) = rs(5).Value
				    i = i + 1
				    ReDim Preserve EventArray(7, i)
			    End If
		    End If
		    rs.MoveNext
	    Loop
	    Set rs = Nothing
    Else
	    If Session("role") = "admin" Then
            sql = "SELECT EventID, EventDate, EventName, EventType, ShowOnline, Logo, Website FROM Events WHERE EventType = '" & lEventType 
            sql = sql & "' ORDER BY EventDate"
        Else
            sql = "SELECT EventID, EventDate, EventName, EventType, ShowOnline, Logo, Website FROM Events WHERE EventType = '" & lEventType 
            sql = sql & "' AND ShowOnline = 'y'ORDER BY EventDate"
        End If
	    Set rs = conn.Execute(sql)
	    Do While Not rs.EOF
		    If CInt(Year(CDate(rs(1).Value))) = CInt(iYear) Then
			    If CInt(Month(CDate(rs(1).Value))) = CInt(iMonth) Then
				    EventArray(0, i) = rs(0).Value
				    EventArray(1, i) = rs(1).Value
				    EventArray(2, i) = rs(2).Value
				    EventArray(3, i) = rs(3).Value
                    EventArray(4, i) = rs(4).Value
                    EventArray(5, i) = rs(5).Value
                    EventArray(7, i) = rs(6).Value
				    i = i + 1
				    ReDim Preserve EventArray(7, i)
			    End If
		    End If
		    rs.MoveNext
	    Loop
	    Set rs = Nothing
    End If

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

    'determine what day to start the calendar
    iFirstDay = Weekday(dSelDate)
Else
    iNumFitness = 0
    iNumCC = 0
    iNumNordic = 0
    iNumTotal = 0

    i = 0
    ReDim EventArray(6, 0)
	If Session("role") = "admin" Then
        sql = "SELECT EventID, EventName, EventDate, EventType, ShowOnline, WebSite, Location FROM Events ORDER BY EventDate"
    Else
        sql = "SELECT EventID, EventName, EventDate, EventType, ShowOnline, WebSite, Location FROM Events WHERE ShowOnline = 'y' ORDER BY EventDate"
    End If
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
 		If CInt(Year(CDate(rs(2).Value))) = CInt(iYear) Then
            iNumFitness = CInt(iNumFitness) + 1

			EventArray(0, i) = rs(0).Value
			EventArray(1, i) = Replace(rs(1).Value, "''", "'") 
            If Not rs(6).Value & "" = "" Then
                EventArray(1, i) = EventArray(1, i) & " <span style='font-weight:normal;'>(" & Replace(rs(6).Value, "''", "'") & ")</span>"
            End If
			EventArray(2, i) = rs(2).Value
			EventArray(3, i) = rs(3).Value
            EventArray(4, i) = rs(4).Value
            EventArray(5, i) = rs(5).Value
			i = i + 1
			ReDim Preserve EventArray(6, i)
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	'now get cc meets		
	If Session("role") = "admin" Then
        sql = "SELECT MeetsID, MeetName, MeetDate, ShowOnline, Sport, Website FROM Meets ORDER BY MeetDate"
    Else
        sql = "SELECT MeetsID, MeetName, MeetDate, ShowOnline, Sport, Website FROM Meets WHERE ShowOnline = 'y' ORDER BY MeetDate"
    End If
	Set rs = conn2.Execute(sql)
	Do While Not rs.EOF
		If CInt(Year(CDate(rs(2).Value))) = CInt(iYear) Then
			EventArray(0, i) = rs(0).Value
			EventArray(1, i) = rs(1).Value
			EventArray(2, i) = rs(2).Value
			EventArray(3, i) = "1"
            EventArray(4, i) = rs(3).Value
            EventArray(5, i) = rs(4).Value
            EventArray(6, i) = rs(5).Value

            If rs(4).Value = "Nordic Ski" Then  
                iNumNordic = CInt(iNumNordic) + 1
            Else
                iNumCC = CInt(iNumCC) + 1
            End If
			i = i + 1
			ReDim Preserve EventArray(6, i)
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing

    iNumTotal = CInt(iNumNordic) + CInt(iNumCC) + CInt(iNumFitness)
End If

Private Function GetThisType(lEventType)
	sql2 = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs2 = conn.Execute(sql2)
	GetThisType = rs2(0).Value
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events (GSE) Calendar of Events</title>
<meta name="description" content="Gopher State Events calendar of events for timing road races, nordic ski, showshoe, mountain bike, fat tire bike, multi-sport, cross-country meets, and other specialty events.">

<style type="text/css">
<!--
td.calendar{
	border:medium double #ececd8;
	width:115px;
	height:75px;
	padding:5px;
	color:#039;
    background-color: none;
    text-align:center;
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
		<div class="col-sm-12">
            <%If sViewList = "y" Then%>
			    <h3 class="h3">Gopher State Events Calendar of Events: List View (<%=iNumTotal%>)</h3>

                <div class="row">
                    <div class="col-sm-9">
	                    <form class="form-inline" name="get_year" method="post" action="calendar.asp?view_list=<%=sViewList%>&amp;year=<%=iYear%>">
			            <label for="year">Year:</label>
				        <select class="form-control" name="year" id="year" onchange="this.form.submit1.click()">
					        <%For i = 2003 to Year(Date) + 1%>
						        <%If CInt(iYear) = CInt(i) Then%>
							        <option value="<%=i%>" selected><%=i%></option>
						        <%Else%>
							        <option value="<%=i%>"><%=i%></option>
						        <%End If%>
					        <%Next%>
                        </select>
				        <input type="hidden" class="form-control" name="submit_year" id="submit_year" value="submit_year">
				        <input type="submit" class="form-control" name="submit1" id="submit1" value="View Races">
                        </form>
                    </div>
                    <div class="col-sm-3">
                        <a href="calendar.asp?view_list=n&amp;year=<%=iYear%>">Switch to Calendar View</a>
                    </div>
                </div>

                <div class="row">
                    <div class="col-sm-6">
                        <h4 class="h4">Fitness Events (<%=iNumFitness%>)</h4>

                        <table class="table table-striped">
                            <tr>
                                <th class="bg-primary">No.</th>
                                <th class="bg-primary">Event</th>
                                <th class="bg-primary">Date</th>
                                <th class="bg-primary">Type</th>
                            </tr>
                            <%For i = 0 To UBound(EventArray, 2) - 1%>
                                <%If Not EventArray(3, i) = "1" Then%>
                                    <tr>
                                        <td style="text-align: right;font-size: 0.85em;"><%=i + 1%>)</td>
                                        <th style="text-align: left;">
                                            <%If EventArray(4, i) = "n" Then%>
                                                <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, i)%>"  onclick="openThis(this.href,1024,768);return false;"
                                                    style="background-color: #ccc;font-size: 0.85em;"><%=EventArray(1, i)%></a>
                                            <%Else%>
                                                <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, i)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                    style="font-size: 0.85em;"><%=EventArray(1, i)%></a>
                                            <%End If%>
                                        </th>
                                        <td style="font-size: 0.85em;"><%=EventArray(2, i)%></td>
                                        <td style="font-size: 0.85em;"><%=GetThisType(EventArray(3, i))%></td>
                                    </tr>
                                <%End If%>
                            <%Next%>
                        </table>
                    </div>
                    <div class="col-sm-6">
                        <h4 class="h4">Cross-Country Running (<%=iNumCC%>)</h4>

                        <table class="table table-striped">
                            <tr>
                                <th class="bg-primary">No.</th>
                                <th class="bg-primary">Event</th>
                                <th class="bg-primary">Date</th>
                            </tr>
                            <%j = 1%>
                            <%For i = 0 To UBound(EventArray, 2) - 1%>
                                <%If EventArray(3, i) = "1" AND EventArray(5, i) = "Cross-Country" Then%>
                                    <tr>
                                        <td style="text-align: right;font-size: 0.85em;"><%=j%>)</td>
                                    <th style="text-align: left;">
                                            <%If EventArray(4, i) = "n" Then%>
                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, i)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                    style="background-color: #ccc;font-size: 0.85em;"><%=EventArray(1, i)%></a>
                                            <%Else%>
                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, i)%>"  onclick="openThis(this.href,1024,768);return false;"
                                                    style="font-size: 0.85em;"><%=EventArray(1, i)%></a>
                                            <%End If%>
                                        </th>

                                        <td style="font-size: 0.85em;"><%=EventArray(2, i)%></td>
                                    </tr>
                                    <%j = j + 1%>
                                <%End If%>
                            <%Next%>
                        </table>

                        <h4 class="h4">Nordic Ski (<%=iNumNordic%>)</h4>

                        <table class="table table-striped">
                            <tr>
                                <th class="bg-primary">No.</th>
                                <th class="bg-primary">Event</th>
                                <th class="bg-primary">Date</th>
                            </tr>
                            <%j = 1%>
                            <%For i = 0 To UBound(EventArray, 2) - 1%>
                                <%If EventArray(3, i) = "1" AND EventArray(5, i) = "Nordic Ski" Then%>
                                    <tr>
                                        <td style="text-align: right;font-size: 0.85em;"><%=j%>)</td>
                                    <th style="text-align: left;">
                                            <%If EventArray(4, i) = "n" Then%>
                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, i)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                    style="background-color: #ccc;font-size: 0.85em;"><%=EventArray(1, i)%></a>
                                            <%Else%>
                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, i)%>"  onclick="openThis(this.href,1024,768);return false;"
                                                    style="font-size: 0.85em;"><%=EventArray(1, i)%></a>
                                            <%End If%>
                                        </th>

                                        <td style="font-size: 0.85em;"><%=EventArray(2, i)%></td>
                                    </tr>
                                    <%j = j + 1%>
                                <%End If%>
                            <%Next%>
                        </table>
                    </div>
                </div>
            <%Else%>
			    <h3 class="h3">Gopher State Events Calendar of Events: Calendar View</h3>

                <div class="row">
                    <div class="col-sm-9">
	                    <form class="form-inline"name="get_month" method="post" action="calendar.asp?view_list=n&amp;year=<%=iYear%>&amp;month=<%=iMonth%>">
			            <a href="calendar.asp?view_list=n&amp;year=<%=iYear%>&amp;month=<%=iMonth%>&amp;prev=y"><img src="/graphics/previous.png" alt="<"></a>
                        <label for="month">Month:</label>&nbsp;
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
			            <label for="event_type">Event Type:&nbsp;</label>
				        <select class="form-control" name="event_type" id="event_type" onchange="this.form.submit2.click()">
					        <option value="0">All</option>
					        <%For i = 0 to UBound(EventTypes, 2) - 1%>
						        <%If CLng(lEventType) = CLng(EventTypes(0, i)) Then%>
							        <option value="<%=EventTypes(0, i)%>" selected><%=EventTypes(1, i)%></option>
						        <%Else%>
							        <option value="<%=EventTypes(0, i)%>"><%=EventTypes(1, i)%></option>
						        <%End If%>
					        <%Next%>
				        </select>
				        <input type="hidden" class="form-control" name="submit_this" id="submit_this" value="submit_this">
				        <input type="submit" class="form-control" name="submit2" id="submit2" value="View This">
                        <a href="calendar.asp?view_list=n&amp;year=<%=iYear%>&amp;month=<%=iMonth%>&amp;next=y"><img src="/graphics/next.png" alt=">"></a>
                        </form>
                    </div>
                    <div class="col-sm-3">
                        <a href="calendar.asp?view_list=y&amp;year=<%=iYear%>">Switch to List View</a>
                    </div>
                </div>
	            </form>

	            <table class="table table-bordered">
		            <tr>
			            <th class="bg-primary">Sun</th>
			            <th class="bg-primary">Mon</th>
			            <th class="bg-primary">Tue</th>
			            <th class="bg-primary">Wed</th>
			            <th class="bg-primary">Thu</th>
			            <th class="bg-primary">Fri</th>
			            <th class="bg-primary">Sat</th>
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
								            <td>
                                                <div style="min-height:100px;">
                                                    <%=x%>
                                                    <br>
                                                    <%n = 0%>
                                                    <%For m = 0 to UBound(EventArray, 2) - 1%>
                                                        <%If CInt(Day(CDate(EventArray(1, m)))) = CInt(x) Then%>
                                                            <%If n > 0 Then%>
                                                                <hr style="margin: 5px 0 5px 0;">
                                                            <%End If%>

                                                            <%If EventArray(3, m) = "1" Then%>
                                                                <%If EventArray(4, m) = "n" Then%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                                <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;background-color: #ccc;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%End If%>
                                                                <%Else%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                                <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;background-color: #ccc;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%End If%>
                                                                <%End If%>
                                                            <%Else%>
                                                                <%If EventArray(4, m) = "n" Then%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;"
                                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <span style="font-weight: bold;background-color: #ccc;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%End If%>
                                                                <%Else%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%End If%>
                                                                <%End If%>
                                                            <%End If%>
                                                            <br>
                                                            <%n = n + 1%>
                                                        <%End If%>
                                                    <%Next%>
                                                    <%x =x + 1%>
                                                </div>
								            </td>
							            <%Else%>
								            <td><div style="min-height:100px;">&nbsp;</div></td>
							            <%End If%>
						            <%Else%>
							            <%If x <= iNumDays Then%>
								            <td>
                                                <div style="min-height:100px;">
                                                    <%=x%>
                                                    <br>
                                                    <%n = 0%>
                                                    <%For m = 0 to UBound(EventArray, 2) - 1%>
                                                        <%If CInt(Day(CDate(EventArray(1, m)))) = CInt(x) Then%>
                                                            <%If n > 0 Then%>
                                                                <hr style="margin: 5px 0 5px 0;">
                                                            <%End If%>

                                                            <%If EventArray(3, m) = "1" Then%>
                                                                <%If EventArray(4, m) = "n" Then%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                                <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="background-color: #ccc;font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;background-color: #ccc;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;background-color: #ccc;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%End If%>
                                                                <%Else%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                                <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/ccmeet_info.asp?meet_id=<%=EventArray(0, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>"   onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(0, m)%>" style="width: 50px;"></a>
                                                                                    <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    </a>
                                                                    <%End If%>
                                                                <%End If%>
                                                                <%If Date >= CDate(EventArray(1, m)) Then%>
                                                                    <br>
                                                                    <a href="/results/cc_rslts/cc_rslts.asp?sport=<%=EventArray(3, m)%>&meet_id=<%=EventArray(0, m)%>&rslts_page=overall_rslts.asp"
                                                                    style="font-size: 0.7em;color: #892700;">(View Results)</a>
                                                                <%End If%>
                                                            <%Else%>
                                                                <%If EventArray(4, m) = "n" Then%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%End If%>
                                                                <%Else%>
                                                                    <%If EventArray(5, m) & "" = "" Then%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                                <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;"
                                                                                style="font-size: 0.85em;font-weight: bold;"><%=EventArray(2, m)%></a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%Else%>
                                                                        <%If EventArray(7, m) & "" = "" Then%>
                                                                            <a href="/events/raceware_events.asp?event_id=<%=EventArray(0, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%Else%>
                                                                            <a href="<%=EventArray(7, m)%>" onclick="openThis(this.href,1024,768);return false;">
                                                                                <img src="/events/logos/<%=EventArray(5, m)%>" alt="<%=EventArray(2, m)%>" style="width: 75px;">
                                                                                <br>
                                                                                <span style="font-weight: bold;"><%=EventArray(2, m)%></span>
                                                                            </a>
                                                                        <%End If%>
                                                                    <%End If%>
                                                                <%End If%>
                                                                <%If Date >= CDate(EventArray(1, m)) Then%>
                                                                    <br>
                                                                    <a href="/results/fitness_events/results.asp?event_type=<%=EventArray(6, m)%>&event_id=<%=EventArray(0, m)%>&first_rcd=1"
                                                                    style="font-size: 0.7em;color: #892700;">(View Results)</a>
                                                                <%End If%>
                                                            <%End If%>
                                                            <br>
                                                            <%n = n + 1%>
                                                        <%End If%>
                                                    <%Next%>
                                                    <%x =x + 1%>
                                                </div>
								            </td>
							            <%Else%>
                                            <td><div style="min-height:100px;">&nbsp;</div></td>								            
							            <%End If%>
						            <%End If%>
				            <%Next%>
			            </tr>
		            <%Next%>
	            </table>
            <%End If%>
        </div>
	</div>
    <!--#include file = "../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
