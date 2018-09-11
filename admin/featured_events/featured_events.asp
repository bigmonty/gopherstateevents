<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sMyEvents
Dim lngTotalViews, lngTotalClicks, lngActiveViews, lngActiveClicks, lngPastViews, lngPastClicks
Dim ActiveEvents(), PastEvents()

If Not Session("role") = "admin" Then 
    If Not Session("role") = "event_dir" Then Response.Redirect "http://www.google.com"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Session("role") = "event_dir" Then
    Set rs = Server.CreateObject("ADODB.REcordset")
    sql = "SELECT EventID FROM Events WHERE EventDirID = " & Session("my_id") & " ORDER BY EventID DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sMyEvents = sMyEvents & rs(0).Value & ", "
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Not sMyEvents = vbNullString Then sMyEvents = Left(sMyEvents, Len(sMyEvents) - 2)
End If

i = 0
lngActiveViews = 0
lngActiveClicks = 0
ReDim ActiveEvents(7, i)
Set rs = Server.CreateObject("ADODB.Recordset")
If Session("role") = "event_dir" Then
    sql = "SELECT FeaturedEventsID, EventName, EventDate, Location, WhenCreated, Views, Clicks, Active FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date 
    sql = sql & "' AND '" & Date + 360 & "') AND EventID IN (" & sMyEvents & ") AND Active = 'y' ORDER BY EventDate DESC"
Else
    sql = "SELECT FeaturedEventsID, EventName, EventDate, Location, WhenCreated, Views, Clicks, Active FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date 
    sql = sql & "' AND '" & Date + 360 & "') AND Active = 'y' ORDER BY EventDate DESC"
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    ActiveEvents(0, i) = rs(0).Value
    ActiveEvents(1, i) = Replace(rs(1).Value, "''", "'")
    ActiveEvents(2, i) = rs(2).Value
    ActiveEvents(3, i) = Replace(rs(3).Value, "''", "'")
    ActiveEvents(4, i) = rs(4).Value
    ActiveEvents(5, i) = rs(5).Value
    ActiveEvents(6, i) = rs(6).Value
    ActiveEvents(7, i) = rs(7).Value

    lngActiveViews = CLng(lngActiveViews) + CLng(rs(5).Value)
    lngActiveClicks = CLng(lngActiveClicks) + CLng(rs(6).Value)

    i = i + 1
    ReDim Preserve ActiveEvents(7, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

i = 0
lngPastViews = 0
lngPastClicks = 0
ReDim PastEvents(7, i)
Set rs = Server.CreateObject("ADODB.Recordset")
If Session("role") = "event_dir" Then
    sql = "SELECT FeaturedEventsID, EventName, EventDate, Location, WhenCreated, Views, Clicks, Active FROM FeaturedEvents WHERE (EventDate < '" & Date 
    sql = sql & "' AND EventID IN (" & sMyEvents & ")) OR  Active = 'n' ORDER BY EventDate"
Else
    sql = "SELECT FeaturedEventsID, EventName, EventDate, Location, WhenCreated, Views, Clicks, Active FROM FeaturedEvents WHERE EventDate < '" & Date 
    sql = sql & "' OR Active = 'n' ORDER BY EventDate"
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    PastEvents(0, i) = rs(0).Value
    PastEvents(1, i) = Replace(rs(1).Value, "''", "'")
    PastEvents(2, i) = rs(2).Value
    PastEvents(3, i) = Replace(rs(3).Value, "''", "'")
    PastEvents(4, i) = rs(4).Value
    PastEvents(5, i) = rs(5).Value
    PastEvents(6, i) = rs(6).Value
    PastEvents(7, i) = rs(7).Value

    lngPastViews = CLng(lngPastViews) + CLng(rs(5).Value)
    lngPastClicks = CLng(lngPastClicks) + CLng(rs(6).Value)

    i = i + 1
    ReDim Preserve PastEvents(7, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

lngTotalViews = CLng(lngActiveViews) + CLng(lngPastViews)
lngTotalClicks = CLng(lngActiveClicks) + CLng(lngPastClicks)
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Featured Events Admin</title>
<meta name="description" content="Gopher State Events featured events admin utility.">
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<%If Session("role") = "event_dir" Then%>
            <!--#include file = "../../includes/event_dir_menu.asp" -->
        <%Else%>
            <!--#include file = "../../includes/admin_menu.asp" -->
        <%End If%>
		
        <div class="col-sm-10">
		    <h3 class="h3">GSE Featured Events</h3>

            <div style="text-align: right;">
                <a href="new_event.asp">Create Featured Event</a>
                |
                <a href="featured_events.asp">Refresh</a>
            </div>

            <%If Session("role") = "event_dir" Then%>
                <div class="bg-info">
                    "Featured Events" are NOT new events.  They are designating an existing event as "Featured" to give it exposure on the GSE home page,
                    results pages, the calendar page, etc.  Designating one of your events as a featured event carries a fee of $50 but, depending on how
                    long before the event happens that you list it, you could literally get your event in front of thousands of participants and get hundreds
                    of clicks to your event's website.  You can establish one of your events as a featured event <a href="/misc/featured_events.asp">here.</a>
                </div>
            <%End If%>

            <h4 class="h4">Views & Clicks Totals</h4>

            <ul class="nav">
                <li class="nav-item">Total Views: <%=lngTotalViews%></li>
                <li class="nav-item">Total Clicks: <%=lngTotalClicks%></li>
            </ul>

            <h4 class="h4">Active Events</h4>

            <ul>
                <li>Active Views: <%=lngActiveViews%></li>
                <li>Active Clicks: <%=lngActiveClicks%></li>
            </ul>

            <table class="table table-condensed table-striped">
                <tr>
                    <th>No.</th>
                    <th>Event</th>
                    <th>Date</th>
                    <th>Location</th>
                    <th>Created</th>
                    <th>Views</th>
                    <th>Clicks</th>
                    <th>Log</th>
                    <th>Active</th>
                </tr>
                <%For i = 0 To UBound(ActiveEvents, 2) - 1%>
                    <tr>
                        <td><%=i + 1%></td>
                        <td><a href="javascript:pop('edit_event.asp?featured_event_id=<%=ActiveEvents(0, i)%>',1000,700)"><%=ActiveEvents(1, i)%></a></td>
                        <td><%=ActiveEvents(2, i)%></td>
                        <td><%=ActiveEvents(3, i)%></td>
                        <td><%=ActiveEvents(4, i)%></td>
                        <td><%=ActiveEvents(5, i)%></td>
                        <td><%=ActiveEvents(6, i)%></td>
                        <td><a href="javascript:pop('click_log.asp?featured_event_id=<%=ActiveEvents(0, i)%>',1000,700)">View</a></td>
                        <td><%=ActiveEvents(7, i)%></td>
                    </tr>
                <%Next%>
            </table>

            <h4 class="h4">Past/Inactive Events</h4>
 
            <ul class="nav">
                <li class="nav-item">Past Views: <%=lngPastViews%></li>
                <li class="nav-item">Past Clicks: <%=lngPastClicks%></li>
            </ul>

           <table class="table table-condensed table-striped">
                <tr>
                    <th>No.</th>
                    <th>Event</th>
                    <th>Date</th>
                    <th>Location</th>
                    <th>Created</th>
                    <th>Views</th>
                    <th>Clicks</th>
                    <th>Log</th>
                    <th>Active</th>
                </tr>
                <%For i = 0 To UBound(PastEvents, 2) - 1%>
                    <tr>
                        <td><%=i + 1%></td>
                        <td><a href="javascript:pop('edit_event.asp?featured_event_id=<%=PastEvents(0, i)%>',1000,700)"><%=PastEvents(1, i)%></a></td>
                        <td><%=PastEvents(2, i)%></td>
                        <td><%=PastEvents(3, i)%></td>
                        <td><%=PastEvents(4, i)%></td>
                        <td><%=PastEvents(5, i)%></td>
                        <td><%=PastEvents(6, i)%></td>
                        <td><a href="javascript:pop('click_log.asp?featured_event_id=<%=PastEvents(0, i)%>',1000,700)">View</a></td>
                        <td><%=PastEvents(7, i)%></td>
                    </tr>
                <%Next%>
            </table>

            <h4>Clicks by Page</h4>
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
