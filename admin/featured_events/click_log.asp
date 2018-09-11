<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lFeaturedEventID
Dim i
Dim sEventName
Dim ClickLog
Dim dEventDate

If Not Session("role") = "admin" Then 
    If Not Session("role") = "event_dir" Then Response.Redirect "http://www.google.com"
End If

lFeaturedEventID = Request.QueryString("featured_event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ClickPage, WhenClick, IPAddress FROM ClickSource WHERE FeaturedEventsID = " & lFeaturedEventID & " ORDER BY WhenClick DESC"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    ClickLog = rs.GetRows()
Else
    ReDim ClickLog(2, 0)
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM FeaturedEvents WHERE FeaturedEventsID = " & lFeaturedEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sEventName = Replace(rs(0).Value, "''", "'")
    dEventDate = rs(1).Value
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Featured Event Click Log</title>
<meta name="description" content="Gopher State Events featured events click log.">
</head>

<body>
<img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
<div class="container">
    <h3 class="h3">GSE Featured Event Click Log</h3>

    <h4 class="h4"><%=sEventName%> on <%=dEventDate%></h4>

    <ul class="nav">
        <li class="nav-item"><a class="nav-link" href="edit_event.asp?featured_event_id=<%=lFeaturedEventID%>">View Control Panel</a></li>
        <li class="nav-item"><a class="nav-link" href="click_log.asp?featured_event_id=<%=lFeaturedEventID%>">Refresh</a></li>
    </ul>

    <h5 class="h5">Num Clicks: <%=UBound(ClickLog, 2) + 1%></h5>
    <table class="table table-condensed table-striped">
        <tr>
            <th>No.</th>
            <th>Click Source</th>
            <th>Timestamp</th>
            <th>IP Address</th>
        </tr>
        <%For i = 0 To UBound(ClickLog, 2)%>
            <tr>
                <td><%=i + 1%></td>
                <td><%=ClickLog(0, i)%></td>
                <td><%=ClickLog(1, i)%></td>
                <td><%=ClickLog(2, i)%></td>
            </tr>
        <%Next%>
    </table>
 </div>
 <!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
