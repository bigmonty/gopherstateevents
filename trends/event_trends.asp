<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim Events

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
If lEventID = vbNullString Then lEventID = "0"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate <= '" & Date & "' ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Trends: Event Trends</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-sm-10">
            <!--#include file = "trends_menu.asp" -->

            <h3 class="h3">Event Trends (Number of Finishers)</h3>
	
			<form role="form" class="form-inline" name="get_event" method="post" action="event_trends.asp" style="margin-bottom: 10px;">
			<div  class="form-group">
				<label for="events">Event:&nbsp;</label>
				<select class="form-control" name="events" id="events" onchange="this.form.get_event.click();">
					<option value="0">&nbsp;</option>
					<%For i = 0 to UBound(Events, 2)%>
						<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
							<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%>&nbsp;(<%=Events(2, i)%>)</option>
						<%Else%>
							<option value="<%=Events(0, i)%>"><%=Events(1, i)%>&nbsp;(<%=Events(2, i)%>)</option>
						<%End If%>
					<%Next%>
				</select>
				<input class="form-control" type="hidden" name="submit_event" id="submit_event" value="submit_event">
				<input class="form-control" type="submit" name="get_event" id="get_event" value="Get This">
			</div>
			</form>				

            <div class="embed-responsive embed-responsive-16by9">
				<iframe name="event_graph" id="event_graph" frameborder="0" 
					src="event_graph.asp?event_id=<%=lEventID%>" style="width:800px;height:400px;"></iframe>
            </div>
        </div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
