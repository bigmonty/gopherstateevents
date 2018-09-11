<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, rs, conn, rs2, sql2
Dim i
Dim lEventID
Dim Events(), SentTo()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("get_event") = "get_event" Then
    lEventID = Request.Form.Item("events")
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

i = 0
ReDim Events(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate, Location FROM Events WHERE EventDate BETWEEN '1/1/2013' AND '" & Date & "' ORDER BY EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & " - " & rs(3).Value & ")"
    Events(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve Events(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not CLng(lEVentID) = 0 Then
    i = 0
    ReDim SentTo(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TargetEvent, WhenSent, NumSent FROM PromoEmail WHERE EventRecips = " & lEventID & " ORDER BY WhenSent"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SentTo(0, i) = GetEventName(rs(0).Value)
        SentTo(1, i) = rs(1).Value
        SentTo(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve SentTo(2, i)
        rs.MoveNext
    Loop
    rs.Close
	Set rs = Nothing
End If

Private Function GetEventName(lThisEvent)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT EventName FROM Events WHERE EventID = " & lThisEvent
    rs2.Open sql2, conn, 1, 2
    GetEventName = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE&copy; Promotional Email Saturatin Screen</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
	th{
		white-space:nowrap;
        text-align: right;
	}
</style>

</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE Event Promotion Saturation Screen</h4>
		
            <table style="margin: 10px 0 0 0;padding: 0;width: 100%;">
                <tr>
                    <td>
                        <form name="select_event" method="post" action="promo_saturation.asp">
                        <span style="font-weight: bold;">Promote:</span>
                        <select name="events" id="events" onchange="this.form.submit2.click();">
                            <option value="0">Non-System Event</option>
                            <%For i = 0 To UBound(Events, 2) - 1%>
                                <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
                                    <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
                                <%Else%>
                                    <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                                <%End If%>
                            <%Next%>
                        </select>
                        <input type="hidden" name="get_event" id="get_event" value="get_event">
                        <input type="submit" name="submit2" id="submit2" value="Get This">
                        </form>
                    </td>
                    <td style="text-align: right;"><a href="email_event_promo.asp">Event Promo Page</a></td>
                </tr>
            </table>

            <%If Not CLng(lEventID) = 0 Then%>
                <hr style="margin: 10px 0 10px 0;">
                
                <h3>Emails Sent To These Participants</h3>
                <br>
                <ol>
                    <%For i = 0 To UBound(SentTo, 2) - 1%>
                        <li><%=SentTo(0, i)%>&nbsp;(<%=SentTo(1, i)%>) Num Sent=<%=SentTo(2, i)%></li>
                    <%Next%>
                </ol>
            <%End If%>
  		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
