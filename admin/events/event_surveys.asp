<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID
Dim sEventName, dEventDate
Dim Events(), OpenResponses(), QuickResponses()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.eOF
    If CDate(rs(2).Value) <= Date Then
	    Events(0, i) = rs(0).Value
	    Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
	    i = i + 1
	    ReDim Preserve Events(1, i)
    End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
	
If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_send") = "submit_send" Then
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

If Not CLng(lEventID) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
    sEventName = Replace(rs(0).Value, "''", "'")
    dEventDate = rs(1).Value
    rs.close
	Set rs = Nothing

    i = 0
    ReDim OpenResponses(7, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SentBy, Role, WhenSent, WordPhrase, FinalThoughts, Pricing, Expectations, Consent FROM EventSurveyResults WHERE EventID = "
    sql = sql & lEventID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        For j = 0 To 7
            If Not rs(j).Value & "" = "" Then OpenResponses(j, i) = Replace(rs(j).Value, "''", "'")
        Next
        i = i + 1
        ReDim Preserve OpenResponses(7, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Event Survey: <%=Replace(sEventName, "''", "'")%></title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td, th{
        padding-left: 5px;
    }
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">Event Survey Results</h3>
			
			<div style="margin:10px;">
				<form class="form-inline" name="which_event" method="post" action="event_surveys.asp">
                <div class="form-group">
    			    <label for="events">Select Event:</label>
				    <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
					    <option value="">&nbsp;</option>
					    <%For i = 0 to UBound(Events, 2) - 1%>
						    <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
							    <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
						    <%Else%>
							    <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
						    <%End If%>
					    <%Next%>
				    </select>
				    <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
				    <input class="form-control" type="submit" name="get_event" id="get_event" value="Get This Event">
                </div>
				</form>
			</div>
			
			<%If Not Clng(lEventID) = 0 Then%>
                <div class="bg-info">
                    <h4 class="h4">Send Survey</h4>
				    <form class="form-inline" name="send_survey" method="post" action="event_surveys.asp?event_id=<%=lEventID%>">
                    <div class="form-group">
                        <label for="send_to">Send To:</label>
                        <input type="text" class="form-control" name="send_to" id="send_to">
				        <input type="hidden" name="submit_send" id="submit_send" value="submit_send">
				        <input class="form-control" type="submit" name="send_this" id="send_this" value="Send Survey">
                    </div>
				    </form>
                </div>
                
                <div>
                    <h4 class="h4">View Survey Results</h4>

                    <h5>Open Ended Responses</h5>

                    <div class="table-responsive">
                        <table class="table">
                            <tr>
                                <th>Sent By</th>
                                <th>Role</th>
                                <th>When Sent</th>
                                <th>Word/Phrase</th>
                                <th>Thoughts</th>
                                <th><img src="/graphics/target.jpg" height="40" alt="Expectations"></th>
                                <th><img src="/graphics/pricing.png" height="40" alt="Pricing"></th>
                                <th><img src="/graphics/consent.png" height="40" alt="Consent"></th>
                            </tr>
                            <%For i = 0 To UBound(OpenResponses, 2) - 1%>
                                <tr>
                                    <td><%=OpenResponses(0, i)%></td>
                                    <td><%=OpenResponses(1, i)%></td>
                                    <td><%=OpenResponses(2, i)%></td>
                                    <td><%=OpenResponses(3, i)%></td>
                                    <td><%=OpenResponses(4, i)%></td>
                                    <td style="text-align: center;"><%=OpenResponses(5, i)%></td>
                                    <td style="text-align: center;"><%=OpenResponses(6, i)%></td>
                                    <td style="text-align: center;"><%=OpenResponses(7, i)%></td>
                                </tr>
                            <%Next%>
                        </table>
                    </div>

                    <h5>Quick Responses</h5>

                    <div class="table-responsive">
                        <table class="table-striped">
                            <tr>
                                <th>No.</th>
                                <th>Sent By</th>
                                <%For j = 0 To 14%>
                                    <th><a href="javascript:pop('survey_prompt.asp',600,400)"><%=j + 1%></a></th>
                                <%Next%>
                            </tr>
                        </table>
                    </div>
                </div>
            <%End If%>
		</div>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>