<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim Events

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = False		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.form.Item("submit_this") = "submit_this" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventGrp, Edition FROM Events"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        rs(1).Value = Request.Form.Item("event_grp_" & rs(0).Value)
        rs(2).Value = Request.Form.Item("edition_" & rs(0).Value)
        rs.Update
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate, EventGrp, Edition FROM Events ORDER BY EventGrp, Edition"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Events By Group</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
            <a href="participation_graph.asp">Participation Graphs</a>
		    <h4 class="h4">Events By Group</h4>

			<form class="form" name="event_by_group" method="Post" action="event_groups.asp">
			<table class="table table-striped">
				<tr>
					<td style="text-align:center;" colspan="5">
						<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
						<input type="submit" name="submit1" id="submit1" value="Submit Changes">
					</td>
				</tr>
				<tr>
					<th>No.</th>
					<th>Event</th>
					<th>Date</th>
					<th>Group</th>
					<th>Edition</th>
				</tr>
				<%For i = 0 To UBound(Events, 2)%>
					<tr>
						<td style="text-align:right;"><%=i + 1%>)</td>
						<td><%=Events(1, i)%></td>
						<td><%=Events(2, i)%></td>
						<td>
 							<select name="event_grp_<%=Events(0, i)%>" id="event_grp_<%=Events(0, i)%>">
                                <option value="0">0</option>
                                <%For j = 1 To 250%>
   									<%If CInt(Events(3, i)) = CInt(j) Then%>
										<option value="<%=j%>" selected><%=j%></option>
									<%Else%>
										<option value="<%=j%>"><%=j%></option>
									<%End If%>
                                <%Next%>
							</select>
                        </td>
						<td>
 							<select name="edition_<%=Events(0, i)%>" id="edition_<%=Events(0, i)%>">
                                <option value="0">0</option>
                                <%For j = 1 To 20%>
   									<%If CInt(Events(4, i)) = CInt(j) Then%>
										<option value="<%=j%>" selected><%=j%></option>
									<%Else%>
										<option value="<%=j%>"><%=j%></option>
									<%End If%>
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
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>