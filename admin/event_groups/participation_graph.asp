<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventGrp
Dim sEventGrp
Dim EventGrps

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventGrp = Request.QueryString("event_grp")
If CStr(lEventGrp) = vbNullString Then lEventGrp = "0"

Response.Buffer = False		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DISTINCT EventGrp, EventName FROM Events ORDER BY EventGrp DESC"
rs.Open sql, conn, 1, 2
EventGrps = rs.GetRows()
rs.Close
Set rs = Nothing

If CLng(lEventGrp) > 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventName FROM Events WHERE EventGrp = " & lEventGrp
    rs.Open sql, conn, 1, 2
    sEventGrp = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Participation Graphs</title>

<!--#include file = "../../includes/js.asp" -->

</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-3">
		    <h4 class="h4 bg-danger">Graph Filters</h4>

            <h5 class="h5">Event Groups</h5>
            <div class="list-group">
                <%For i = 0 To UBound(EventGrps, 2)%>
                    <%If CLng(lEventGrp) = CLng(EventGrps(0, i)) Then%>
                        <a href="participation_graph.asp?event_grp=<%=EventGrps(0, i)%>" class="list-group-item active"><%=EventGrps(1, i)%></a>
                    <%Else%>
                        <a href="participation_graph.asp?event_grp=<%=EventGrps(0, i)%>" class="list-group-item"><%=EventGrps(1, i)%></a>
                    <%End If%>
                <%Next%>
            </div>
		</div>
		<div class="col-md-7">
		    <h4 class="h4 bg-success">Yearly Participants</h4>
            <h5><%=sEventGrp%></h5>
            <div class="embed-responsive embed-responsive-16by9">
              <iframe class="embed-responsive-item" src="graph.asp"></iframe>
            </div>
		</div>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>