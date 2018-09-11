<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID
Dim Events(), TimeLine(), Tasks(22)
Dim dDateFrom, dDateTo

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

dDateFrom = Request.QueryString("date_from")
dDateTo = REquest.QueryString("date_to")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_dates") = "submit_dates" Then
    dDateFrom = Request.Form.Item("date_from")
    dDateTo = Request.Form.Item("date_to")
End If

If CStr(dDateFrom) = vbNullString Then dDateFrom = Date - 2
If Not IsDate(dDateFrom) Then Response.Redirect("http://www.google.com")

If CStr(dDateTo) = vbNullString Then dDateTo = Date + 8
If Not IsDate(dDateTo) Then Response.Redirect("http://www.google.com")

i = 0
ReDim Events(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName FROM Events WHERE EventDate >= '" & dDateFrom & "' AND EventDate <= '" & dDateTo & "' ORDER By EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.eOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sql = "SELECT ContractSent, DepositReceived, DataCollected, BibsPrepped, EventEmail, EventPromo, StaffEmail, PartData, BibLabels, PreRaceParts, "
sql = sql & "BibList, PacketPrep, ChargeTimers, ChargeClocks, ChargeBatteries, ChargeGoPro, ChargeCamera, ResolveErrors, UploadPix, PixNotif, SendInvoice, "
sql = sql & "UpdateFinances, UpdateSeries FROM EventTimeline"
Set rs = conn.Execute(sql)
For i = 0 To 22
    Tasks(i) = rs(i).Name
Next
rs.Close
Set rs = Nothing

Private Function CompletedTask(lEventID, sThisTask)
    CompletedTask = "n"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT " & sThisTask & " FROM EventTimeline WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        If Not rs(0).Value & "" = "" Then
            If CDate(rs(0).Value) Then CompletedTask ="y"
        End If
    End If
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Fitness Event TimeLine Matrix</title>
<script>
$(function() {
    $( "#date_from" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#date_to" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">Fitness Event TimeLine Matrix</h3>
			
            <form class="form-inline" name="get_dates" method="post" action="timeline_matrix.asp">
            <label for="date_from">Date From:</label>
            <input type="text" class="form-control" name="date_from" id="date_from" value="<%=dDateFrom%>">
            <label for="date_to">Date To:</label>
            <input type="text" class="form-control" name="date_to" id="date_to" value="<%=dDateTo%>">
			<input type="hidden" class="form-control" name="submit_dates" id="submit_dates" value="submit_dates">
			<input type="submit" class="form-control" name="submit1" id="submit1" value="View These">
            </form>

            <br>

            <table class="table table-condensed table-striped">
                <tr>
                    <th>Task</th>
                    <%For i = 0 To UBOund(Events, 2) - 1%>
                        <td><a href="event_timeline.asp?event_id=<%=Events(0, i)%>"><%=Events(1, i)%></a></td>
                    <%Next%>
                </tr>
                <%For j = 0 To 22%>
                    <tr>
                        <td><%=Tasks(j)%></td>
                        <%For i = 0 To UBOund(Events, 2) - 1%>
                            <%Call CompletedTask(Events(0, i), Tasks(j))%>
                            <td>
                                <%If CompletedTask(Events(0, i), Tasks(j)) = "y" Then%>
                                    <img src="/graphics/stopwatch.png" class="img-responsive" alt="Stopwatch">
                                <%End If%>
                            </td>
                        <%Next%>
                    </tr>
                <%Next%>
            </table>
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