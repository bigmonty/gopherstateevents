<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lSeriesID
Dim sSeriesName, sSeriesRaces, sSeriesEvntsID
Dim iSeriesYear
Dim Series(), SeriesEvents(), EventRaces(), Delete()

lSeriesID = Request.QueryString("series_id")
		
Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

ReDim SeriesEvents(0)

If Request.form.Item("delete_races") = "delete_races" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesEventsID FROM SeriesEvents WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If sSeriesEvntsID = vbNullString Then
            sSeriesEvntsID = rs(0).Value & ", "
        Else
            sSeriesEvntsID = sSeriesEvntsID & rs(0).Value & ", "
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If Len(sSeriesEvntsID) > 0 Then sSeriesEvntsID = Left(sSeriesEvntsID, Len(sSeriesEvntsID) - 2)

    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesRacesID FROM SeriesRaces WHERE SeriesEventsID IN (" & sSeriesEvntsID & ")"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM SeriesRaces WHERE SeriesRacesID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
ElseIf Request.form.Item("submit_series") = "submit_series" Then
    lSeriesID = Request.Form.Item("series")
End If

If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID, SeriesName, SeriesYear FROM Series ORDER BY SeriesYear DESC, SeriesName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Series(0, i) = rs(0).Value
	Series(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve Series(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not CLng(lSeriesID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing

    i = 0
    ReDim SeriesEvents(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesEventsID, EventName, EventDate FROM SeriesEvents WHERE SeriesID = " & lSeriesID & " ORDER BY EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesEvents(0, i) = rs(0).Value
        SeriesEvents(1, i) = Replace(rs(1).Value, "''", "'")
        SeriesEvents(2, i) = rs(2).Value
	    i = i + 1
	    ReDim Preserve SeriesEvents(2, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Sub GetEventRaces(lThisEvent)
    Dim x

    x = 0
    ReDim EventRaces(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesRacesID, RaceName, Dist FROM SeriesRaces WHERE SeriesEventsID = " & lThisEvent
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        EventRaces(0, x) = rs(0).Value
        EventRaces(1, x) = Replace(rs(1).Value, "''", "'")
        EventRaces(2, x) = rs(2).Value
        x = x + 1
        ReDim Preserve EventRaces(2, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Series Event/Race Manager</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <!--#include file = "series_nav.asp" -->

			<h4 class="h4">GSE Series Events/Races</h4>

            <%If UBound(Series, 2) > 0 Then%>
   			    <form name="select_series" method="Post" action="event_mgr.asp">
                <span style="font-weight: bold;">Select Series:</span>
                <select name="series" id="series" onchange="this.form.submit1.click();">
                    <option value="">&nbsp;</option>
                    <%For i = 0 To UBound(Series, 2) - 1%>
                        <%If CLng(lSeriesID) = CLng(Series(0, i)) Then%>
                            <option value="<%=Series(0, i)%>" selected><%=Series(1, i)%></option>
                        <%Else%>
                            <option value="<%=Series(0, i)%>"><%=Series(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
			    <input type="hidden" name="submit_series" id="submit_series" value="submit_series">
			    <input type="submit" name="submit1" id="submit1" value="Select Series To View">
			    </form>
            <%End If%>

            <%If Not CLng(lSeriesID) = 0 Then%>
                <br>
                <hr>
                <br>
                (check box to delete race)
                <form name="series_races" method="post" action="event_mgr.asp?series_id=<%=lSeriesID%>">
                <%For i = 0 To UBound(SeriesEvents, 2) - 1%>
                    <%Call GetEventRaces(SeriesEvents(0, i))%>
                    <table style="margin-top: 10px;">
                        <tr><th colspan="2"><%=SeriesEvents(1, i)%> (<%=SeriesEvents(2, i)%>)</th></tr>
                        <%For j = 0 To UBound(EventRaces, 2) - 1%>
                            <tr>
                                <td style="width: 25px;"><input type="checkbox" name="delete_<%=EventRaces(0, j)%>" id="delete_<%=EventRaces(0, j)%>"></td>
                                <td style="text-align: left;"><%=EventRaces(1, j)%></td>
                            </tr>
                        <%Next%>
                    </table>
                <%Next%>
                <div>
			        <input type="hidden" name="delete_races" id="delete_races" value="delete_races">
			        <input type="submit" name="submit2" id="submit2" value="Delete Checked Races">
                </div>
                </form>
            <%End If%>
		</div>
	</div>
<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>