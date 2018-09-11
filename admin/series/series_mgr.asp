<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, conn2, rs, sql, sql2, rs2
Dim i, j
Dim lSeriesID, lSeriesEventsID
Dim sSeriesName, sPartAwds, sPerfAwds, sSeriesNotes
Dim iSeriesYear, iMinParticip
Dim AvailEvents(), Series(), SeriesEvents(), AddEvents(), RemoveEvents(), EventRaces()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lSeriesID = Request.QueryString("series_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
							
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_events") = "submit_events" Then
    i = 0
    ReDim AddEvents(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventName, EventDate, Location FROM Events WHERE EventDate BETWEEN '1/1/" & Year(Date) & "' AND '12/31/" & Year(Date) & "'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("add_event_" & rs(0).Value) = "on" Then
            AddEvents(0, i) = rs(0).Value
            AddEvents(1, i) = Replace(rs(1).Value, "''", "'")
            AddEvents(2, i) = rs(2).Value
            If Not rs(3).Value & "" = "" Then AddEvents(3, i) = Replace(rs(3).Value, "''", "'")
            i = i + 1
            ReDim Preserve AddEvents(3, i)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(AddEvents, 2) - 1
        If AddEvents(3, i) & "" = "" Then
            sql = "INSERT INTO SeriesEvents(SeriesID, EventID, EventName, EventDate, Location) VALUES (" & lSeriesID & ", " & AddEvents(0, i) 
            sql = sql & ", '" & Replace(AddEvents(1, i), "'", "''") & "', '"  & AddEvents(2, i) & "', '"  & AddEvents(3, i) & "')"
        Else
            sql = "INSERT INTO SeriesEvents(SeriesID, EventID, EventName, EventDate, Location) VALUES (" & lSeriesID & ", " & AddEvents(0, i) 
            sql = sql & ", '" & Replace(AddEvents(1, i), "'", "''") & "', '"  & AddEvents(2, i) & "', '"  & Replace(AddEvents(3, i), "'", "''") & "')"
        End If
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        'get SeriesEventsID
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT SeriesEventsID FROM SeriesEvents ORDER BY SeriesEventsID DESC"
        rs.Open sql, conn, 1, 2
        lSeriesEventsID = rs(0).Value
        rs.Close
        Set rs = Nothing

        'now add races to series race table
        j = 0
        ReDim EventRaces(2, 0)
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RaceID, RaceName, Dist FROM RaceData WHERE EventID = " & AddEvents(0, i)
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            EventRaces(0, j) = rs(0).Value
            EventRaces(1, j) = rs(1).Value
            EventRaces(2, j) = rs(2).Value
            j = j + 1
            ReDim Preserve EventRaces(2, j)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        For j = 0 To UBound(EventRaces, 2) - 1
            sql = "INSERT INTO SeriesRaces(SeriesEventsID, RaceID, RaceName, Dist) VALUES (" & lSeriesEventsID & ", " & EventRaces(0, j) & ", '" 
            sql = sql & EventRaces(1, j) & "', '" & EventRaces(2, j) & "')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        Next
    Next
ElseIf Request.form.Item("submit_event_edits") = "submit_event_edits" Then
    i = 0
    ReDim RemoveEvents(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID FROM Events WHERE EventDate BETWEEN '1/1/" & Year(Date) & "' AND '12/31/" & Year(Date) & "'" 
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("remove_" & rs(0).Value) = "on" Then
            RemoveEvents(i) = rs(0).Value
            i = i + 1
            ReDim Preserve RemoveEvents(i)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(RemoveEvents) - 1
        sql = "DELETE FROM SeriesEvents WHERE SeriesID = " & lSeriesID & " AND EventID = " & RemoveEvents(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
ElseIf Request.form.Item("submit_settings") = "submit_settings" Then
    If Request.Form.Item("delete") = "on" Then
        sql = "DELETE FROM Series WHERE SeriesID = " & lSeriesID
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        lSeriesID = 0
    Else
        sSeriesNotes = Request.Form.Item("series_notes")
        If Not sSeriesNotes & "" = "" Then sSeriesNotes = Replace(sSeriesNotes, "'", "''")

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT SeriesName, SeriesYear, MinParticip, PartAwds, PerfAwds, SeriesNotes FROM Series WHERE SeriesID = " & lSeriesID
        rs.Open sql, conn, 1, 2
        rs(0).Value = Replace(Request.Form.Item("series_name"), "'", "''")
        rs(1).Value = Request.Form.Item("series_year")
        rs(2).Value = Request.Form.Item("min_particip")
        rs(3).Value = Request.Form.Item("part_awds")
        rs(4).Value = Request.Form.Item("perf_awds")
        rs(5).Value = Request.Form.Item("series_notes")
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.form.Item("submit_new_series") = "submit_new_series" Then
    sSeriesName = Replace(Request.Form.Item("series_name"), "'", "''")

    sql = "INSERT INTO Series(SeriesName, SeriesYear) VALUES('" & sSeriesName & "', " & Year(Date) & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesID FROM Series ORDER BY SeriesID DESC"
    rs.Open sql, conn, 1, 2
    lSeriesID = rs(0).Value
    rs.Close
    Set rs = Nothing
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

ReDim AvailEvents(7, 0)
ReDim SeriesEvents(7, 0)
If Not CLng(lSeriesID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName, SeriesYear, MinParticip, PartAwds, PerfAwds, SeriesNotes FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    iSeriesYear = rs(1).Value
    iMinParticip = rs(2).Value
    sPartAwds = rs(3).Value
    sPerfAwds = rs(4).Value
    sSeriesNotes = rs(5).Value
    rs.Close
    Set rs = Nothing

    i = 0
    j = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT e.EventID, e.EventName, e.EventDate, e.Location, e.EventFamilyID, rd.FirstName, rd.LastName, rd.Email, rd.Phone "
    sql = sql & "FROM Events e INNER JOIN EventDir rd ON e.EventDirID = rd.EventDirID WHERE e.EventDate BETWEEN '1/1/" & Year(Date) & "' AND '12/31/" 
    sql = sql & Year(Date) & "' ORDER BY EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If InSeries(rs(0).Value) = False Then
            AvailEvents(0, i) = rs(0).Value
	        AvailEvents(1, i) = Replace(rs(1).Value, "''", "'")
            AvailEvents(2, i) = rs(2).Value
            AvailEvents(3, i) = Replace(rs(3).Value, "''", "'")
            AvailEvents(4, i) = GetFamily(rs(4).Value)
	        AvailEvents(5, i) = Replace(rs(5).Value, "''", "'") & " " & Replace(rs(6).Value, "''", "'")
            AvailEvents(6, i) = rs(7).Value
            AvailEvents(7, i) = rs(8).Value
	        i = i + 1
	        ReDim Preserve AvailEvents(7, i)
        Else
            SeriesEvents(0, j) = rs(0).Value
	        SeriesEvents(1, j) = Replace(rs(1).Value, "''", "'")
            SeriesEvents(2, j) = rs(2).Value
            SeriesEvents(3, j) = Replace(rs(3).Value, "''", "'")
            SeriesEvents(4, j) = GetFamily(rs(4).Value)
	        SeriesEvents(5, j) = Replace(rs(5).Value, "''", "'") & " " & Replace(rs(6).Value, "''", "'")
            SeriesEvents(6, j) = rs(7).Value
            SeriesEvents(7, j) = rs(8).Value
	        j = j + 1
	        ReDim Preserve SeriesEvents(7, j)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function GetFamily(lFamilyID)
    If CStr(lFamilyID) = vbNullString Then lFamilyID = "0"

    If Not CLng(lFamilyID) = 0 Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT FamilyName FROM EventFamily WHERE EventFamilyID = " & lFamilyID
        rs2.Open sql2, conn, 1, 2
        GetFamily = rs2(0).Value
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

Private Function InSeries(lEventID)
    InSeries = False

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT SeriesID FROM SeriesEvents WHERE EventID = " & lEventID & " AND SeriesID = " & lSeriesID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then InSeries = True
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Date Series Manager</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <!--#include file = "series_nav.asp" -->

			<h4 class="h4">GSE Series Manager</h4>

            <br>
            
            <h5 class="h5">Create New Series</h5>
            <form class="form-inline" name="create_series" method="Post" action="series_mgr.asp">
            <label for="series_name">Series Name:</label>
            <input class="form-control" type="text" name="series_name" id="series_name" size="30">
            <input type="hidden" name="submit_new_series" id="submit_new_series" value="submit_new_series">
            <input class="form-control" type="submit" name="submit1" id="submit1" value="Create Series">
            </form>

            <%If UBound(Series, 2) > 0 Then%>
                <h5 class="h5">Manage Series</h5>
   			    <form class="form-inline" role="form" name="select_series" method="Post" action="series_mgr.asp">
                <label for="series">Select Series:</label>
                <select class="form-control" name="series" id="series" onchange="this.form.submit2.click();">
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
			    <input class="form-control" type="submit" name="submit2" id="submit2" value="Select Series To Edit">
			    </form>
            <%End If%>

            <%If Not CLng(lSeriesID) = 0 Then%>
                <h5  class="h5">Series Settings:</h5>
                <form class="form-horizontal" role="form" name="edit_settings" method="Post" action="series_mgr.asp?series_id=<%=lSeriesID%>">
                <div class="row form-group">
                    <label for="series_name" class="control-label col-sm-2">Series Name:</label>
                    <div class="col-sm-4">
                        <input class="form-control" type="text" name="series_name" id="series_name" value="<%=sSeriesName%>">
                    </div>
                    <label for="series_year" class="control-label col-sm-2">Series Year:</label>
                    <div class="col-sm-4">
                        <select class="form-control" name="series_year" id="series_year">
                            <%For i = 2013 To Year(Date)%>
                                <%If CInt(iSeriesYear) = CInt(i) Then%>
                                    <option value="<%=i%>" selected><%=i%></option>
                                <%Else%>
                                    <option value="<%=i%>"><%=i%></option>
                                <%End If%>
                            <%Next%>
                        </select>
                    </div>
                </div>
                <div class="row form-group">
                    <label for="min_particip" class="control-label col-sm-2">Minimum Particip:</label>
                    <div class="col-sm-4">
                        <select class="form-control" name="min_particip" id="min_particip">
                            <%For i = 1 To 10%>
                                <%If CInt(iMinParticip) = CInt(i) Then%>
                                    <option value="<%=i%>" selected><%=i%></option>
                                <%Else%>
                                    <option value="<%=i%>"><%=i%></option>
                                <%End If%>
                            <%Next%>
                        </select>
                    </div>
                    <label for="part_awds" class="control-label col-sm-2">Particip Awards:</label>
                    <div class="col-sm-4">
                        <select class="form-control" name="part_awds" id="part_awds">
                            <%If sPartAwds ="y" Then%>
                                <option value="y" selected>Yes</option>
                                <option value="n">No</option>
                            <%Else%>
                                <option value="y">Yes</option>
                                <option value="n" selected>No</option>
                            <%End If%>
                        </select>
                    </div>
                </div>
                <div class="row form-group">
                    <label for="perf_awds" class="control-label col-sm-2">Perf Awards:</label>
                    <div class="col-sm-4">
                        <select class="form-control" name="perf_awds" id="perf_awds">
                            <%If sPerfAwds ="y" Then%>
                                <option value="y" selected>Yes</option>
                                <option value="n">No</option>
                            <%Else%>
                                <option value="y">Yes</option>
                                <option value="n" selected>No</option>
                            <%End If%>
                        </select>
                    </div>
                    <label for="perf_awds" class="control-label col-sm-2">Series Notes:</label>
                    <div class="col-sm-4">
                        <textarea class="form-control" name="series_notes" id="series_notes" rows="3"><%=sSeriesNotes%></textarea>
                    </div>
                </div>
                <div class="row form-group">
                    <label for="perf_awds" class="control-label col-sm-2">Delete Series:</label>
                    <div class="col-sm-10">
                        <input type="checkbox" name="delete" id="delete"> 
                    </div>
                </div>
                <div class="row form-group">
                    <input type="hidden" name="submit_settings" id="submit_settings" value="submit_settings">
                    <input class="form-control" type="submit" name="submit3" id="submit3" value="Save Settings">
                </div>
                </form>

                <h5 class="h5">Manage Series Events:</h5>
			    <form class="form" name="edit_events" method="Post" action="series_mgr.asp?series_id=<%=lSeriesID%>">
			    <table class="table table-striped">
				    <tr>
					    <td style="text-align:center;" colspan="7">
						    <input type="hidden" name="submit_event_edits" id="submit_event_edits" value="submit_event_edits">
						    <input type="submit" name="submit4" id="submit4" value="Remove Selected Events">
					    </td>
				    </tr>
				    <tr>
					    <th>No.</th>
					    <th>Event</th>
					    <th>Date</th>
					    <th>Location</th>
                        <th>Family</th>
					    <th>Director</th>
					    <th>Phone</th>
                        <th>Delete</th>
				    </tr>
				    <%For i = 0 To UBound(SeriesEvents, 2) - 1%>
                        <tr>
                            <td style="text-align:right;"><%=i + 1%>)</td>
                            <td><%=SeriesEvents(1, i)%></td>
                            <td><%=SeriesEvents(2, i)%></td>
                            <td><%=SeriesEvents(3, i)%></td>
                            <td><%=SeriesEvents(4, i)%></td>
                            <td><a href="mailto:<%=SeriesEvents(6, i)%>"><%=SeriesEvents(5, i)%></a></td>
                            <td><%=SeriesEvents(7, i)%></td>
                            <td style="text-align: center;">
                                <input type="checkbox" name="remove_<%=SeriesEvents(0, i)%>" id="remove_<%=SeriesEvents(0, i)%>">
                            </td>
                        </tr>
				    <%Next%>
			    </table>
			    </form>

                <h4 class="h4">Add Series Events:</h4>
			    <form class="form" name="add_events" method="Post" action="series_mgr.asp?series_id=<%=lSeriesID%>">
			    <table class="table table-striped">
				    <tr>
					    <td colspan="7">
						    <input type="hidden" name="submit_events" id="submit_events" value="submit_events">
						    <input type="submit" name="submit5" id="submit5" value="Add Selected Events">
					    </td>
				    </tr>
				    <tr>
					    <th>No.</th>
					    <th>Event</th>
					    <th>Date</th>
					    <th>Location</th>
                        <th>Family</th>
					    <th>Director</th>
					    <th>Phone</th>
                        <th>Add</th>
				    </tr>
				    <%For i = 0 To UBound(AvailEvents, 2) - 1%>
                        <tr>
                            <td style="text-align:right;"><%=i + 1%>)</td>
                            <td><%=AvailEvents(1, i)%></td>
                            <td><%=AvailEvents(2, i)%></td>
                            <td><%=AvailEvents(3, i)%></td>
                            <td><%=AvailEvents(4, i)%></td>
                            <td><a href="mailto:<%=AvailEvents(6, i)%>"><%=AvailEvents(5, i)%></a></td>
                            <td><%=AvailEvents(7, i)%></td>
                            <td style="text-align: center;">
                                <input type="checkbox" name="add_event_<%=AvailEvents(0, i)%>" id="add_event_<%=AvailEvents(0, i)%>">
                            </td>
                        </tr>
				    <%Next%>
			    </table>
			    </form>
            <%End If%>
		</div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>