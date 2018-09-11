<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, conn2, rs, sql, sql2, rs2
Dim i, j
Dim lSeriesID
Dim sSeriesName, sPartAwds, sPerfAwds, sLogo, sInfoSheet, sSeriesNotes
Dim iYear, iSeriesYear, iMinParticip
Dim Series(), SeriesEvents()

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = "0"
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"
If CLng(lSeriesID) < 0 Then Response.Redirect "http://www.google.com"
		
iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If Not IsNumeric(iYear) Then Response.Redirect "http://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
							
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.form.Item("submit_series") = "submit_series" Then
    lSeriesID = Request.Form.Item("series")
End If

If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID, SeriesName FROM Series WHERE SeriesYear = " & iYear & " ORDER BY SeriesYear DESC, SeriesName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Series(0, i) = rs(0).Value
	Series(1, i) = Replace(rs(1).Value, "''", "'") & " (" & iYear & ")"
	i = i + 1
	ReDim Preserve Series(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

ReDim SeriesEvents(6, 0)
If Not CLng(lSeriesID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName, SeriesYear, MinParticip, PartAwds, PerfAwds, Logo, InfoSheet, SeriesNotes FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    iSeriesYear = rs(1).Value
    iMinParticip = rs(2).Value
    sPartAwds = rs(3).Value
    sPerfAwds = rs(4).Value
    sLogo = rs(5).Value
    sInfoSheet = rs(6).Value
    If Not rs(7).Value & "" = "" Then sSeriesNotes = Replace(rs(7).Value, "''", "'")
    rs.Close
    Set rs = Nothing

    j = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT se.EventID, se.EventName, se.EventDate, se.Location, e.Logo, e.Website FROM SeriesEvents se INNER JOIN Events e ON se.EventID = e.EventID "
    sql = sql & "WHERE se.SeriesID = " & lSeriesID & " ORDER BY e.EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesEvents(0, j) = rs(0).Value
	    SeriesEvents(1, j) = Replace(rs(1).Value, "''", "'")
        SeriesEvents(2, j) = rs(2).Value
        SeriesEvents(3, j) = Replace(rs(3).Value, "''", "'")
        SeriesEvents(4, j) = GetRaceDist(rs(0).Value)
        SeriesEvents(5, j) = rs(4).Value
        SeriesEvents(6, j) = rs(5).Value
	    j = j + 1
	    ReDim Preserve SeriesEvents(6, j)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function GetRaceDist(lEventID)
    GetRaceDist = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Dist FROM SeriesRaces sr INNER JOIN SeriesEvents se ON sr.SeriesEventsID = se.SeriesEventsID WHERE se.EventID = " & lEventID
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        GetRaceDist = GetRaceDist & Replace(rs2(0).Value, "_", " ") & ", "
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If Not GetRaceDist = vbNullString Then
        GetRaceDist = Trim(GetRaceDist)
        GetRaceDist = Left(GetRaceDist, Len(GetRaceDist) - 1)
    End If
End Function
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE (Gopher State Events) Series Manager</title>
<meta name="description" content="GSE race series for road races, nordic ski, showshoe events, mountain bike, duathlon, and cross-country meet management (timing).">
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div class="row">
		<div class="col-md-9">
			<h1 class="h1">Gopher State Events Race Series</h1>

            <div>
                <ul class="nav">
                    <li class="nav-item"><a class="nav-link" href="/series/series_results.asp?series_id=<%=lSeriesID%>" onclick="openThis(this.href,1024,768);return false;" 
                        style="font-weight: bold;color: #f00;">Series Standings</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</li>
                    <li class="nav-item"><span  class="nav-link">Select Year:</span></li>
                    <%For i = 2014 To Year(Date) + 1%>
                        <li class="nav-item"><a class="nav-link" href="series_info.asp?year=<%=i%>"><%=i%></a>&nbsp;&nbsp;</li>
                    <%Next%>
                    <li><a class="nav-link" href="javascript:pop('how_it_works.asp',600,650)">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;How It Works</a></li>
                </ul>
            </div>

   			<form role="form" class="form-inline" name="select_series" method="Post" action="series_info.asp?year=<%=iYear%>&amp;series_id=<%=lSeriesID%>">
			<div class="form-group">
                <label for="series">Select Series:</label>
                <select class="form-control" name="series" id="series" onchange="this.form.submit1.click();">
                    <option value="">&nbsp;</option>
                    <%For i = 0 To UBound(Series, 2) - 1%>
                        <%If CLng(lSeriesID) = CLng(Series(0, i)) Then%>
                            <option value="<%=Series(0, i)%>" selected><%=Series(1, i)%></option>
                        <%Else%>
                            <option value="<%=Series(0, i)%>"><%=Series(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </div>
            <div class="form-group">
			    <input class="form-control" type="hidden" name="submit_series" id="submit_series" value="submit_series">
			    <input class="form-control" type="submit" name="submit1" id="submit1" value="Select This">
            </div>
			</form>

            <%If Not CLng(lSeriesID) = 0 Then%>
                <br>
                <h4 class="h4"><%=sSeriesName%></h4>

                <div class="row">
                    <div class="col-sm-6 bg-danger">
                        <h5 class="h5">Series Description:</h5>
                        <%If Not sSeriesNotes = vbNullString Then%>
                            <%=sSeriesNotes%>
                        <%End If%>
                    </div>
                    <div class="col-sm-6 bg-success">
                        <h5 class="h5">Series Award Structure:</h5>
                    
                        <p>
                            This series <span style="font-weight: bold;">DOES
                            <%If sPerfAwds = "n" Then%>
                                NOT
                            <%End If%>
                            </span>have performance-based awards.
                            This series <span style="font-weight: bold;">DOES
                            <%If sPartAwds = "n" Then%>
                                NOT
                            <%End If%>
                            </span>have participation-based awards.
                        </p>
                    </div>
                </div>

                <h4 class="h4">Series Events:</h4>
			    <table class="table table-striped">
				    <tr>
					    <th style="text-align:right;">No.</th>
					    <th>Event</th>
					    <th>Date</th>
					    <th>Location</th>
                        <th>Distance(s)</th>
				    </tr>
				    <%For i = 0 To UBound(SeriesEvents, 2) - 1%>
						<tr>
							<td style="text-align:right;"><%=i + 1%>)</td>
							<td>
                                <%If Not SeriesEvents(6, i) & "" = "" Then%>
                                    <a href="<%=SeriesEvents(6, i)%>" onclick="openThis(this.href,1024,768);return false;"><%=SeriesEvents(1, i)%></a>
                                <%Else%>
                                    <a href="javascript:pop('/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>',800,600)" 
                                    rel="nofollow"><%=SeriesEvents(1, i)%></a>
                                <%End If%>
                            </td>
							<td><%=SeriesEvents(2, i)%></td>
							<td><%=SeriesEvents(3, i)%></td>
                            <td><%=SeriesEvents(4, i)%></td>
                        </tr>
				    <%Next%>
			    </table>
            <%End If%>
		</div>
        <%If CLng(lSeriesID) = 0 Then%>		
            <!--#include file = "../includes/vira_sponsors.asp" -->
        <%Else%>
            <div class="col-md-3">
                <%If sInfoSheet & "" = "" Then%>
                    <%If Not sLogo & "" = "" Then%>
                        <img class ="img-responsive" src="/admin/series/logos/<%=sLogo%>" alt="Info Sheet">
                    <%End If%>
                <%Else%>
                    <%If sLogo & "" = "" Then%>
                        <a href="/admin/series/info_sheets/<%=sInfoSheet%>" target="_blank" style="margin: 0 0 10px 50px;">Info Sheet</a>
                    <%Else%>
                        <a href="/admin/series/info_sheets/<%=sInfoSheet%>" target="_blank">
                            <img class ="img-responsive" src="/admin/series/logos/<%=sLogo%>" alt="Info Sheet">
                        </a>
                    <%End If%>
                <%End If%>
                <ul style="margin:0;padding:0;list-style:none;">
				    <%For i = 0 To UBound(SeriesEvents, 2) - 1%>
                        <%If SeriesEvents(6, i) & "" = "" Then%>
                            <%If SeriesEvents(5, i) & "" = "" Then%>
                                <li>
					                <a href="javascript:pop('/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>',800,700)" rel="nofollow">
                                        <%=SeriesEvents(1, i)%>
                                    </a>
				                </li>
                            <%Else%>
                                <li>
					                <a href="javascript:pop('/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>',800,700)">
                                        <img class ="img-responsive" src="/events/logos/<%=SeriesEvents(5, i)%>" alt="<%=SeriesEvents(1, i)%>">
                                    </a>
				                </li>
                            <%End If%>
                        <%Else%>
                            <%If SeriesEvents(5, i) & "" = "" Then%>
                                <li>
					                <a href="<%=SeriesEvents(6, i)%>" rel="nofollow" onclick="openThis(this.href,1024,768);return false;">
                                        <%=SeriesEvents(1, i)%>
                                    </a>
				                </li>
                            <%Else%>
                                <li>
					                <a href="<%=SeriesEvents(6, i)%>" onclick="openThis(this.href,1024,768);return false;">
                                        <img class ="img-responsive" src="/events/logos/<%=SeriesEvents(5, i)%>" alt="<%=SeriesEvents(1, i)%>">
                                    </a>
				                </li>
                            <%End If%>
                        <%End If%>
                        <hr>
                    <%Next%>
			    </ul>
   		    </div>
        <%End If%>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>