<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim TodayEvents()

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim TodayEvents(7, 0)
sql = "SELECT e.EventID, e.EventName, e.Logo, e.EventType, e.Website, si.MapLink "
sql = sql & "FROM Events e INNER JOIN SiteInfo si ON e.EventID = si.EventID WHERE e.EventDate = '" & Date & "'"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	TodayEvents(0, i) = rs(0).Value
	TodayEvents(1, i) = Replace(rs(1).Value, "''", "'") 
    TodayEvents(2, i) = rs(2).Value
    TodayEvents(3, i) = "fitness"
    TodayEvents(4, i) = rs(3).Value
 	TodayEvents(5, i) = rs(4).Value   
	TodayEvents(7, i) = rs(5).Value
	i = i + 1
	ReDim Preserve TodayEvents(7, i)
	rs.MoveNext
Loop
Set rs = Nothing
	
'now get cc/nordic	
sql = "SELECT m.MeetsID, m.MeetName, m.Sport, m.Logo, m.Website, m.StartList, ml.MaplInk "
sql = sql & "FROM Meets m INNER JOIN MapLinks ml ON m.MeetsID = ml.MeetsID "
sql = sql & "WHERE m.ShowOnline = 'y' AND m.MeetDate = '" & Date & "'"
Set rs = conn2.Execute(sql)
Do While Not rs.EOF
	TodayEvents(0, i) = rs(0).Value
	TodayEvents(1, i) = rs(1).Value
    TodayEvents(2, i) = rs(3).Value
    TodayEvents(3, i) = rs(2).Value
    TodayEvents(5, i) = rs(4).Value
    TodayEvents(6, i) = rs(5).Value
	TodayEvents(7, i) = rs(6).Value
	i = i + 1
	ReDim Preserve TodayEvents(7, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Function GetEventType(lThisEventType)
	sql2 = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lThisEventType
	Set rs2 = conn.Execute(sql2)
	GetEventType = rs2(0).Value
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; Featured Events</title>
<meta name="description" content="Today's GSE Events">

<style type="text/css">
	td, th{text-align:center;}
</style>
</head>

<body onload="javascript:request_feature.your_name.focus()">
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div class="row">
		<div class="col-sm-12">
			<h3 class="h3">Today's GSE Events</h3>

			<%If UBound(TodayEvents, 2) > 0 Then%>
				<table class="table">
					<tr>
						<th colspan="2">Event Name</th>
						<th>Results</th>
						<th>Type</th>
						<th>Info Link</th>
						<th>Website</th>
						<th>Location</th>
						<th>Start List</th>
					</tr>
					<%For i = 0 To UBound(TodayEvents, 2) - 1%>
						<tr>
							<td>
								<%If TodayEvents(3, i) = "fitness" Then%>
									<%If TodayEvents(2, i) & "" = "" Then%>
										<img src="/graphics/info.jpg" alt="Info" style="height:50px;">
									<%Else%>
										<img src="/events/logos/<%=TodayEvents(2, i)%>" alt="<%=TodayEvents(1, i)%>" style="height:50px;">
									<%End If%>
								<%Else%>
									<%If TodayEvents(2, i) & "" = "" Then%>
										<img src="/graphics/info.jpg" alt="Info" style="height:50px;"> 
									<%Else%>
										<img src="/events/logos/<%=TodayEvents(2, i)%>" alt="<%=TodayEvents(1, i)%>" style="height:50px;">
									<%End If%>
								<%End If%>
							</td>
							<td>
								<%If TodayEvents(3, i) = "fitness" Then%>
									<a href="/results/fitness_events/results.asp?event_type=<%=TodayEvents(4, i)%>&amp;event_id=<%=TodayEvents(0, i)%>&first_rcd=1">
										<img src="/graphics/race_results.jpg" alt="Race Results" style="height: 30px;">
									</a>
								<%Else%>
									<a href="/results/cc_rslts/cc_rslts.asp?meet_id=<%=TodayEvents(0, i)%>&amp;sport=<%=TodayEvents(3, i)%>&amp;rslts_page=overall_rslts.asp">
										<img src="/graphics/race_results.jpg" alt="Race Results" style="height: 30px;">
									</a>
								<%End If%>
							</td>
							<th style="text-align:left;"><%=TodayEvents(1, i)%></th>
							<td>
								<%If TodayEvents(3, i) = "fitness" Then%>
									<%=TodayEvents(4, i)%>
								<%Else%>
									<%=TodayEvents(3, i)%>
								<%End If%>
							</td>
							<td>
								<%If TodayEvents(3, i) = "fitness" Then%>
									<a href="/events/raceware_events.asp?event_id=<%=TodayEvents(0, i)%>"  
										onclick="openThis(this.href,1024,768);return false;"><img src="/graphics/info.jpg" alt="Info" style="height:40px;">
									</a>
								<%Else%>
									<a href="/events/ccmeet_info.asp?meet_id=<%=TodayEvents(0, i)%>"  
										onclick="openThis(this.href,1024,768);return false;"><img src="/graphics/info.jpg" alt="Info" style="height:40px;">
									</a>
								<%End If%>
							</td>
							<td>
								<%If TodayEvents(5, i) & "" = "" Then%>
									n/a
								<%Else%>
									<a href="<%=TodayEvents(5, i)%>" onclick="openThis(this.href,1024,768);return false;">
										<img src="/graphics/social_media/web_logo.jpg" alt="Info" style="height:60px;">
									</a>
								<%End If%>
							</td>
							<td>
								<%If TodayEvents(7, i) & "" = "" Then%>
									n/a
								<%Else%>
									<a href="<%=TodayEvents(7, i)%>" onclick="openThis(this.href,1024,768);return false;">
										<img src="/graphics/social_media/map_quest.jpg" alt="Info" style="height:50px;">
									</a>
								<%End If%>
							</td>
							<td>
								<%If TodayEvents(3, i) = "Nordic Ski" Then%>
									<a href="/ccmeet_admin/manage_meet/run_order/<%=TodayEvents(6, i)%>" target="_blank">
										<img src="http://www.gopherstateevents.com/graphics/social_media/list.png" alt="View" style="height: 30px;">
									</a>
								<%Else%>
									n/a
								<%End If%>
							</td>
						</tr>
					<%Next%>
				</table>
			<%Else%>
				<p>We have no events scheduled today.</p>
			<%End If%>
		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
