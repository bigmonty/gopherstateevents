<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisType, lEventID, lCloneEvent
Dim sEventName, dEventDate, sEventSite, sClub, sWebSite, sOnlineReg, sWaiver, sMapLink, sAddress, sComments, sKeywords, sDescription, sWelcome
Dim sShowOnline
Dim EventTypes(), EventArray(11), Events()
Dim iEventGrp, iEdition
Dim dWhenShutDown

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lCloneEvent = Request.QueryString("clone_event")
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
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'") & " " & Year(rs(1).Value)
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_clone_event") = "submit_clone_event" Then
	lCloneEvent = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	If Not Request.Form.Item("event_site") & "" = "" Then sEventSite = Replace(Request.Form.Item("event_site"), "'", "''")
	If Not Request.Form.Item("club") & "" = "" Then sClub = Replace(Request.Form.Item("club"), "'", "''")
	If Not Request.Form.Item("web_site") & "" = "" Then sWebSite = Replace(Request.Form.Item("web_site"), "'", "''")
	If Not Request.Form.Item("comments") & "" = "" Then sComments = Replace(Request.Form.Item("comments"), "'", "''")
	lThisType = Request.Form.Item("event_type")
	sOnlineReg = Request.Form.Item("online_reg")
    sShowOnline = Request.Form.Item("show_online")
	If Not Request.Form.Item("waiver") & "" = "" Then sWaiver = Replace(Request.Form.Item("waiver"), "'", "''")
	If Not Request.Form.Item("map_link") & "" = "" Then sMapLink = Replace(Request.Form.Item("map_link"), "'", "''")
	If Not Request.Form.Item("meta_kywds") & "" = "" Then sKeywords = Replace(Request.Form.Item("meta_kywds"), "'", "''")
	If Not Request.Form.Item("meta_descr") & "" = "" Then sDescription = Replace(Request.Form.Item("meta_descr"), "'", "''")
	If Not Request.Form.Item("wlcme_msg") & "" = "" Then sWelcome = Replace(Request.Form.Item("wlcme_msg"), "'", "''")
	If Not Request.Form.Item("address") & "" = "" Then sAddress = Replace(Request.Form.Item("address"), "'", "''")
	dWhenShutDown = CDate(dEventDate) - 1 & " 4:00.00 PM"

    'get event grp
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventGrp FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
    iEventGrp = rs(0).Value
	rs.Close
	Set rs = Nothing
	
    'get next edition
    iEdition = 0
    Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Edition FROM Events WHERE EventGrp = " & iEventGrp & " AND EventID <> " & lEventID & " ORDER By Edition DESC"
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iEdition = CInt(rs(0).Value) + 1
	rs.Close
	Set rs = Nothing

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventType, Club, Website, Comments, ShowOnline, OnlineReg, WhenShutdown, EventGrp, Edition FROM Events WHERE EventID = " 
    sql = sql & lEventID
	rs.Open sql, conn, 1, 2
	rs(0).value = lThisType
	rs(1).value = sClub
	rs(2).value = sWebSite
	rs(3).value = sComments
	rs(4).value = sShowOnline
	rs(5).value = sOnlineReg
	rs(6).Value = dWhenShutdown
    rs(7).Value = iEventGrp
    rs(8).Value = iEdition
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT SiteName, MapLink, Address FROM SiteInfo WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	rs(0).value = sEventSite
	If sMapLink & "" = "" Then
		rs(1).Value = ""
	Else
		rs(1).Value = Replace(sMapLink, "'", "''")
	End If
	rs(2).value = sAddress
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Waiver FROM Waiver WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	If sWaiver & "" = "" Then
		rs(0).Value = rs(0).OriginalValue
	Else
		rs(0).value = Replace(sWaiver, "'", "''")
	End If
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MetaKeywords, MetaDescription, WelcomeMsg FROM EventsWeb WHERE EventsID = " & lEventID
	rs.Open sql, conn, 1, 2
	If Not sKeywords & "" = "" Then rs(0).Value  = sKeywords
	If Not sDescription & "" = "" Then rs(1).Value  = sDescription
	If Not sWelcome & "" = "" Then rs(2).Value  = sWelcome
	rs.Update
	rs.Close
	Set rs = Nothing

    Response.redirect "edit_event.asp?event_id=" & lEventID
End If

If CStr(lCloneEvent) = vbNullString Then lCloneEvent = 0

If Not CLng(lCloneEvent) = 0 Then
	sql = "SELECT s.SiteName, e.EventType, e.Club, e.Website, e.Comments, e.ShowOnline, e.OnlineReg, s.MapLink, w.Waiver, s.Address "
    sql = sql & "FROM Events e INNER JOIN SiteInfo s ON e.EventID = s.EventID INNER JOIN Waiver w ON w.EventID = e.EventID WHERE e.EventID = " 
    sql = sql & lCloneEvent
	Set rs = conn.Execute(sql)
	If Not rs(0).Value = vbNullString Then EventArray(0) = Replace(rs(0).Value, "''", "'")
	EventArray(1) = rs(1).Value
	For i = 2 to 9
		If rs(i).Value & "" = "" Then
			EventArray(i) = rs(i).Value
		Else
			EventArray(i) = Replace(rs(i).Value, "''", "'")
		End If
	Next
	Set rs = Nothing
	
	'get site web info
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MetaKeywords, MetaDescription, WelcomeMsg FROM EventsWeb WHERE EventsID = " & lEventID
	rs.Open sql, conn, 1, 2
	If Not rs(0).Value & "" = "" Then sKeywords = Replace(rs(0).Value, "''", "'")
	If Not rs(1).Value & "" = "" Then sDescription = Replace(rs(1).Value, "''", "'")
	If Not rs(2).Value & "" = "" Then sWelcome = Replace(rs(2).Value, "''", "'")
	rs.Close
	Set rs = Nothing
End If

'get event types
i = 0
ReDim EventTypes(1, 0)
sql = "SELECT EvntRaceTypesID, EvntRaceType FROM EvntRaceTypes ORDER BY EvntRaceType"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventTypes(0, i) = rs(0).Value
	EventTypes(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve EventTypes(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Function GetThisType(lEventType)
	sql = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs = conn.Execute(sql)
	GetThisType = rs(0).Value
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Clone <%=sEventName%> Event Info</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">Clone Event Information: <%=sEventName%></h4>
			
			<div style="margin:10px;">
				<form name="eventtoclone" method="post" action="clone_event.asp?event_id=<%=lEventID%>">
				<span style="font-weight:bold;">Select Event To Clone:</span>
				<select name="events" id="events" onchange="this.form.get_event.click()">
					<option value="">&nbsp;</option>
					<%For i = 0 to UBound(Events, 2) - 1%>
						<%If CLng(lCloneEvent) = CLng(Events(0, i)) Then%>
							<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
						<%Else%>
							<option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
						<%End If%>
					<%Next%>
				</select>
				<input type="hidden" name="submit_clone_event" id="submit_clone_event" value="submit_clone_event">
				<input type="submit" name="get_event" id="get_event" value="Get This Event" style="font-size:0.8em;">
				</form>
			</div>
			
			<%If Not Clng(lCloneEvent) = 0 Then%>
				<form name="clone_this" method="post" action="clone_event.asp?event_id=<%=lEventID%>&amp;clone_event=<%=lCloneEvent%>">
				<table style="margin:10px;">
					<tr>	
						<th>Event Type:</th>
						<td>
							<select name="event_type" id="event_type">
								<%For i = 0 to UBound(EventTypes, 2) - 1%>
									<%If CLng(EventArray(1)) = CLng(EventTypes(0, i)) Then%>
										<option value="<%=EventTypes(0, i)%>" selected><%=EventTypes(1, i)%></option>
									<%Else%>
										<option value="<%=EventTypes(0, i)%>"><%=EventTypes(1, i)%></option>
									<%End If%>
								<%Next%>
							</select>
                        </th>
						<th>Club:</th>
						<td><input name="club" id="club" maxlength="50" size="50" value="<%=EventArray(2)%>"></td>
					</tr>
					
					<tr>	
						<th valign="top" rowspan="3"><span style="color:#d62002">*</span>Event Site:</th>
						<td rowspan="3"><textarea name="event_site" id="event_site" rows="3" cols="30"><%=EventArray(0)%></textarea></td>
						<th valign="top">Website:</th>
						<td valign="top"><input name="web_site" id="web_site" maxlength="50" size="50" value="<%=EventArray(3)%>"></td>
					</tr>
					<tr>
						<th><span style="color:#d62002">*</span>Show Online? &nbsp;</th>
						<td>
							<%If EventArray(5) = "y" Then%>
								<input type="radio" name="show_online" id="show_online" value="y" checked>Yes &nbsp;
								<input type="radio" name="show_online" id="show_online" value="n">No
							<%Else%>
								<input type="radio" name="show_online" id="show_online" value="y">Yes &nbsp;
								<input type="radio" name="show_online" id="show_online" value="n" checked>No
							<%End If%>
						</td>
					</tr>
					<tr>
						<th><span style="color:#d62002">*</span>Online Part Reg? &nbsp;</th>
						<td>
							<%If EventArray(6) = "y" Then%>
								<input type="radio" name="online_reg" id="online_reg" value="y" checked>Yes &nbsp;
								<input type="radio" name="online_reg" id="online_reg" value="n">No
							<%Else%>
								<input type="radio" name="online_reg" id="online_reg" value="y">Yes &nbsp;
								<input type="radio" name="online_reg" id="online_reg" value="n" checked>No
							<%End If%>
						</td>
					</tr>
					<tr>	
						<th valign="top">Comments:</th>
						<td><textarea name="comments" id="comments" rows="6" cols="30"><%=EventArray(4)%></textarea></td>
					</tr>
					<tr>	
						<th valign="top">Map to Site:</th>
						<td><textarea name="map_link" id="map_link" rows="6" cols="30"><%=EventArray(7)%></textarea></td>
						<th valign="top">Event Waiver:</th>
						<td><textarea name="waiver" id="waiver" rows="6" cols="30"><%=EventArray(8)%></textarea></td>
					</tr>
					<tr>	
						<th valign="top">Welcome:</th>
						<td><textarea name="wlcme_msg" id="wlcme_msg" rows="6" cols="30"><%=sWelcome%></textarea></td>
						<th valign="top">Address:</th>
						<td><textarea name="address" id="address" rows="6" cols="30"><%=EventArray(9)%></textarea></td>
					</tr>
					<tr>	
						<th valign="top">Keywords:</th>
						<td><textarea name="meta_kywds" id="meta_kywds" rows="6" cols="30"><%=sKeywords%></textarea></td>
						<th valign="top">Description:</th>
						<td><textarea name="meta_descr" id="meta_descr" rows="6" cols="30"><%=sDescription%></textarea></td>
					</tr>
					<tr>
						<td colspan="4">
							<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
							<input type="submit" name="submit" id="submit" value="Clone Data">
						</td>
					</tr>
				</table>
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