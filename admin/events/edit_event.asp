<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisType, lEventID, lEventDirID, lEventFamilyID
Dim sEventName, dEventDate, sEventSite, sClub, sWebSite, sWeather, sComments, sOnlineReg, sShowOnline, sWaiver, sRsltsOfficial, sWelcome, sStaffNotes
Dim sNeedBibs, sNeedPins, sMapLink, sAddress, sEventClass, sAdminNotes, sEventDirEmail, sFoundBy, sOptOut, sRsltsSort, sStartTime, sPacketPickup
Dim sStartBox, sAnnouncer, sDigitalDisplay, sLocation, sLocalPower, sTearOffs, sRaceReport, sGallery, sDynamicBibAssign, sNeedTruss, sPixSponsor
Dim sSixDigitsOnly
Dim iEdition, iEventGrp, iMaxStart, iAntFieldSize
Dim sngInvoice, sngDuplRange, sngMinTime
Dim EventDir(), EventTypes(), EventArray(37), EventGrps(), EventFams()
Dim dWhenShutdown
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim EventFams(1, 0)
sql = "SELECT EventFamilyID, FamilyName FROM EventFamily ORDER BY FamilyName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventFams(0, i) = rs(0).Value
	EventFams(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve EventFams(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim EventDir(1, 0)
sql = "SELECT EventDirID, FirstName, LastName FROM EventDir WHERE Active = 'y' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventDir(0, i) = rs(0).Value
	EventDir(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve EventDir(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

Dim Events
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing
	
If Request.Form.Item("submit_official") = "submit_official" Then
	sql = "DELETE FROM OfficialRslts WHERE EventID = " & lEventID 
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	If Request.Form.Item("official_rslts") = "y" Then
		sql = "INSERT INTO OfficialRslts (EventID) VALUES (" & lEventID & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	End If
ELseIf Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	If Request.Form.Item("delete") = "on" Then
		sql = "DELETE FROM Events WHERE EventID = " & lEventID
		Set rs = conn.Execute(sql)
		Set rs = Nothing
		
		lEventID = 0
	Else
		If Not Request.Form.Item("event_name") & "" = "" Then sEventName = Replace(Request.Form.Item("event_name"), "'", "''")
		dEventDate = Request.Form.Item("date_month") & "/" & Request.Form.Item("date_day") & "/" & Request.Form.Item("date_year")
		If Not Request.Form.Item("event_site") & "" = "" Then sEventSite = Replace(Request.Form.Item("event_site"), "'", "''")
		If Not Request.Form.Item("club") & "" = "" Then sClub = Replace(Request.Form.Item("club"), "'", "''")
		If Not Request.Form.Item("web_site") & "" = "" Then sWebSite = Replace(Request.Form.Item("web_site"), "'", "''")
		If Not Request.Form.Item("weather") & "" = "" Then sWeather = Replace(Request.Form.Item("weather"), "'", "''")
		If Not Request.Form.Item("comments") & "" = "" Then sComments = Replace(Request.Form.Item("comments"), "'", "''")
		lThisType = Request.Form.Item("event_type")
		sShowOnline = Request.Form.Item("show_online")
		sOnlineReg = Request.Form.Item("online_reg")
		If Not Request.Form.Item("waiver") & "" = "" Then sWaiver = Replace(Request.Form.Item("waiver"), "'", "''")
		If Not Request.Form.Item("map_link") & "" = "" Then sMapLink = Replace(Request.Form.Item("map_link"), "'", "''")
		If Not Request.Form.Item("wlcme_msg") & "" = "" Then sWelcome = Replace(Request.Form.Item("wlcme_msg"), "'", "''")
		If Not Request.Form.Item("address") & "" = "" Then sAddress = Replace(Request.Form.Item("address"), "'", "''")
		dWhenShutDown = Request.Form.Item("shutdown_month") & "/" & Request.Form.Item("shutdown_day") & "/" & Request.Form.Item("shutdown_year")
		dWhenShutDown = dWhenShutDown & " " & Request.Form.Item("shutdown_hour") & ":00.00 " & Request.Form.Item("shutdown_ampm")
		iEventGrp = Request.Form.Item("event_grp")
		iEdition = Request.Form.Item("edition")
		sngInvoice = Request.Form.Item("invoice")
		sNeedBibs = Request.Form.Item("need_bibs")
		sNeedPins = Request.Form.Item("need_pins")
        lEventDirID = Request.Form.Item("event_dir")
        lEventFamilyID = Request.Form.Item("event_family")
        sEventClass = Request.Form.Item("event_class")
        sAdminNotes = Request.Form.Item("admin_notes")
        If Not sAdminNotes = vbNullString Then sAdminNotes = Replace(sAdminNotes, "'", "''")
        sOptOut = Request.Form.Item("opt_out")
        sFoundBy = Request.Form.Item("found_by")
        If Not sFoundBy = vbNullString Then sFoundBy = Replace(sFoundBy, "'", "''")
        If Not Request.Form.Item("packet_pickup") & "" = "" Then sPacketPickup = Replace(Request.Form.Item("packet_pickup"), "'", "''")
        sStartBox = Request.Form.Item("start_box")
        sAnnouncer = Request.Form.Item("announcer")
        sDigitalDisplay = Request.Form.Item("digital_display")
        sLocation = Request.Form.Item("location")
        sLocalPower = Request.Form.Item("local_power")
        sTearOffs = Request.Form.Item("tear_offs")
        iAntFieldSize = Request.Form.Item("ant_field_size")
        sDynamicBibAssign = Request.Form.Item("dynamic_bib_assign")
        sNeedTruss = Request.Form.Item("need_truss")
        sPixSponsor = Request.Form.Item("pix_sponsor")
        sStaffNotes = Request.Form.Item("staff_notes")
        sSixDigitsOnly = Request.Form.Item("six_digits_only")

        'rfid settings
        sRsltsSort = Request.Form.Item("rslts_sort")
        sStartTime = Request.Form.Item("start_time")
        sngDuplRange = Request.Form.Item("dupl_range")
        sngMinTime = Request.Form.Item("min_time")
        iMaxStart = Request.Form.Item("max_start")

        'race report settings
	    If Not Request.Form.Item("weather") & "" = "" Then sWeather = Replace(Request.Form.Item("weather"), "'", "''")
        If Not Request.Form.Item("race_report") & "" = "" Then sRaceReport = Replace(Request.Form.Item("race_report"), "'", "''")
        If Not Request.Form.Item("gallery") & "" = "" Then sGallery = Replace(Request.Form.Item("gallery"), "'", "''")

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT EventName, EventDate, EventType, Club, Website, Weather, Comments, ShowOnline, OnlineReg, WhenShutdown, NeedBibs, "
		sql = sql & "NeedPins, EventGrp, Edition, Invoice, EventDirID, EventFamilyID, EventClass, AdminNotes, FoundBy, OptOut, "
        sql = sql & "PacketPickup, StartBox, Announcer, DigitalDisplay, Location, LocalPower, TearOffs, AntFieldSize, DynamicBibAssign, NeedTruss, "
        sql = sql & "PixSponsor, StaffNotes, SixDigitsOnly FROM Events WHERE EventID = " & lEventID
		rs.Open sql, conn, 1, 2
		If sEventName & "" = "" Then
			rs(0).Value = Replace(rs(0).OriginalValue, "''", "'")
		Else
			rs(0).value = sEventName
		End If
		rs(1).value = dEventDate
		rs(2).value = lThisType
		rs(3).value = sClub
		rs(4).value = sWebSite
		rs(5).value = sWeather
		rs(6).value = sComments
		rs(7).value = sShowOnline
		rs(8).value = sOnlineReg
		If IsDate(dWhenShutdown) Then
			rs(9).Value = dWhenShutdown
		Else
			rs(9).Value = rs(9).OriginalValue
		End If
		rs(10).value = sNeedBibs
		rs(11).value = sNeedPins
		rs(12).value = iEventGrp
		rs(13).value = iEdition
        rs(14).value = sngInvoice
        rs(15).value = lEventDirID
        rs(16).value = lEventFamilyID
        rs(17).value = sEventClass
        rs(18).value = sAdminNotes
        rs(19).value = sFoundBy
        rs(20).value = sOptOut
        rs(21).value = sPacketPickup
        rs(22).value = sStartBox
        rs(23).value = sAnnouncer
        rs(24).value = sDigitalDisplay
        rs(25).value = sLocation
        rs(26).value = sLocalPower
        rs(27).Value = sTearOffs
        rs(28).Value = iAntFieldSize
        rs(29).Value = sDynamicBibAssign
        rs(30).Value = sNeedTruss
        rs(31).Value = sPixSponsor
        rs(32).Value = sStaffNotes
        rs(33).Value = sSixDigitsOnly
		rs.Update
		rs.Close
		Set rs = Nothing
		
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT SiteName, MapLink, Address FROM SiteInfo WHERE EventID = " & lEventID
		rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
		    rs(0).value = sEventSite
		    If sMapLink & "" = "" Then
			    rs(1).Value = ""
		    Else
			    rs(1).Value = Replace(sMapLink, "'", "''")
		    End If
		    rs(2).value = sAddress
		    rs.Update
        End If
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
	
	    'get site web info
	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT WelcomeMsg FROM EventsWeb WHERE EventsID = " & lEventID
	    rs.Open sql, conn, 1, 2
	    If rs.RecordCount > 0 Then
		    If Not sWelcome & "" = "" Then 
                rs(0).Value = Replace(sWelcome, "''", "'")
                rs.Update
            End If
	    End If
	    rs.Close
	    Set rs = Nothing
	
	    'get rfid settings
	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT StartTime, DuplRange, MinTime, MaxStart, RsltsSort FROM RFIDSettings WHERE EventID = " & lEventID
	    rs.Open sql, conn, 1, 2
        rs(0).Value = sStartTime
        rs(1).Value = sngDuplRange
        rs(2).Value = sngMinTime
        rs(3).Value = iMaxStart
        rs(4).Value = sRsltsSort
        rs.Update
	    rs.Close
	    Set rs = Nothing

	    Set rs=Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT RaceReport, Weather, Gallery FROM RaceReport WHERE EventID = " & lEventID
	    rs.Open sql, conn, 1, 2
	    If rs.RecordCount > 0 Then
            rs(0).value = sRaceReport
            rs(1).value = sWeather
            rs(2).value = sGallery
	        rs.Update
        End If
	    rs.Close
	    Set rs = Nothing
	End If
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

If Not CLng(lEventID) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT e.EventName, e.EventDate, s.SiteName, e.EventType, e.Club, e.Website, e.Weather, e.Comments, e.ShowOnline, e.OnlineReg, "
    sql = sql & "s.MapLink, w.Waiver, s.Address, e.WhenShutdown, e.NeedBibs, e.NeedPins, e.EventGrp, e.Edition, e.Invoice, e.EventDirID, "
    sql = sql & "e.EventFamilyID, e.EventClass, e.AdminNotes, e.FoundBy, e.OptOut, e.PacketPickup, e.StartBox, e.Announcer, e.DigitalDisplay, e.Location, "
    sql = sql & "e.LocalPower, e.TearOffs, e.AntFieldSize, e.DynamicBibAssign, e.NeedTruss, e.PixSponsor, StaffNotes, e.SixDigitsOnly "
    sql = sql & "FROM Events e INNER JOIN SiteInfo s ON e.EventID = s.EventID INNER JOIN Waiver w ON w.EventID = e.EventID WHERE e.EventID = " & lEventID
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
	    For i = 0 to 37
		    If Not rs(i).Value & "" = "" Then  EventArray(i) = Replace(rs(i).Value, "''", "'")
	    Next
    End If
    rs.close
	Set rs = Nothing
	
	If EventArray(14) & "" = "" Then EventArray(14) = "y"
	If EventArray(15) & "" = "" Then EventArray(15) = "n"
    If EventArray(30) & "" = "" Then EventArray(30) = "n"
    If EventArray(31) & "" = "" Then EventArray(31) = "n"
	If EventArray(32) & "" = "" Then EventArray(32) = "0"

	'get site web info
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT WelcomeMsg FROM EventsWeb WHERE EventsID = " & lEventID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		If Not rs(0).Value & "" = "" Then sWelcome = Replace(rs(0).Value, "''", "'")
	End If
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

'get event types
i = 0
ReDim EventGrps(0)
sql = "SELECT DISTINCT EventGrp FROM Events ORDER BY EventGrp DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventGrps(i) = rs(0).Value
	i = i + 1
	ReDim Preserve EventGrps(i)
	rs.MoveNext
Loop
Set rs = Nothing

EventGrps(i) = CInt(EventGrps(i - 1)) + 1

sRsltsOfficial = "n"
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID FROM OfficialRslts WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then sRsltsOfficial = "y"
rs.Close
Set rs = Nothing

'get rfid settings
bFound = True
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT StartTime, DuplRange, MinTime, MaxStart,  RsltsSort FROM RFIDSettings WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then 
    sStartTime = rs(0).Value
    sngDuplRange = rs(1).Value
    sngMinTime = rs(2).Value
    iMaxStart = rs(3).Value
    sRsltsSort = rs(4).Value
Else
    bFound = False
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceReport, Weather, Gallery FROM RaceReport WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    If Not rs(0).Value & "" = "" Then sRaceReport = Replace(rs(0).Value, "''", "'")
    If Not rs(1).Value & "" = "" Then sWeather = Replace(rs(1).Value, "''", "'")
    If Not rs(2).Value & "" = "" Then sGallery = Replace(rs(2).Value, "''", "'")
End If
rs.Close
Set rs = Nothing

If bFound = False Then
    iMaxStart = "300"
    sStartTime = "00:00:00.000"
    sngDuplRange = "0"
    sngMinTime = "0"
        
    sql = "INSERT INTO RFIDSettings(EventID) VALUES (" & lEventID & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If

Private Function GetThisType(lEventType)
	sql = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs = conn.Execute(sql)
	GetThisType = rs(0).Value
	Set rs = Nothing
End Function

Private Function EventDirEmail(lEventDir)
	sql = "SELECT Email FROM EventDir WHERE EventDirID = " & lEventDir
	Set rs = conn.Execute(sql)
	EventDirEmail = rs(0).Value
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title><%=Replace(sEventName, "''", "'")%> Event Info</title>

<script>
function checkFields() {
 	if (document.update_info.event_name.value == '')
		{
  		alert('Please fill in all required fields!');
  		return false
  		}
	else
		if (isNaN(document.update_info.date_month.value) ||
		   isNaN(document.update_info.date_day.value) ||
		   isNaN(document.update_info.date_year.value))
    		{
			alert('All event date fields must be numeric values');
			return false
			} 	
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h3 class="h3">Edit Event Information</h3>
			
			<form class="form-inline" name="which_event" method="post" action="edit_event.asp?event_id=<%=lEventID%>">
			<label for="events">Events:</label>
			<select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Events, 2)%>
					<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%Else%>
						<option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			<input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
			</form>
			<br>

			<%If Not Clng(lEventID) = 0 Then%>
				<!--#include file = "../../includes/event_nav.asp" -->
				
				<div style="text-align:right;padding-right: 10px;">
					<form class="form-inline" name="rsults_official" method="post" action="edit_event.asp?event_id=<%=lEventID%>">
					<label for="official_rslts">Rslts Official:&nbsp;</label>
					<select class="form-control" name="official_rslts" id="official_rslts">
						<%If sRsltsOfficial = "n" Then%> 
							<option value="n" selected>n</option>
							<option value="y">y</option>
						<%Else%>
							<option value="n">n</option>
							<option value="y" selected>y</option>
						<%End If%>
					</select>
					<input type="hidden" name="submit_official" id="submit_official" value="submit_official">
					<input class="form-control" type="submit" name="submit_official_event" id="submit_official_event" value="Submit Results As Official">
					</form>
				</div>
				
                <div class="bg-danger">
                    <a href="clone_event.asp?event_id=<%=lEventID%>" style="color:#fff;">Clone Existing Event</a>
                    |
                    <a style="color:#fff;" href="javascript:pop('http://www.gopherstateevents.com/events/raceware_events.asp?event_id=<%=lEventID%>',800,600)">Info Link</a>
                </div>
				    
				<form class="form" name="update_info" method="post" action="edit_event.asp?event_id=<%=lEventID%>" onsubmit="return checkFields()">
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="event_name">Event Name:&nbsp;</label>
					<div class="col-sm-4"><input class="form-control" name="event_name" id="event_name" maxlength="25" value="<%=EventArray(0)%>"></div>
					<label class="col-sm-2 form-control-label text-nowrap" for="event_type">Event Type:</label>
					<div class="col-sm-4">
						<select class="form-control" name="event_type" id="event_type">
							<%For i = 0 to UBound(EventTypes, 2) - 1%>
								<%If CLng(EventArray(3)) = CLng(EventTypes(0, i)) Then%>
									<option value="<%=EventTypes(0, i)%>" selected><%=EventTypes(1, i)%></option>
								<%Else%>
									<option value="<%=EventTypes(0, i)%>"><%=EventTypes(1, i)%></option>
								<%End If%>
							<%Next%>
						</select>
					</div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="date_month"><span style="color:#d62002">*</span>Event Date:</label>
					<div class="col-sm-4">
						<input class="form-control" name="date_month" id="date_month" maxLength="2" value="<%=Month(EventArray(1))%>">&nbsp;
						<input class="form-control" name="date_day" id="date_day" maxLength="2" value="<%=Day(EventArray(1))%>">&nbsp;
						<input class="form-control" name="date_year" id="date_year" maxLength="4" value="<%=Year(EventArray(1))%>">
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="location">Location:</label>
					<div class="col-sm-4"><input class="form-control" type="text" name="location" id="location" value="<%=EventArray(29)%>"></div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="admin_notes">Admin Notes:</label>
					<div class="col-sm-10"><textarea class="form-control" name="admin_notes" id="admin_notes" rows="4"><%=EventArray(22)%></textarea></div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="staff_notes">Staff Notes:</label>
					<div class="col-sm-10"><textarea class="form-control" name="staff_notes" id="staff_notes" rows="4"><%=EventArray(36)%></textarea></div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="club">Club:</label>
					<div class="col-sm-4"><input class="form-control" name="club" id="club" maxlength="50" value="<%=EventArray(4)%>"></div>
					<label class="col-sm-2 form-control-label text-nowrap" for="ant_field_size">Antic Size:</label>
					<div class="col-sm-4"><input class="form-control" type="text" name="ant_field_size" id="ant-field_size" maxlength="4" value="<%=EventArray(32)%>"></div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="need_truss">Need Truss:</label>
					<div class="col-sm-4">
						<select class="form-control" name="need_truss" id="need_truss">
							<%If EventArray(34) = "y" Then%>
								<option value="y" selected>Yes</option>
								<option value="n">No</option>
							<%Else%>
								<option value="y">Yes</option>
								<option value="n" selected>No</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="event_family">Event Family:</label>
					<div class="col-sm-4">
						<select class="form-control" name="event_family" id="event_family">
							<option value="0">&nbsp;</otpion>
							<%For i = 0 To UBound(EventFams, 2) - 1%>
								<%If CLng(EventArray(20)) = CLng(EventFams(0, i)) Then%>
									<option value="<%=EventFams(0, i)%>" selected><%=EventFams(1, i)%></otpion>
								<%Else%>
									<option value="<%=EventFams(0, i)%>"><%=EventFams(1, i)%></otpion>
								<%End IF%>
							<%Next%>
						</select>
					</div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="announcer">Announcer:</label>
					<div class="col-sm-4">
						<select class="form-control" name="announcer" id="announcer">
							<%If EventArray(27) = "y" Then%>
								<option value="n">No</option>
								<option value="y" selected>Yes</option>
							<%Else%>
								<option value="n">No</option>
								<option value="y">Yes</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="event_grp">Event Group:</label>
					<div class="col-sm-4">
						<select class="form-control" name="event_grp" id="event_grp">
							<option value="<%=CInt(EventGrps(0)) + 1%>"><%=CInt(EventGrps(0)) + 1%></option>
							<%For i = 0 To UBound(EventGrps)%>
								<%If CInt(EventArray(16)) = CInt(EventGrps(i)) Then%>
									<option value="<%=EventGrps(i)%>" selected><%=EventGrps(i)%></option>
								<%Else%>
									<option value="<%=EventGrps(i)%>"><%=EventGrps(i)%></option>
								<%End If%>
							<%Next%>
						</select>
					</div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="start_box">Extra Box:</label>
					<div class="col-sm-4">
						<select class="form-control" name="start_box" id="start_box">
							<%If EventArray(26) = "y" Then%>
								<option value="n">No</option>
								<option value="y" selected>Yes</option>
							<%Else%>
								<option value="n">No</option>
								<option value="y">Yes</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="edition">Edition:</label>
					<div class="col-sm-4">
						<select class="form-control" name="edition" id="edition">
							<%For i = 0 to UBound(Events, 2)%>
								<%If CInt(EventArray(17)) = CInt(i) Then%>
									<option value="<%=i%>" selected><%=i%></option>
								<%Else%>
									<option value="<%=i%>"><%=i%></option>
								<%End If%>
							<%Next%>
						</select>
					</div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="digital_display">Digital Display:</label>
					<div class="col-sm-4">
						<select class="form-control" name="digital_display" id="digital_display">
							<%If EventArray(28) = "y" Then%>
								<option value="n">No</option>
								<option value="y" selected>Yes</option>
							<%Else%>
								<option value="n">No</option>
								<option value="y">Yes</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="event_class">Event Class:</label>
					<div class="col-sm-4">
						<select class="form-control" name="event_class" id="event_class">
							<option value="B">B</otpion>
							<%If EventArray(21) = "A" Then%>
								<option value="A" selected>A</otpion>
							<%Else%>
								<option value="A">A</otpion>
							<%End IF%>
						</select>
					</div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="dynamic_bib_assign">Dynamic Assign:</label>
					<div class="col-sm-4">
						<select class="form-control" name="dynamic_bib_assign" id="dynamic_bib_assign">
							<%If EventArray(33) = "y" Then%>
								<option value="y" selected>Yes</option>
								<option value="n">No</option>
							<%Else%>
								<option value="y">Yes</option>
								<option value="n" selected>No</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="event_dir">Event Dir:</label>
					<div class="col-sm-4">
						<select class="form-control" name="event_dir" id="event_dir">
							<%For i = 0 To UBound(EventDir, 2) - 1%>
								<%If CLng(EventDir(0, i)) = CLng(EventArray(19)) Then%>
									<option value="<%=EventDir(0, i)%>" selected><%=EventDir(1, i)%></option>
								<%Else%>
									<option value="<%=EventDir(0, i)%>"><%=EventDir(1, i)%></option>
								<%End If%>
							<%Next%>
						</select>
					</div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="opt_out">Email Opt Out:</label>
					<div class="col-sm-4">
						<select class="form-control" name="opt_out" id="opt_out">
							<%If EventArray(24) = "y" Then%>
								<option value="y" selected>Yes</option>
								<option value="n">No</option>
							<%Else%>
								<option value="y">Yes</option>
								<option value="n" selected>No</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap">Email:</label>
					<div class="col-sm-4">
						<a href="mailto:<%=EventDirEmail(EventArray(19))%>"><%=EventDirEmail(EventArray(19))%></a>
					</div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="coach_first">
						Pix Sponsor:
					</label>
					<div class="col-sm-4">
						<select class="form-control" name="pix_sponsor" id="pix_sponsor">
							<%If EventArray(35) = "y" Then%>
								<option value="n">No</option>
								<option value="y" selected>Yes</option>
							<%Else%>
								<option value="n">No</option>
								<option value="y">Yes</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="found_by">Found By:</label>
					<div class="col-sm-4"><input class="form-control" type="text" name="found_by" id="found_by" value="<%=EventArray(23)%>"></div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="six_digits_only">6-Digit Bibs Only:</label>
					<div class="col-sm-4">
						<select class="form-control" name="six_digits_only" id="six_digits_only">
							<%If EventArray(37) = "y" Then%>
								<option value="y" selected>Yes</option>
								<option value="n">No</option>
							<%Else%>
								<option value="y">Yes</option>
								<option value="n" selected>No</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="invoice">Invoice:</label>
					<div class="col-sm-4">$&nbsp;<input class="form-control" type="text"name ="invoice" id="invoice" value="<%=EventArray(18)%>"></div>
				</div>
				<div class="row">
					<div class="col-sm-12">
						<h5 class="h5">RFID Settings</h5>
						<div class="form-group row">
							<label class="col-sm-2 form-control-label text-nowrap" for="start_time">Start Time:</label>
							<div class="col-sm-2"><input class="form-control" type="text" name="start_time" id="start_time" value="<%=sStartTime%>" size="12"></div>
							<label class="col-sm-2 form-control-label text-nowrap" for="max_start">Max Start:</label>
							<div class="col-sm-2"><input class="form-control" type="text" name="max_start" id="max_start" value="<%=iMaxStart%>" size="3"></div>
							<label class="col-sm-2 form-control-label text-nowrap" for="dupl_range">Duplicate Range:</label>
							<div class="col-sm-2"><input class="form-control" type="text" name="dupl_range" id="dupl_range" value="<%=sngDuplRange%>" size="4"></div>
						</div>
						<div class="form-group row">
							<label class="col-sm-2 form-control-label text-nowrap" for="min_time">Min Time:</label>
							<div class="col-sm-2"><input class="form-control" type="text" name="min_time" id="min_time" value="<%=sngMinTime%>" size="3"></div>
							<label class="col-sm-2 form-control-label text-nowrap" for="rslts_sort">Sort By If Chip Start:</label>
							<div class="col-sm-4">
								<select class="form-control" name="rslts_sort" id="rslts_sort">
									<%If sRsltsSort = "chip" Then%>
										<option value="chip" selected>Chip Time</option>
										<option value="gun">Gun Time</option>
									<%Else%>
										<option value="chip">Chip Time</option>
										<option value="gun" selected>Gun Time</option>
									<%End If%>
								</select>
							</div>
						</div>
					</div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="event_site">Event Site:</label>
					<div class="col-sm-10"><textarea class="form-control" name="event_site" id="event_site" rows="3"><%=EventArray(2)%></textarea></div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="map_link">Map to Site:</label>
					<div class="col-sm-10"><textarea class="form-control" name="map_link" id="map_link" rows="3"><%=EventArray(10)%></textarea></div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="weather">Weather:</label>
					<div class="col-sm-10"><textarea class="form-control" name="weather" id="weather" rows="5"><%=EventArray(6)%></textarea></div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="web_site">Website:</label>
					<div class="col-sm-10"><input class="form-control" name="web_site" id="web_site" maxlength="500" value="<%=EventArray(5)%>"></div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="local_power">Local Power? &nbsp;</label>
					<div class="col-sm-2">
						<%If EventArray(31) = "y" Then%>
							<input type="radio" name="local_power" id="local_power" value="y" checked>Yes &nbsp;
							<input type="radio" name="local_power" id="local_power" value="n">No
						<%Else%>
							<input type="radio" name="local_power" id="local_power" value="y">Yes &nbsp;
							<input type="radio" name="local_power" id="local_power" value="n" checked>No
						<%End If%>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="show_online"><span style="color:#d62002">*</span>Show Online? &nbsp;</label>
					<div class="col-sm-2">
						<%If EventArray(8) = "y" Then%>
							<input type="radio" name="show_online" id="show_online" value="y" checked>Yes &nbsp;
							<input type="radio" name="show_online" id="show_online" value="n">No
						<%Else%>
							<input type="radio" name="show_online" id="show_online" value="y">Yes &nbsp;
							<input type="radio" name="show_online" id="show_online" value="n" checked>No
						<%End If%>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="online_reg"><span style="color:#d62002">*</span>Online Part Reg? &nbsp;</label>
					<div class="col-sm-2">
						<%If EventArray(9) = "y" Then%>
							<input type="radio" name="online_reg" id="online_reg" value="y" checked>Yes &nbsp;
							<input type="radio" name="online_reg" id="online_reg" value="n">No
						<%Else%>
							<input type="radio" name="online_reg" id="online_reg" value="y">Yes &nbsp;
							<input type="radio" name="online_reg" id="online_reg" value="n" checked>No
						<%End If%>
					</div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="need_bibs">Need Bibs? &nbsp;</label>
					<div class="col-sm-2">
						<%If EventArray(14) = "y" Then%>
							<input type="radio" name="need_bibs" id="need_bibs" value="y" checked>Yes &nbsp;
							<input type="radio" name="need_bibs" id="need_bibs" value="n">No
						<%Else%>
							<input type="radio" name="need_bibs" id="need_bibs" value="y">Yes &nbsp;
							<input type="radio" name="need_bibs" id="need_bibs" value="n" checked>No
						<%End If%>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="need_pins">Need Pins? &nbsp;</label>
					<div class="col-sm-2">
						<%If EventArray(15) = "y" Then%>
							<input type="radio" name="need_pins" id="need_pins" value="y" checked>Yes &nbsp;
							<input type="radio" name="need_pins" id="need_pins" value="n">No
						<%Else%>
							<input type="radio" name="need_pins" id="need_pins" value="y">Yes &nbsp;
							<input type="radio" name="need_pins" id="need_pins" value="n" checked>No
						<%End If%>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="tear_offs">Tear-Offs? &nbsp;</label>
					<div class="col-sm-2">
						<%If EventArray(31) = "y" Then%>
							<input type="radio" name="tear_offs" id="tear_offs" value="y" checked>Yes &nbsp;
							<input type="radio" name="tear_offs" id="tear_offs" value="n">No
						<%Else%>
							<input type="radio" name="tear_offs" id="tear_offs" value="y">Yes &nbsp;
							<input type="radio" name="tear_offs" id="tear_offs" value="n" checked>No
						<%End If%>
					</div>
				</div>
				<div class="form-group row">
					<label class="col-sm-2 form-control-label text-nowrap" for="race_report">Race Report:</label>
					<div class="col-sm-4"><textarea class="form-control" name="race_report" id="race_report" rows="4"><%=sRaceReport%></textarea></div>
					<label class="col-sm-2 form-control-label text-nowrap" for="gallery">Race Gallery:</label>
					<div class="col-sm-4"><textarea class="form-control" name="gallery" id="gallery" rows="4"><%=sGallery%></textarea></div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="waiver">Event Waiver:</label>
					<div class="col-sm-4"><textarea class="form-control" name="waiver" id="waiver" rows="5"><%=EventArray(11)%></textarea></div>
					<label class="col-sm-2 form-control-label text-nowrap" for="comments">Comments:</label>
					<div class="col-sm-4"><textarea class="form-control" name="comments" id="comments" rows="5"><%=EventArray(7)%></textarea></div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="wlcme_msg">Welcome:</label>
					<div class="col-sm-4"><textarea class="form-control" name="wlcme_msg" id="wlcme_msg" rows="5"><%=sWelcome%></textarea></div>
					<label class="col-sm-2 form-control-label text-nowrap" for="address">Address:</label>
					<div class="col-sm-4"><textarea class="form-control" name="address" id="address" rows="3"><%=EventArray(12)%></textarea></div>
				</div>
				<div class="form-group row">	
					<label class="col-sm-2 form-control-label text-nowrap" for="shutdown_month">End Pre-Reg:</label>
					<div class="col-sm-4">
						<select class="form-control" name="shutdown_month" id="shutdown_month">
							<%For i = 1 To 12%>
								<%If Month(EventArray(13)) = i Then%>
									<option value="<%=i%>" selected><%=i%></option>
								<%Else%>
									<option value="<%=i%>"><%=i%></option>
								<%End If%>
							<%Next%>
						</select>
						/
						<select class="form-control" name="shutdown_day" id="shutdown_day">
							<%For i = 1 To 31%>
								<%If Day(EventArray(13)) = i Then%>
									<option value="<%=i%>" selected><%=i%></option>
								<%Else%>
									<option value="<%=i%>"><%=i%></option>
								<%End If%>
							<%Next%>
						</select>
						/
						<select class="form-control" name="shutdown_year" id="shutdown_year">
							<%For i = 2001 To Year(Date) + 1%>
								<%If Year(EventArray(13)) = i Then%>
									<option value="<%=i%>" selected><%=i%></option>
								<%Else%>
									<option value="<%=i%>"><%=i%></option>
								<%End If%>
							<%Next%>
						</select>
						at
						<select class="form-control" name="shutdown_hour" id="shutdown_hour">
							<%For i = 1 To 24%>
								<%If i < 10 Then%>
									<%If Hour(EventArray(13)) = i Then%>
										<option value="0<%=i%>" selected>0<%=i%></option>
									<%Else%>
										<option value="0<%=i%>">0<%=i%></option>
									<%End If%>
								<%ElseIf i > 12 Then%>
									<%If Hour(EventArray(13)) = i Then%>
										<option value="<%=i - 12%>" selected><%=i - 12%></option>
									<%Else%>
										<option value="<%=i - 12%>"><%=i - 12%></option>
									<%End If%>
								<%Else%>
									<%If Hour(EventArray(13)) = i Then%>
										<option value="<%=i%>" selected><%=i%></option>
									<%Else%>
										<option value="<%=i%>"><%=i%></option>
									<%End If%>
								<%End If%>
							<%Next%>
						</select>
						:00.00
						<select class="form-control" name="shutdown_ampm" id="shutdown_ampm">
							<%If Right(CStr(EventArray(13)), 2) = "PM" Then%>
								<option value="AM">AM</option>
								<option value="PM" selected>PM</option>
							<%Else%>
								<option value="AM">AM</option>
								<option value="PM">PM</option>
							<%End If%>
						</select>
					</div>
					<label class="col-sm-2 form-control-label text-nowrap" for="packet_pickup">Packet Pickup:</label>
					<div class="col-sm-4"><textarea class="form-control" name="packet_pickup" id="packet_pickup" rows="3"><%=EventArray(25)%></textarea></div>
				</div>
				<div class="form-group row">
					<div class="col-sm-12">
						<input type="checkbox" name="delete" id="delete">&nbsp;Delete This Event! (There is no undo for this action and if any results
						have been recorded they will be deleted!)
					</div>
				</div>
				<div class="form-group row">
					<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
					<input class="form-control" type="submit" name="submit" id="submit" value="Make Changes">
				</div>
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
%>