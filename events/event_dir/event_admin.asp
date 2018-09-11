<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim sEventName, sEventSite, sClub, sWebSite, sWeather, sComments, sOnlineReg, sWaiver, sStartBox, sAnnouncer, sDigitalDisplay, sLocation, sLocalPower
Dim sNeedBibs, sNeedPins, sMapLink, sAddress, sOptOut, sRsltsSort, sStartTime, sPacketPickup, sWhichTab, sInfoLink, sThisPage, sLogo, sInfoSheet
Dim sTearOffs, sRaceReport, sGallery
Dim iAntFieldSize
Dim EventArray(22), Events()
Dim sngDeposit
Dim dWhenShutdown, dEventDate
Dim bFound, bChangesLocked

If Not Session("role") = "event_dir" Then Response.Redirect "/default.asp?sign_out=y"

sThisPage = "event_admin.asp"
lEventID = Request.QueryString("event_id")

sWhichTab = Request.QueryString("which_tab")
If sWhichTab = vbNullString Then sWhichTab = "General"

sngDeposit = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_post_race") = "submit_post_race" Then
	If Not Request.Form.Item("weather") & "" = "" Then sWeather = Replace(Request.Form.Item("weather"), "'", "''")
    If Not Request.Form.Item("race_report") & "" = "" Then sRaceReport = Replace(Request.Form.Item("race_report"), "'", "''")
    If Not Request.Form.Item("gallery") & "" = "" Then sGallery = Replace(Request.Form.Item("gallery"), "'", "''")

    bFound = False
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RaceReport, Weather, Gallery FROM RaceReport WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
        rs(0).value = sRaceReport
        rs(1).value = sWeather
        If Not sGallery & "" = "" Then rs(2).value = sGallery
        bFound = True
	    rs.Update
    End If
	rs.Close
	Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO RaceReport (EventID, RaceReport, Weather, Gallery) VALUES (" & lEventID & ", '" & sRaceReport & "', '"
        sql = sql & sWeather & "', '" & sGallery & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_registration") = "submit_registration" Then
	sOnlineReg = Request.Form.Item("online_reg")
	dWhenShutDown = Request.Form.Item("when_shutdown")


	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT OnlineReg, WhenShutdown FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	rs(0).value = sOnlineReg
	If IsDate(dWhenShutdown) Then
		rs(1).Value = dWhenShutdown
	Else
		rs(1).Value = rs(1).OriginalValue
	End If
	rs.Update
	rs.Close
	Set rs = Nothing
ElseIf Request.Form.Item("submit_venue") = "submit_venue" Then
	If Not Request.Form.Item("event_site") & "" = "" Then sEventSite = Replace(Request.Form.Item("event_site"), "'", "''")
	If Not Request.Form.Item("map_link") & "" = "" Then sMapLink = Replace(Request.Form.Item("map_link"), "'", "''")
	If Not Request.Form.Item("address") & "" = "" Then sAddress = Replace(Request.Form.Item("address"), "'", "''")
    sLocation = Request.Form.Item("location")
    sLocalPower = Request.Form.Item("local_power")

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Location, LocalPower FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
    rs(0).Value = sLocation
    rs(1).Value = sLocalPower
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
ElseIf Request.Form.Item("submit_general") = "submit_general" Then
	If Not Request.Form.Item("event_name") & "" = "" Then sEventName = Replace(Request.Form.Item("event_name"), "'", "''")
	dEventDate = Request.Form.Item("event_date")
	If Not Request.Form.Item("club") & "" = "" Then sClub = Replace(Request.Form.Item("club"), "'", "''")
	If Not Request.Form.Item("web_site") & "" = "" Then sWebSite = Replace(Request.Form.Item("web_site"), "'", "''")
	If Not Request.Form.Item("comments") & "" = "" Then sComments = Replace(Request.Form.Item("comments"), "'", "''")
	If Not Request.Form.Item("waiver") & "" = "" Then sWaiver = Replace(Request.Form.Item("waiver"), "'", "''")
    If Not Request.Form.Item("packet_pickup") & "" = "" Then sPacketPickup = Replace(Request.Form.Item("packet_pickup"), "'", "''")
    iAntFieldSize = Request.Form.Item("ant_field_size")

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventName, EventDate, Club, Website, Comments, PacketPickup, AntFieldSize FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	If sEventName & "" = "" Then
		rs(0).Value = Replace(rs(0).OriginalValue, "''", "'")
	Else
		rs(0).value = sEventName
	End If
	rs(1).value = dEventDate
	rs(2).value = sClub
	rs(3).value = sWebSite
	rs(4).value = Left(sComments, 2000)
    rs(5).value = sPacketPickup
    rs(6).Value = iAntFieldSize
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
ElseIf Request.Form.Item("submit_preferences") = "submit_preferences" Then
	sNeedBibs = Request.Form.Item("need_bibs")
	sNeedPins = Request.Form.Item("need_pins")
    sOptOut = Request.Form.Item("opt_out")
    sAnnouncer = Request.Form.Item("announcer")
    sDigitalDisplay = Request.Form.Item("digital_display")
    sRsltsSort = Request.Form.Item("rslts_sort")

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT NeedBibs, NeedPins, OptOut, Announcer, DigitalDisplay FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	rs(0).value = sNeedBibs
	rs(1).value = sNeedPins
    rs(2).value = sOptOut
    rs(3).value = sAnnouncer
    rs(4).value = sDigitalDisplay
	rs.Update
	rs.Close
	Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RsltsSort FROM RFIDSettings WHERE EventID = " & lEventID
    rs.Open sql, conn,  1, 2
    rs(0).Value = sRsltsSort
    rs.Update
    rs.Close
    Set rs = Nothing
End If

i = 0
ReDim Events(1, 0)
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDirID = " & Session("my_id") & " ORDER By EventDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.eOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
Set rs = Nothing
	
If UBound(Events, 2) = 1 Then lEventID = Events(0, 0)
If CStr(lEventID) = vbNullString Then lEventID = Events(0, 0)

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventName, e.EventDate, s.SiteName, e.Club, e.Website, e.Weather, e.Comments, e.OnlineReg, s.MapLink, w.Waiver, s.Address, e.WhenShutdown, "
sql = sql & "e.NeedBibs, e.NeedPins, e.OptOut, e.PacketPickup, e.Announcer, e.DigitalDisplay, e.Location, e.LocalPower, e.TearOffs, e.AntFieldSize, "
sql = sql & "e.Deposit, e.Logo FROM Events e INNER JOIN SiteInfo s ON e.EventID = s.EventID INNER JOIN Waiver w "
sql = sql & "ON w.EventID = e.EventID WHERE e.EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
	For i = 0 to 21
		If Not rs(i).Value & "" = "" Then  EventArray(i) = Replace(rs(i).Value, "''", "'")
	Next

    sngDeposit = rs(22).Value
    sLogo = rs(23).Value
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RsltsSort FROM RFIDSettings WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
EventArray(22) = rs(0).Value
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
	
bChangesLocked = False
If Date >= CDate(EventArray(1)) - 5 Then bChangesLocked = True

If EventArray(12) & "" = "" Then EventArray(12) = "y"
If EventArray(13) & "" = "" Then EventArray(13) = "n"
If EventArray(21) & "" = "" Then EventArray(21) = "50"

If EventArray(4) & "" = "" Then
    sInfoLink = "http://www.gopherstateevents.com/events/raceware_events.asp?event_id=" & lEventID
Else
    sInfoLink = EventArray(4)
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT InfoSheet FROM InfoSheet WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then sInfoSheet = rs(0).Value
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=Replace(EventArray(0), "''", "'")%> Event Info</title>

<!--#include file = "event_css.asp" -->

<script>
function chkGeneral() {
 	if (document.update_general.event_name.value == '' || 
	 	document.update_general.event_date.value == '' ||
	 	document.update_general.event_site.value == '')
		{
  		alert('Please fill in all required fields!');
  		return false
  		}
	else
   		return true
}

function chkRegistration() {
	if (isNaN(document.update_registration.date_month.value) ||
		isNaN(document.update_registration.date_day.value) ||
		isNaN(document.update_registration.date_year.value))
    	{
		alert('All date fields must be numeric values');
		return false
		} 	
	else
   		return true
}

$(function() {
    $( "#event_date" ).datepicker({
      autoclose: true
    });
}); 


$(function() {
    $( "#when_shutdown" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/event_dir_menu.asp" -->
		<div class="col-sm-10">
			<h3 class="h3">GSE Edit/Manage Event Information: <%=EventArray(0)%></span></h3>
		
			<!--#include file = "event_select.asp" -->
										
			<div class="row">
				<!--#include file = "event_dir_tabs.asp" -->
				<%Select Case sWhichTab%>
					<%Case "General"%>
						<form class="form" role="form" name="update_general" 
							method="post" action="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General" 
							onsubmit="return chkGeneral()">
						<div class="row">	
							<div class="col-sm-2">Event Name:</div>
							<div class="col-sm-4"><input class="form-control" name="event_name" id="event_name" maxlength="25" size="45" value="<%=EventArray(0)%>"></div>
							<div class="col-sm-6">&nbsp;</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Event Date:</div>
							<div class="col-sm-4">
								<input class="form-control" name="event_date" id="event_date" value="<%=EventArray(1)%>">
							</div>
							<div class="col-sm-6">&nbsp;</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Field Size:</div>
							<div class="col-sm-4">
								<input class="form-control" name="ant_field_size" id="ant_field_size" maxlength="4" size="4" value="<%=EventArray(21)%>">
							</div>
							<div class="col-sm-6">
								Roughly how many participants are you expecting. <span style="color: red;">(must be numeric)</span>
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Organization:</div>
							<div class="col-sm-4"><input class="form-control" name="club" id="club" maxlength="50" size="50" value="<%=EventArray(3)%>"></div>
							<div class="col-sm-6">
								The club, group, or organization that is sponsoring this event.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Website:</div>
							<div class="col-sm-4"><input class="form-control" name="web_site" id="web_site" maxlength="500" size="65" value="<%=EventArray(4)%>"></div>
							<div class="col-sm-6">
								If your event has it's own website please enter it here (include "http:").  If you enter a website for your
								event it will be displayed when visitors click on your event on our calendar.  If not, a GSE event info page will
								be displayed.
							</div>
						</div>
						<div class="row">	
							<div class="col-sm-2">Event Waiver:</div>
							<div class="col-sm-4"><textarea class="form-control" name="waiver" id="waiver" rows="5"><%=EventArray(9)%></textarea></div>
							<div class="col-sm-6">
								We supply a generic waiver in the event that you might need us to print some registration forms.  This rarely
								happens.  If you wish to supply your own waiver, please enter it to the left.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Comments:</div>
							<div class="col-sm-4"><textarea class="form-control" name="comments" id="comments" rows="5"><%=EventArray(6)%></textarea></div>
							<div class="col-sm-6">
								This information will show up on your page's info page.  The info page only appears if you do not supply a website
								for your event.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Packet Pickup:</div>
							<div class="col-sm-4"><textarea class="form-control" name="packet_pickup" id="packet_pickup" rows="10"><%=EventArray(15)%></textarea></div>
							<div class="col-sm-6">
								This information will appear on your event's info page if you do not supply a website.  We will also use it to
								guide us if you want us at packet pick-up.  We can send a representative to packet pick-up to enter walk-up
								registrations if you wish.  The fee for that is round trip mileage and $25/hour.  NOTE:  If you expect a significant
								number of walk-up registrations at packet pick-up we would like to get that information entered into our system
								prior to race morning.
							</div>
						</div>
						<div class="row">
							<input type="hidden" name="submit_general" id="submit_general" value="submit_general">
							<%If bChangesLocked = False Then%>
								<input class="form-control" type="submit" name="submit1" id="submit1" value="Save Changes">
							<%Else%>
								<input class="form-control" type="submit" name="submit1" id="submit1" value="Save Changes" disabled>
							<%End If%>
						</div>
						</form>
					<%Case "Venue"%>
						<form class="form" role="form" name="update_venue" method="post" action="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">
						<div class="row">
							<div class="col-sm-2">City, St:</div>
							<div class="col-sm-4"><input class="form-control" type="text" name="location" id="location" value="<%=EventArray(18)%>"></div>
							<div class="col-sm-6">
								Please supply this information for geo-location purposes.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Event Site:</div>
							<div class="col-sm-4"><textarea class="form-control" name="event_site" id="event_site" rows="5"><%=EventArray(2)%></textarea></div>
							<div class="col-sm-6">
								This field is for a site label such as "City Park" "Behind the High School" or something to that effect.  It may 
								help us and your participants find the venue.  If your event does not provide a website this information
								will appear on your event's information page.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Address:</div>
							<div class="col-sm-4"><textarea class="form-control" name="address" id="address" rows="5"><%=EventArray(10)%></textarea></div>
							<div class="col-sm-6">
								Please supply the physical address for your venue.  If your event does not provide a website this information
								will appear on your event's information page.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Map to Site:</div>
							<div class="col-sm-4"><input class="form-control" type="text" name="map_link" id="map_link" value="<%=EventArray(8)%>"></div>
							<div class="col-sm-6">
								Please supply a MapQuest or Google Maps link to the exact site for your event.  If your event doesn't provide a 
								website this will appear on your GSE information page.  It will also be useful as our staff
								attempts to find your event on race day.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Local Power:</div>
							<div class="col-sm-4">
								<select class="form-control" name="local_power" id="local_power">
									<%If EventArray(19) = "y" Then%>
										<option value="n">No</option>
										<option value="y" selected>Yes</option>
									<%Else%>
										<option value="n">No</option>
										<option value="y">Yes</option>
									<%End If%>
								</select>
							</div>
							<div class="col-sm-6">
								If the venue has local electrical power available within 100' please indicate this.  We can supply our own power but
								will tap in to local power if available.
							</div>
						</div>
						<div class="row">
							<input type="hidden" name="submit_venue" id="submit_venue" value="submit_venue">
							<%If bChangesLocked = False Then%>
								<input class="form-control" type="submit" name="submit2" id="submit2" value="Save Changes">
							<%Else%>
								<input class="form-control" type="submit" name="submit2" id="submit2" value="Save Changes" disabled>
							<%End If%>
						</div>
						</form>
					<%Case "Preferences"%>
						<form class="form" role="form" name="update_preferences" method="post" action="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">
						<div class="row">
							<div class="col-sm-2">Sort Results By:</div>
							<div class="col-sm-4">
								<select class="form-control" name="rslts_sort" id="rslts_sort">
									<%If EventArray(22) = "chip" Then%>
										<option value="gun">Gun Time</option>
										<option value="chip" selected>Chip Time</option>
									<%Else%>
										<option value="gun">Gun Time</option>
										<option value="chip">Chip Time</option>
									<%End If%>
								</select>
							</div>
							<div class="col-sm-6">
								If your event utilizes a chip start, race results can be sorted by gun time or by "net" time (from when a participant 
								crosses the starting line until they start the finish line).  Note that only gun time is used by most records management
								groups (USATF, etc).  However, most participants expect results to be sorted by net time.  Please indicate your
								preference here.  (See "Race Data" for our guidelines on chip starts.)
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Email Opt Out:</div>
							<div class="col-sm-4">
								<select class="form-control" name="opt_out" id="opt_out">
									<%If EventArray(14) = "y" Then%>
										<option value="y" selected>Yes</option>
										<option value="n">No</option>
									<%Else%>
										<option value="y">Yes</option>
										<option value="n" selected>No</option>
									<%End If%>
								</select>
							</div>
							<div class="col-sm-6">
								Gopher State Events will send out a pre-race email to all participants within 72 hours prior to the event (usually
								the day before).  This is a good way to ensure that names are spelled correctly and other data is accurate.  If
								you select "No" here an email will not be sent to your participants.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Need Bibs? &nbsp;</div>
							<div class="col-sm-4">
								<%If EventArray(12) = "y" Then%>
									<input type="radio" name="need_bibs" id="need_bibs" value="y" checked>Yes &nbsp;
									<input type="radio" name="need_bibs" id="need_bibs" value="n">No
								<%Else%>
									<input type="radio" name="need_bibs" id="need_bibs" value="y">Yes &nbsp;
									<input type="radio" name="need_bibs" id="need_bibs" value="n" checked>No
								<%End If%>
							</div>
							<div class="col-sm-6">
								If you are providing your own bibs, custom or otherwise, select "No".  If that is the case it is imperative that you
								LET US KNOW THE RANGE OF NUMBERS ON YOUR BIBS AT LEAST TWO WEEKS PRIOR TO THE EVENT so that we can ensure that we
								have the necessary RFID tags.  We will need to get the bibs from you in time to apply the RFID tags to them.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Need Pins? &nbsp;</div>
							<div class="col-sm-4">
								<%If EventArray(13) = "y" Then%>
									<input type="radio" name="need_pins" id="need_pins" value="y" checked>Yes &nbsp;
									<input type="radio" name="need_pins" id="need_pins" value="n">No
								<%Else%>
									<input type="radio" name="need_pins" id="need_pins" value="y">Yes &nbsp;
									<input type="radio" name="need_pins" id="need_pins" value="n" checked>No
								<%End If%>
							</div>
							<div class="col-sm-6">
								Gopher State Events can supply pins at a cost of $15/box. Each box accommodates enough pins for about 350
								participants.  Only whole boxes are provided.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Announcer Portal:</div>
							<div class="col-sm-4">
								<select class="form-control" name="announcer" id="announcer">
									<%If EventArray(16) = "y" Then%>
										<option value="n">No</option>
										<option value="y" selected>Yes</option>
									<%Else%>
										<option value="n">No</option>
										<option value="y">Yes</option>
									<%End If%>
								</select>
							</div>
							<div class="col-sm-6">
								The Announcer Portal allows an announcer to view participants as they approach the finish line (or shortly
								after they have finished).  This feature carries an additional fee of $150 and has limited availability.  The
								announcer must provide their own device and sound system.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-2">Digital Display:</div>
							<div class="col-sm-4">
								<select class="form-control" name="digital_display" id="digital_display">
									<%If EventArray(17) = "y" Then%>
										<option value="n">No</option>
										<option value="y" selected>Yes</option>
									<%Else%>
										<option value="n">No</option>
										<option value="y">Yes</option>
									<%End If%>
								</select>
							</div>
							<div class="col-sm-6">
								Gopher State Events, LLC can supply a large-screen monitor and a laptop for viewing results electronically.  The
								cost for this feature is $150.  Alternatively, local event management can provide the laptop and/or monitor at
								no charge.  GSE staff will provide the link to the electronic results page.
							</div>
						</div>
						<div class="row">
							<input type="hidden" name="submit_preferences" id="submit_preferences" value="submit_preferences">
							<%If bChangesLocked = False Then%>
								<input class="form-control" type="submit" name="submit3" id="submit3" value="Save Changes">
							<%Else%>
								<input class="form-control" type="submit" name="submit3" id="submit3" value="Save Changes" disabled>
							<%End If%>
						</div>
						</form>
					<%Case "Registration"%>
						<form class="form" role="form" name="update_registration" method="post" action="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration" 
							onsubmit="return chkRegistration()">
						<div class="row">
							<div class="col-sm-2">Online Part Reg? &nbsp;</div>
							<div class="col-sm-4">
								<%If EventArray(7) = "y" Then%>
									<input type="radio" name="online_reg" id="online_reg" value="y" checked>Yes &nbsp;
									<input type="radio" name="online_reg" id="online_reg" value="n">No
								<%Else%>
									<input type="radio" name="online_reg" id="online_reg" value="y">Yes &nbsp;
									<input type="radio" name="online_reg" id="online_reg" value="n" checked>No
								<%End If%>
							</div>
							<div class="col-sm-6">If you answer "Yes" you can add the links to the pre-registration site on the "Race Info" tab.</div>
						</div>
						<div class="row">	
							<div class="col-sm-2">End Pre-Reg:</div>
							<div class="col-sm-4">
								<input class="form-control" name="when_shutdown" id="when_shutdown" value="<%=EventArray(11)%>">
							</div>
							<div class="col-sm-6">Please follow the data entry protocol carefully or an error may occur.</div>
						</div>
						<div class="row">
							<input type="hidden" name="submit_registration" id="submit_registration" value="submit_registration">
							<%If bChangesLocked = False Then%>
								<input class="form-control" type="submit" name="submit5" id="submit5" value="Save Changes">
							<%Else%>
								<input class="form-control" type="submit" name="submit5" id="submit5" value="Save Changes" disabled>
							<%End If%>
						</div>
						</form>
					<%Case "Post Race"%>
						<p class="descriptor">
							Completing these fields adds a sense of historical reference for your event.  If you supply a race report it will
							show up in our "At The Races" feature.
						</p>

						<form class="form" role="form" name="update_post_race" method="post" action="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">
						<div class="row">
							<div class="col-sm-4">Weather:</div>
							<div class="col-sm-4"><textarea name="weather" id="weather" rows="5" cols="75"><%=sWeather%></textarea></div>
							<div class="col-sm-4">
								Give a brief summary of the weather that day.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-4">Race Report:</div>
							<div class="col-sm-4"><textarea name="race_report" id="race_report" rows="10" cols="75"><%=sRaceReport%></textarea></div>
							<div class="col-sm-4">
								Provide any information on the event from the standpoint of competition, field size, unique attributes of this event,
								or anything else that you would like to record for this year's event.
							</div>
						</div>
						<div class="row">
							<div class="col-sm-4">Race Gallery:</div>
							<div class="col-sm-4"><textarea name="gallery" id="gallery" rows="3" cols="75"><%=sGallery%></textarea></div>
							<div class="col-sm-4">
								If you have taken pix of your event and would like to post a link to them you can enter that here.  Note that the
								link to our finish line pictures will also show up on the "At The Races" summary for your event.</div>
						</div>
						<div class="row">
							<td colspan="3">
								<input type="hidden" name="submit_post_race" id="submit_post_race" value="submit_post_race">
								<input type="submit" name="submit6" id="submit6" value="Save Changes">
							</div>
						</div>
						</form>
					<%Case "Documents"%>
						<div class="row">
							<div class="col-sm-4">Race Logo:</div>
							<div class="col-sm-4"><a href="javascript:pop('logo.asp?event_id=<%=lEventID%>',600,400)">Upload</a></div>
							<div class="col-sm-4">        
								<%If sLogo & "" = "" Then%>
									&nbsp;
								<%Else%>
									<img src="/events/logos/<%=sLogo%>" style="width: 150px;">
								<%End If%>
							</div>
						</div>
						<div class="row">
							<div class="col-sm-4">Race Information Sheet:</div>
							<div class="col-sm-4"><a href="javascript:pop('info_sheet.asp?event_id=<%=lEventID%>',600,400)">Upload</a></div>
							<div class="col-sm-4">
								<%If sInfoSheet & "" = "" Then%>
									&nbsp;
								<%Else%>
									<a href="/events/info_sheets/<%=sInfoSheet%>" onclick="openThis(this.href,1024,768);return false;">View</a>
								<%End If%>
							</div>
						</div>
				<%End Select%>
			</div>
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