<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, sql2, rs2
Dim i, j
Dim lEventID, lEventDirID, lEventType
Dim sEventName, sSiteName, sClub, sWebsite, sShowOnline, sOnlineReg, sComments, sWaiver, sErrMsg, sMsg, sLocation
Dim dThisDate, dEventDate
Dim EventTypes(), EventDirs()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

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

If Request.Form.Item("submit_event") = "submit_event" Then
    sWaiver = "I understand that running a road race is a potentially dangerous activity. I do hereby "
    sWaiver = sWaiver & "waive and release any and all claims for damages that I may incur as a result of my "
    sWaiver = sWaiver & "participation in this event against the event and its organizers, all sponsors, "
    sWaiver = sWaiver & "employees, volunteers, or officials of these organizations. I further certify that have full "
    sWaiver = sWaiver & "knowledge of the risks involved in this event and that I am physically fit and sufficiently "
    sWaiver = sWaiver & "trained to participate. If, however, as a result of my participation in the race I require "
    sWaiver = sWaiver & "medical attention, I hereby give consent to authorize medical personnel to provide "
    sWaiver = sWaiver & "such medical care as deemed necessary.  " & vbCrLf & vbCrLf
    sWaiver = sWaiver & "I have read the foregoing and certify my agreement by clicking the button below. "

	dThisDate = Request.Form.Item("month") & "/" & Request.Form.Item("day") & "/" & Request.Form.Item("year")

	If IsDate(dThisDate) Then
		dEventDate = dThisDate
		sEventName = Replace(Request.Form.Item("event_name"), "'", "''")
		lEventType = Request.Form.Item("event_type")
		If Not Request.Form.Item("club") & "" = "" Then sClub = Replace(Request.Form.Item("club"), "'", "''")
		sWebsite = Request.Form.Item("website")
		sShowOnline = Request.Form.Item("show_online")
		If Not Request.Form.Item("site_name") & "" = "" Then sSiteName = Replace(Request.Form.Item("site_name"), "'", "''")
		If Not Request.Form.Item("comments") & "" = "" Then sComments = Replace(Request.Form.Item("comments"), "'", "''")
		lEventDirID = Request.Form.Item("event_dir")
		sLocation = Request.Form.Item("location")

		If sWebsite & "" = "" Then sWebsite = "http://www.gopherstateevents.com"
		If sSiteName & "" = "" Then sSiteName = "tbd"
		
		sql = "INSERT INTO Events (EventName, EventDate, EventType, Club, Website, ShowOnline, EventDirID, Comments, DateReg, WhenShutdown, "
		sql = sql & "FeeIncrDate, Location) VALUES ('" & sEventName & "', '" & dEventDate & "', " & lEventType & ", '" & sClub & "', '" & sWebsite & "', '" 
		sql = sql & sShowOnline & "', " & lEventDirID & ", '" & sComments & "', '" & Date & "', '" & CDate(dEventDate) - 1  & "', '" & Date 
        sql = sql & "', '" & sLocation & "')"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
		
		'get event id
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT EventID FROM Events WHERE EventName = '" & sEventname & "' AND EventDate = '" & dEventDate & "'"
		rs.Open sql, conn, 1, 2
		lEventID = rs(0).Value
		rs.Close
		Set rs = Nothing
         
         'insert into waiver table
        sql = "INSERT INTO Waiver (EventID, Waiver) VALUES (" & lEventID & ", '" & sWaiver & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
         
         'insert into site info table
        sql = "INSERT INTO SiteInfo (EventID) VALUES (" & lEventID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
         
         'insert into site events web table
        sql = "INSERT INTO EventsWeb (EventsID, MetaDescription) VALUES (" & lEventID & ", 'Meta Description for " & sEventname & " on " & dEventDate & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
         
         'insert into event asgmt table
        sql = "INSERT INTO EventAsgmt (EventID, EventType, EventDate) VALUES (" & lEventID & ", 'fitness', '" & dEventDate & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        'insert into race report
        sql = "INSERT INTO RaceReport (EventID) VALUES (" & lEventID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
	
	    sMsg = vbCrLf & "This is notification that a new event has been added Gopher State Events:" & vbCrLf & vbCrLf
	    sMsg = sMsg & "Event Name: " & sEventName & vbCrLf & vbCrLf
	    sMsg = sMsg & "Event Date: " & dEventDate & vbCrLf & vbCrLf
	    sMsg = sMsg & "Location: " & sLocation & vbCrLf & vbCrLf
		
		Response.Redirect "edit_event.asp?event_id=" & lEventID
	Else
		sErrMsg = "This is not a valid date."

		sEventName = Request.Form.Item("event_name")
		lEventType = Request.Form.Item("event_type")
		sClub = Request.Form.Item("club")
		sWebsite = Request.Form.Item("website")
		sShowOnline = Request.Form.Item("show_online")
		sSiteName = Request.Form.Item("site_name")
		sComments = Request.Form.Item("comments")
		lEventDirID = Request.Form.Item("event_dir")
	End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Create Event</title>

<script>
function chkFlds(){
 	if (document.new_event.event_name.value == '' || 
 	    document.new_event.event_dir.value == '' ||
 	    document.new_event.month.value == '' ||
	 	document.new_event.day.value == '' || 
	 	document.new_event.year.value == '' || 
        document.new_event.location.value == '' || 
	 	document.new_event.event_type.value == '')
		{
  		alert('All fields except event site, club, comments, and website are required.');
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
		<div class="col-md-10">
			<h4 class="h4">Create a New GSE Event</h4>
		
			<%If Not sErrMsg = vbNullString Then%>
				<p><%=sErrMsg%></p>
			<%End If%>
			
			<form name="new_event" method="Post" action="create_event.asp" onsubmit="return chkFlds();">
			<table>
				<tr>
					<th>Event Name:</th>
					<td>
						<input type="text" name="event_name" id="event_name" value="<%=sEventName%>">
					</td>
					<th>Event Director:</th>
					<td>
						<select name="event_dir" id="event_dir">
							<option value="">&nbsp;</option>
							<%For i = 0 To UBound(EventDir, 2) - 1%>
								<option value="<%=EventDir(0, i)%>"><%=EventDir(1, i)%></option>
							<%Next%>
						</select>
					</td>
				</tr>
				<tr>
					<th>Event Date:</th>
					<td>
						<select name="month" id="month">
							<option value="">&nbsp;</option>
							<%For i = 1 To 12%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
						/
						<select name="day" id="day">
							<option value="">&nbsp;</option>
							<%For i = 1 To 31%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
						/
						<select name="year" id="year">
							<option value="">&nbsp;</option>
							<%For i = Year(Date) To Year(Date) + 2%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
					</td>
					<th>Event Type:</th>
					<td>
						<select name="event_type" id="event_type">
							<option value="">&nbsp;</option>
							<%For i = 0 To UBound(EventTypes, 2) - 1%>
								<option value="<%=EventTypes(0, i)%>"><%=EventTypes (1, i)%></option>
							<%Next%>
						</select>
					</td>
				</tr>
				<tr>
					<th>Sponsoring Club/Org:</th>
					<td>
						<input type="text" name="club" id="club" value="<%=sClub%>">
					</td>
					<th>Website:</th>
					<td>
						<input type="text" name="website" id="website" value="<%=sWebsite%>">
					</td>
				</tr>
				<tr>
					<th>Location:</th>
					<td colspan="3">
						<input type="text" name="location" id="location" value="<%=sLocation%>">
					</td>
				</tr>
				<tr>
					<th valign="top">Comments:</th>
					<td colspan="3">
						<textarea name="comments" id="comments" cols="75" rows="3"><%=sComments%></textarea>
					</td>
				</tr>
				<tr>
					<th style="white-space:nowrap;" colspan="2">Show this event online?</th>
					<td colspan="2">
						<%If sShowOnline = "n" Then%>
							<input type="radio" name="show_online" id="show_online" value="y">Yes
							<input type="radio" name="show_online" id="show_online" value="n" checked>No
						<%Else%>
							<input type="radio" name="show_online" id="show_online" value="y" checked>Yes
							<input type="radio" name="show_online" id="show_online" value="n">No
						<%End If%>
					</td>
				</tr>
				<tr>
					<td style="background-color:#ececd8;text-align:center;" colspan="4">
						<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
						<input type="submit" name="submit1" id="submit1" value="Submit Event">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%	
conn.Close
Set conn = Nothing
%>
</body>
</html>
