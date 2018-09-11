<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lThisType, lEventID, lEventDirID
Dim iLastAge, iBibsFrom, iBibsTo
Dim sEventName, dEventDate, sEventSite, sClub, sWebSite, sWeather, sComments, sOnlineReg, sShowOnline, sWaiver, sMapLink, sAddress, sNeedBibs
Dim EventTypes(), EventArray(14), EventDir(11), Events(), Races(), MAgeGrps(), FAgeGrps(), TShirts(8)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0

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
	Events(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
	If CStr(lEventID) = vbNullString Then lEventID = 0
End If

If Not CLng(lEventID) = 0 Then
	sql = "SELECT e.EventName, e.EventType, e.EventDate, e.Club, e.Website, e.WhenShutdown, e.ShowOnline, e.OnlineReg, s.Address, s.MapLink, s.SiteName, "
	sql = sql & "w.Waiver, e.Comments, e.NeedBibs, e.NeedPins, e.EventDirID FROM Events e INNER JOIN SiteInfo s ON e.EventID = s.EventID "
	sql = sql & "INNER JOIN Waiver w ON w.EventID = e.EventID WHERE e.EventID = " & lEventID
	Set rs = conn.Execute(sql)
	For i = 0 to 14
		If Not rs(i).Value & "" = "" Then EventArray(i) = Replace(rs(i).Value, "''", "'")
	Next
	lEventDirID = rs(15).Value
	Set rs = Nothing

	If EventArray(13) & "" = "" Then EventArray(13) = "n"
	If EventArray(14) & "" = "" Then EventArray(14) = "n"
	
	'get event dir for this event
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Phone, Address, City, State, Zip, Email, UserID, Password, Comments, Mobile FROM EventDir WHERE EventDirID = " 
	sql = sql & lEventDirID
	rs.Open sql, conn, 1, 2
	For i = 0 to 11
		If not rs(i).Value & "" = "" Then EventDir(i) =  Replace(rs(i).Value, "''", "'")
	Next
	rs.Close
	Set rs = Nothing
	
	'get races for this event
	i = 0
	ReDim Races(11, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Dist, Type, EntryFeePre, EntryFee, StartTime, Certified, StartType, MAwds, FAwds, RaceID, RaceName, OnlineRegLink FROM RaceData "
	sql = sql & "WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		For j = 0 to 11
			Races(j, i) = rs(j).Value
		Next
			
		i = i + 1
		ReDim Preserve Races(11, i)
			
		rs.MoveNext
	Loop
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

Private Sub MoreRaceData(lRaceID)
	Dim x
	
	'get t-shirt information
	sql = "SELECT IsOption, Small, Medium, Large, XLarge, XXLarge, Short, Long, ChooseNone FROM TShirtData WHERE RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	For x = 0 to 8
		TShirts(x) = rs(x).Value
	Next
	Set rs = Nothing
	
	'get bib range
	iBibsFrom = 0
	iBibsTo = 0
	Set rs=Server.CreateObject("ADODB.Recordset")
    sql = "SELECT BibsFrom, BibsTo FROM RaceData WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    iBibsFrom = rs(0).Value
    iBibsTo = rs(1).Value
    rs.Close
    Set rs = Nothing
End Sub

Private Function GetThisType(lEventType)
	sql = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs = conn.Execute(sql)
	GetThisType = rs(0).Value
	Set rs = Nothing
End Function

Private Sub GetMAgeGrps(lRaceID)
	ReDim MAgeGrps(0)
	k = 0

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EndAge, NumAwds FROM AgeGroups WHERE (Gender = 'm') AND (RaceID = " & lRaceID
	sql = sql & ") ORDER BY EndAge"
	rs.Open sql, conn, 1, 2

	If rs.RecordCount = 1 Then
		MAgeGrps(0) = "None"
			
		k = k + 1
		ReDim Preserve MAgeGrps(k)
	Else
		Do While Not rs.EOF
			If k = 0 Then
				MAgeGrps(k) = rs(0).Value & " and Under (" & rs(1).Value & "), "
				iLastAge = rs(0).Value
			Else
				If rs(0).Value = 110 Then
					MAgeGrps(k) = CInt(iLastAge) + 1 & " and Over (" & rs(1).Value & ")"
				Else
					MAgeGrps(k) = CInt(iLastAge) + 1 & " - " & rs(0).Value & " (" & rs(1).Value & "), "
					iLastAge = rs(0).Value
				End If
			End If
			
			k = k + 1
			ReDim Preserve MAgeGrps(k)
			
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
End Sub

Private Sub GetFAgeGrps(lRaceID)
	ReDim FAgeGrps(0)
	k = 0
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EndAge, NumAwds FROM AgeGroups WHERE (Gender = 'f') AND (RaceID = " & lRaceID
	sql = sql & ") ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	
	If rs.RecordCount = 1 Then
		FAgeGrps(0) = "None"
			
		k = k + 1
		ReDim Preserve FAgeGrps(k)
	Else
		Do While Not rs.EOF
			If k = 0 Then
				FAgeGrps(k) = rs(0).Value & " and Under (" & rs(1).Value & "), "
				iLastAge = rs(0).Value
			Else
				If rs(0).Value = 110 Then
					FAgeGrps(k) = CInt(iLastAge) + 1 & " and Over (" & rs(1).Value & ")"
				Else
					FAgeGrps(k) = CInt(iLastAge) + 1 & " - " & rs(0).Value & " (" & rs(1).Value & "), "
					iLastAge = rs(0).Value
				End If
			End If
			
			k = k + 1
			ReDim Preserve FAgeGrps(k)
			
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Event Data</title>

<!--#include file = "../../includes/js.asp" -->
<style type="text/css">
	th{
		text-align:right;
		white-space:nowrap;
		padding:5px 0 0 10px;
	}

	td{
		padding:5px 0 0 5px;
	}
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->
	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">Event Data</h4>
			
			<div style="margin:10px;">
				<form name="which_event" method="post" action="event_data.asp?event_id=<%=lEventID%>">
				<span style="font-weight:bold;">Select Event:</span>
				<select name="events" id="events" onchange="this.form.get_event.click()">
					<option value="">&nbsp;</option>
					<%For i = 0 to UBound(Events, 2) - 1%>
						<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
							<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
						<%Else%>
							<option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
						<%End If%>
					<%Next%>
				</select>
				<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
				<input type="submit" name="get_event" id="get_event" value="Get This Event" style="font-size:0.8em;">
				</form>
			</div>
			
			<%If Not Clng(lEventID) = 0 Then%>	
				<div style="text-align:right;font-size:0.8em;">
					<a href="print_event_data.asp?event_id=<%=lEventID%>" onclick="openThis(this.href,1024,768);return false;">Print Data</a>
					&nbsp;|&nbsp;
					<a href="email_event_data.asp?event_id=<%=lEventID%>" onclick="openThis(this.href,1024,768);return false;">Email This</a>
				</div>
	
				<h4 style="background-color:#ececec;">Event Director Information</h4>
				
				<table style="margin:10px;">
					<tr>
						<th>First Name:</th>
						<td><%=EventDir(0)%></td>
						<th>Last Name:</th>
						<td><%=EventDir(1)%></td>
						<th>Phone:</th>
						<td><%=EventDir(2)%></td>
						<th>Mobile:</th>
						<td><%=EventDir(11)%></td>
					</tr>
					<tr>
						<th>Address:</th>
						<td><%=EventDir(3)%></td>
						<th>City:</th>
						<td><%=EventDir(4)%></td>
						<th>State:</th>
						<td><%=EventDir(5)%></td>
						<th>Zip:</th>
						<td><%=EventDir(6)%></td>
					</tr>
					<tr>
						<th valign="top">User Name:</th>
						<td valign="top"><%=EventDir(8)%></td>
						<th valign="top">Password:</th>
						<td valign="top"><%=EventDir(9)%></td>
						<th>Email:</th>
						<td colspan="3"><a href="mailto:<%=EventDir(7)%>"><%=EventDir(7)%></a></td>
					</tr>
					<tr>
						<th valign="top">Comments:</th>
						<td colspan="5"><%=EventDir(10)%></td>
					</tr>
				</table>

				<h4 style="background-color:#ececec;">Event Information</h4>
				
				<table>
					<tr>	
						<th>Event Name:</th>
						<td style="white-space:nowrap;"><%=EventArray(0)%></td>
						<th>Event Type:</th>
						<td style="white-space:nowrap;">
							<%For i = 0 to UBound(EventTypes, 2) - 1%>
								<%If CLng(EventArray(1)) = CLng(EventTypes(0, i)) Then%>
									<%=EventTypes(1, i)%>
								<%End If%>
							<%Next%>
						</td>
						<th>Event Date:</th>
						<td style="white-space:nowrap;"><%=EventArray(2)%></td>
					</tr>
					<tr>	
						<th valign="top">Show Online?</th>
						<td valign="top"><%= EventArray(6)%></td>
						<th valign="top">Online Part Reg?</th>
						<td valign="top"><%=EventArray(7)%></td>
						<th valign="top">Address:</th>
						<td valign="top"><%=EventArray(8)%></td>
					</tr>
					<tr>	
						<th valign="top">Club:</th>
						<td valign="top"><%=EventArray(3)%></td>
						<th valign="top">Website:</th>
						<td valign="top"><a href="<%=EventArray(4)%>" onclick="openThis(this.href,1024,768);return false;"><%=EventArray(4)%></a></td>
						<th valign="top">End Pre-Reg:</th>
						<td style="white-space:nowrap;" valign="top"><%=EventArray(5)%></td>
					</tr>
					<tr>	
						<th valign="top">Map to Site:</th>
						<td valign="top"><a href="<%=EventArray(9)%>" onclick="openThis(this.href,1024,768);return false;"><%=EventArray(9)%></a></td>
						<th valign="top">Event Site:</th>
						<td valign="top"><%=EventArray(10)%></td>
						<th valign="top">Need Bibs?&nbsp;<span style="font-weight:normal;"><%=EventArray(13)%></span></th>
						<th style="text-align:left;" valign="top">Need Pins?&nbsp;<span style="font-weight:normal;"><%=EventArray(14)%></span></th>
					</tr>
					<tr>	
						<th valign="top">Waiver:</th>
						<td colspan="5"><%=EventArray(11)%></td>
					</tr>
					<tr>
						<th valign="top">Comments:</th>
						<td colspan="5"><%=EventArray(12)%></td>
					</tr>
				</table>
				
				<h4 style="background-color:#ececec;margin-top:10px;">Race Information</h4>
				
				<%For i = 0 To UBound(Races, 2) - 1%>
					<%Call MoreRaceData(Races(9, i))%>
					
					<h4 style="margin:10px 0 0 10px;"><%=Races(10, i)%></h4>
					
					<table>
						<tr>
							<th>Distance:</th>
							<td><%=Races(0, i)%></td>
							<th style="white-space:nowrap;">Race Type:</th>
							<td><%=GetThisType(Races(1, i))%></td>
							<th style="white-space:nowrap;">Start Time:</th>
							<td><%=Races(4, i)%></td>
							<th>Certified?</th>
							<td><%=Races(5, i)%></td>
						</tr>
						<tr>
							<th>T-Shirts?</th>
							<td><%=TShirts(0)%></td>
							<th>Sleeve Length:</th>
							<td>
								<%If TShirts(6) = "y" Then%>
									Short
								<%End If%>
								<%If TShirts(6) = "y" And TShirts(7) = "y" Then%>
									&nbsp;&&nbsp;
								<%End If%>
								<%If TShirts(7) = "y" Then%>
									Long
								<%End If%>
							</td>
							<th>Sizes:</th>
							<td>
								<%If TShirts(1) = "y" Then%>
									S, 
								<%End If%>
								<%If TShirts(2) = "y" Then%>
									M, 
								<%End If%>
								<%If TShirts(3) = "y" Then%>
									L, 
								<%End If%>
								<%If TShirts(4) = "y" Then%>
									XL, 
								<%End If%>
								<%If TShirts(5) = "y" Then%>
									XXL
								<%End If%>
							</td>
							<th>Choose None?</th>
							<td><%=TShirts(8)%></td>
						</tr>
						<tr>
							<th style="white-space:nowrap;">Pre-Reg Fee:</th>
							<td>$<%=Races(2, i)%></td>
							<th style="white-space:nowrap;">Race Day Fee:</th>
							<td>$<%=Races(3, i)%></td>
							<th style="white-space:nowrap;">Start Type:</th>
							<td><%=Races(6, i)%></td>
							<th>Open Awards:</th>
							<td>M:&nbsp;<%=Races(7, i)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;F:&nbsp;<%=Races(8, i)%></td>
						</tr>
						<tr>
							<th>Bib Range:</th>
							<td colspan="3">From <%=iBibsFrom%> to <%=iBibsTo%></td>
							<th>Reg Link:</th>
							<td colspan="3"><%=Races(11, i)%></td>
						</tr>
						<tr>
							<th style="text-align:left;" colspan="6">Mens Age Groups (Awards):</th>
						</tr>
						<tr>
							<td style="text-align:left;padding-left:10px;" colspan="6">
								<%Call GetMAgeGrps(Races(9,i))%>
								<%For j = 0 to UBound(MAgeGrps) - 1%>
									<%=MAgeGrps(j)%>
								<%Next	%>
							</td>
						</tr>
						<tr>
							<th style="text-align:left;" colspan="6">Womens Age Groups (Awards):</th>
						</tr>
						<tr>
							<td style="text-align:left;padding-left:10px;" colspan="6">
								<%Call GetFAgeGrps(Races(9,i))%>
								<%For j = 0 to UBound(FAgeGrps) - 1%>
									<%=FAgeGrps(j)%>
								<%Next	%>
							</td>
						</tr>
					</table>
				<%Next%>
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