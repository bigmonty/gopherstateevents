<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, rs, conn, rs2, sql2
Dim i, j
Dim lEventID, lEventDirID
Dim sMyEmail, sSuppMsg, sEventName
Dim RaceArr(), FinishersArr(), PartInfo()
Dim cdoMessage, cdoConfig, objEmail, xmlhttp, EmailContents, sPageToSend, sEventDirEmail
Dim dWhenSent
Dim bFound
Dim bSendThis

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDirID FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
lEventDirID = rs(1).Value
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Email FROM EventDir WHERE EventDirID = " & lEventDirID
rs.Open sql, conn, 1, 2
sEventDirEmail = rs(0).Value
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_ind") = "submit_ind" Then
	i = 0
	ReDim RaceArr(0)
	sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		RaceArr(i) = rs(0).Value
		i = i + 1
		ReDim Preserve RaceArr(i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	For i = 0 to UBound(RaceArr) - 1
		j = 0
		ReDim FinishersArr(0)
		sql = "SELECT ParticipantID FROM IndResults WHERE RaceID = " & RaceArr(i)
		Set rs = Conn.Execute(sql)
		Do While Not rs.EOF
			If Request.Form.Item("send_" & rs(0).Value) = "on" Then
				FinishersArr(j) = rs(0).Value
				j = j + 1
				ReDim Preserve FinishersArr(j)
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
		
		For j = 0 to UBound(FinishersArr) - 1
			'get email address
			sMyEmail = vbNullString
			
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT Email FROM Participant WHERE ParticipantID = " & FinishersArr(j)
			rs.Open sql, conn, 1, 2
			If Not rs(0).Value & "" = "" Then sMyEmail = rs(0).Value
			rs.Close
			Set rs = Nothing
			
			If Not sMyEmail = vbNullString Then
				If ValidEmail(sMyEmail) = True Then
                    sPageToSend = "http://www.gopherstateevents.com/race_vids/view_vids.asp?event_id=" & lEventID & "&part_id=" & FinishersArr(j)

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

                    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	                    xmlhttp.open "GET", sPageToSend, false
	                    xmlhttp.send ""
	                    EmailContents = xmlhttp.responseText
                    Set xmlhttp = nothing

			        Set cdoMessage = CreateObject("CDO.Message")
			        With cdoMessage
				        Set .Configuration = cdoConfig
				        .To = sMyEmail
				        .From = "bob.schneider@gopherstateevents.com"
   				        .Subject = "Videos Online For " & sEventName
				        .HTMLBody = EmailContents
				        .Send
			        End With
			        Set cdoMessage = Nothing
                    Set cdoConfig = Nothing

					'insert into email sent
					sql = "INSERT INTO VideosSent (ParticipantID, EventID, WhenSent) VALUES (" & FinishersArr(j) & ", "
					sql = sql & lEventID & ", '" & Now() & "')"
					Set rs = conn.Execute(sql)
					Set rs = Nothing
				End If
			End If
		Next
	Next
ElseIf Request.Form.Item("submit_all") = "submit_all" Then
	i = 0
	ReDim RaceArr(0)
	sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		RaceArr(i) = rs(0).Value
		i = i + 1
		ReDim Preserve RaceArr(i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	For i = 0 to UBound(RaceArr) - 1
		j = 0
		ReDim FinishersArr(0)
		sql = "SELECT ParticipantID FROM IndResults WHERE RaceID = " & RaceArr(i)
		Set rs = Conn.Execute(sql)
		Do While Not rs.EOF
			FinishersArr(j) = rs(0).Value
			j = j + 1
			ReDim Preserve FinishersArr(j)
			rs.MoveNext
		Loop
		Set rs = Nothing
		
		For j = 0 to UBound(FinishersArr) - 1
			bSendThis = True
			
			If Request.Form.Item("no_resend") = "on" Then
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT ParticipantID, RaceID FROM ResultsSent WHERE ParticipantID = " & FinishersArr(j) & " AND RaceID = " & RaceArr(i)
				rs.Open sql, conn, 1, 2
				If rs.RecordCount > 0 Then 	bSendThis = False
				rs.Close
				Set rs = Nothing
			End If
			
			If bSendThis = True Then
				'get email address
				sMyEmail = vbNullString
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT Email FROM Participant WHERE ParticipantID = " & FinishersArr(j)
				rs.Open sql, conn, 1, 2
				If Not rs(0).Value & "" = "" Then sMyEmail = rs(0).Value
				rs.Close
				Set rs = Nothing

				If Not sMyEmail = vbNullString Then
					If ValidEmail(sMyEmail) = True Then
                        sPageToSend = "http://www.gopherstateevents.com/race_vids/view_vids.asp?event_id=" & lEventID & "&part_id=" & FinishersArr(j)

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

                        Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	                        xmlhttp.open "GET", sPageToSend, false
	                        xmlhttp.send ""
	                        EmailContents = xmlhttp.responseText
                        Set xmlhttp = nothing

			            Set cdoMessage = CreateObject("CDO.Message")
			            With cdoMessage
				            Set .Configuration = cdoConfig
				            .To = sMyEmail
				            .From = "bob.schneider@gopherstateevents.com"
                            If j = 0 Then .BCC = "bob.schneider@gopherstateevents.com;" & sEventDirEmail
   				            .Subject = "Videos Online For " & sEventName
				            .HTMLBody = EmailContents
				            .Send
			            End With
			            Set cdoMessage = Nothing
                        Set cdoConfig = Nothing

						'insert into email sent
						sql = "INSERT INTO VideosSent (ParticipantID, EventID, WhenSent) VALUES (" & FinishersArr(j) & ", "
						sql = sql & lEventID & ", '" & Now() & "')"
						Set rs = conn.Execute(sql)
						Set rs = Nothing
					End If
				End If
			End If
		Next
	Next
	
	'add this event to the emailrslts table
	sql = "INSERT INTO VideoNotifSent(EventID, DateSent) VALUES (" & lEventID & ", '" & Now() & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DateSent FROM EmailRslts WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then dWhenSent = rs(0).Value
rs.Close
Set rs = Nothing
	
'get races in this event
i = 0
ReDim RaceArr(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	RaceArr(0, i) = rs(0).Value
	RaceArr(1, i) = Replace(rs(0).Value, "''", "'")
	i = i + 1
	ReDim Preserve RaceArr(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

%>
<!--#include file = "../../includes/valid_email.asp" -->
<%

Private Function GetMySend(lThisPart, lThisRaceID)
	GetMySend = vbNullString
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT WhenSent FROM VideosSent WHERE EventID = " & lEventID & " AND ParticipantID = " & lThisPart & " ORDER BY WhenSent DESC"
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetMySend = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function

Private Sub GetPartInfo(lThisRace)
	Dim x
	
	x = 0
	ReDim PartInfo(4, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, p.Gender, p.Email FROM Participant p INNER JOIN IndResults ir "
	sql = sql & "ON p.ParticipantID = ir.ParticipantID INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID WHERE ir.RaceID = " & lThisRace
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		PartInfo(0, x) = rs(0).Value
		PartInfo(1, x) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
		PartInfo(2, x) = rs(3).Value
		PartInfo(3, x) = GetMySend(rs(0).Value, lThisRace)
		PartInfo(4, x) = rs(4).Value
        x = x + 1
		ReDim Preserve PartInfo(4, x)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>GSE&copy; Video Notification</title>
<!--#include file = "../../includes/meta2.asp" -->




<style type="text/css">
	td, th{
		white-space:nowrap;
	}
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">Video Notification:&nbsp;<%=sEventName%></h4>
				
            <!--#include file = "../../includes/event_nav.asp" -->

			<div style="margin: 0;padding: 0;font-size: 0.85em;">
				<a href="/admin/fitness_vids.asp?event_id=<%=lEventID%>" rel="nofollow">Add/Edit Videos</a>
            </div>

			<%If CStr(dWhenSent) = vbNullString Then%>
				<p>Video notification for this event have not been sent.</p>
			<%Else%>
				<p>Video notification for this event were last sent on <%=dWhenSent%>.</p>
			<%End If%>
				
			<div style="background-color:#ececd8;">
				<form name="send_all" Method="Post" action="video_notification.asp?event_id=<%=lEventID%>">
				<input type="checkbox" name="no_resend" id="no_resend">&nbsp;No Resend
                <br>
				<input type="hidden" name="submit_all" id="submit_all" value="submit_all">
				<input type="submit" name="submit2" id="submit2" value="Send All">
				</form>
			</div>
				
			<form name="send_ind" method="Post" action="video_notification.asp?event_id=<%=lEventID%>">
			<div style="text-align:center;">
				<input type="hidden" name="submit_ind" id="submit_ind" value="submit_ind">
				<input type="submit" name="submit2" id="submit2" value="Send Selected">
			</div>
			<%For i = 0 To UBound(RaceArr, 2) - 1%>
				<%Call GetPartInfo(RaceArr(0, i))%>
					
				<h4 class="h4"><%=UBound(PartInfo, 2)%> Finishers</h4>
					
				<table>
					<tr>	
						<th>Pl.</th>
						<th>Name</th>
						<th>M/F</th>
                        <th>Email</th>
						<th>Last Send</th>
						<th>Send</th>
					</tr>
					<%For j = 0 To UBound(PartInfo, 2) - 1%>
						<%If j mod 2 = 0 Then%>
							<tr>	
								<td class="alt"><%=j +1%>)</td>
								<td class="alt"><%=PartInfo(1, j)%></td>
								<td class="alt"><%=PartInfo(2, j)%></td>
								<td class="alt"><a href="mailto:<%=PartInfo(4, j)%>"><%=PartInfo(4, j)%></a></td>
                                <td class="alt"><%=PartInfo(3, j)%></td>
								<td class="alt" style="text-align:center;">
									<input type="checkbox" name="send_<%=PartInfo(0, j)%>" id="send_<%=PartInfo(0, j)%>">
								</td>
							</tr>
						<%Else%>
							<tr>	
								<td><%=j +1%>)</td>
								<td><%=PartInfo(1, j)%></td>
								<td><%=PartInfo(2, j)%></td>
								<td><a href="mailto:<%=PartInfo(4, j)%>"><%=PartInfo(4, j)%></a></td>
                                <td><%=PartInfo(3, j)%></td>
								<td style="text-align:center;">
									<input type="checkbox" name="send_<%=PartInfo(0, j)%>" id="send_<%=PartInfo(0, j)%>">
								</td>
							</tr>
						<%End If%>
					<%Next%>
				</table>
			<%Next%>
			</form>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
