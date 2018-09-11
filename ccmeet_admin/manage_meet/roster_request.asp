<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisMeet, sMeetName, dMeetDate
Dim MeetTeams(), RosterTeams(), RosterReady()
Dim i, j
Dim cdoMessage, cdoConfig
Dim sMsg
Dim bRosterExists, bRosterReady
Dim SendTo()
Dim lTeamID

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

'get meet teams array
i = 0
ReDim MeetTeams(1, 0)
sql = "SELECT mt.TeamsID, t.TeamName, t.Gender FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve MeetTeams(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("request_from") = "request_from" Then
	i = 0
	ReDim SendTo(5, 0)
	For j = 0 To UBound(MeetTeams, 2) - 1
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT t.TeamsID, t.TeamName, c.LastName, c.Email, c.UserID, c.Password FROM Teams t INNER JOIN Coaches c "
		sql = sql & "ON t.CoachesID = c.CoachesID  WHERE t.TeamsID = " & MeetTeams(0, j)
		rs.Open sql, conn, 1, 2
		Do While Not rs.EOF
			If Request.Form.Item("send_all") = "on" Then
				SendTo(0, i) = rs(0).Value
				SendTo(1, i) = rs(1).Value
				SendTo(2, i) = rs(2).Value
				SendTo(3, i) = rs(3).Value
				SendTo(4, i) = rs(4).Value
				SendTo(5, i) = rs(5).Value
				i = i + 1
				ReDim Preserve SendTo(5, i)
			Else
				If Request.Form.Item("request_" & rs(0).Value) = "on" Then
					SendTo(0, i) = rs(0).Value
					SendTo(1, i) = rs(1).Value
					SendTo(2, i) = rs(2).Value
					SendTo(3, i) = rs(3).Value
					SendTo(4, i) = rs(4).Value
					SendTo(5, i) = rs(5).Value
					i = i + 1
					ReDim Preserve SendTo(5, i)
				End If
			End If
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
	Next
	For i = 0 to UBound(SendTo, 2) - 1
		If Not SendTo(3, i) = vbNullString Then
			sMsg = vbCrLf
			sMsg = sMsg & "Dear Coach " & SendTo(2, i) & ": " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "You are receiving this email because your team does not yet have a roster on file with GSE "
			sMsg = sMsg & "and you are scheduled to compete in the " & sMeetName & " on " & dMeetDate & ".  It is "
			sMsg = sMsg & "imperative that we have a roster on file for your team at least 72 hours prior to this event.  As "
			sMsg = sMsg & "well, at least 24 hours prior to this event you will need to submit a meet line-up indicating which "
			sMsg = sMsg & "athletes you expect to be running. " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "There are 2 ways you can submit your roster to us: " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "1) Go to www.gopherstateevents.com, log in, and hand-enter your roster one-at-a-time.  Again, more specific "
			sMSg = sMsg & "instructions for this can be found on the site. " & vbCrLf & vbCrLf
			
			sMsg = sMsg & "User ID: " & SendTo(4, i) & " " & vbCrLf
			sMsg = sMsg & "Password: " & SendTo(5, i) & " " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "2) Attach the roster to an email addressed to bob.schneider@gopherstateevents.com.  Make sure that the "
			sMsg = sMsg & "attachment is an MS Excel file and contains first name, last name, gender, and grade. " & vbCrLf & vbCrLf
			
			sMsg = sMsg & "Thank you in advance for taking care of this in a timely fashion to ensure that all is ready to go "
			sMsg = sMsg &  "on race day.  You may call or email me for assistance regarding this process. " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
			sMsg = sMSg & "Bob Schneider " & vbCrLf
			sMsg = sMSg & "GSE (Gopher State Events) www.gopherstateevents.com " & vbCrLf
			sMsg = sMsg & "612-720-8427 " & vbCrLf
	
			'write these to db
			sql = "INSERT INTO RosterRequest (TeamsID, DateSent) VALUES (" & CLng(SendTo(0, i)) & ", '" & Date & "')"
			Set rs = conn.Execute(sql)
			Set rs = Nothing

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = SendTo(3, i)
'				.BCC = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
				.Subject = "GSE CC Roster Request"
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
			Set cdoConfig = Nothing
		End If
	Next
End If

'identify which teams have rosters uploaded
i = 0
ReDim RosterTeams(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID FROM RosterUpload ORDER BY TeamsID"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	RosterTeams(i) = rs(0).Value
	i = i + 1
	ReDim Preserve RosterTeams(i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'identify which teams have rosters ready
i = 0
ReDim RosterReady(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DISTINCT TeamsID FROM Roster ORDER BY TeamsID"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	RosterReady(i) = rs(0).Value
	i = i + 1
	ReDim Preserve RosterReady(i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Function DatesSent(lTeamID)
	DatesSent = vbNullString
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT DateSent FROM RosterRequest WHERE TeamsID = " & lTeamID & " ORDER BY DateSent"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		Do While Not rs.EOF
			If DatesSent = vbNullString Then
				If Year(rs(0).Value) = Year(Date) Then
					DatesSent = rs(0).Value
				End If
			Else
				If Year(rs(0).Value) = Year(Date) Then
					DatesSent = DatesSent & "<br>" & rs(0).Value
				End If
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Roster Request</title>

</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		
		<div class="col-sm-10">
			<h4 class="h4">CCMeet Roster Request: <%=sMeetName%></h4>

			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>

			<div class="row">
				<div class="col-sm-6">			
					<form class="form" name="request_roster" method="post" action="roster_request.asp?meet_id=<%=lThisMeet%>">
					<table class="table table-striped">
						<tr>
							<td colspan="5">	
								<input type="hidden" name="request_from" id="request_from" value="request_from">
								<input type="submit" class="form-control" name="submit" id="submit" value="Request Roster(s)">
							</td>
						</tr>
						<tr>
							<td colspan="5">	
								<input type="checkbox" name="send_all" id="send_all">&nbsp;Send All
							</td>
						</tr>
						<tr>
							<th>Exists</th>
							<th>Ready</th>
							<th>Team</th>
							<th>Date Sent</th>
							<th>Send</th>
						</tr>
						<%For i = 0 to UBound(MeetTeams, 2) - 1%>
							<%bRosterExists = False%>
							<%bRosterReady = False%>
							<tr>
								<td>
									<%For j = 0 to UBound(RosterTeams) - 1%>
										<%If CLng(RosterTeams(j)) = CLng(MeetTeams(0, i)) Then%>
											<%bRosterExists = True%>
											<input type="radio" name="exists_<%=MeetTeams(0, i)%>" id="exists_<%=MeetTeams(0, i)%>" 
														checked>
											<%Exit For%>
										<%Else%>
											<%If j = UBound(RosterTeams) - 1 Then%>
												<input type="radio" name="exists_<%=MeetTeams(0, i)%>" id="exists_<%=MeetTeams(0, i)%>">
											<%End If%>
										<%End If%>
									<%Next%>
								</td>
								<td>
									<%For j = 0 to UBound(RosterReady) - 1%>
										<%If CLng(RosterReady(j)) = CLng(MeetTeams(0, i)) Then%>
											<%bRosterReady = True%>
											<input type="radio" name="ready_<%=MeetTeams(0, i)%>" id="ready_<%=MeetTeams(0, i)%>" 
														checked>
											<%Exit For%>
										<%Else%>
											<%If j = UBound(RosterReady) - 1 Then%>
												<input type="radio" name="ready_<%=MeetTeams(0, i)%>" id="ready_<%=MeetTeams(0, i)%>">
											<%End If%>
										<%End If%>
									<%Next%>
								</td>
								<td><%=MeetTeams(1, i)%></td>
								<td><%=DatesSent(MeetTeams(0, i))%></td>
								<td><input type="checkbox" name="request_<%=MeetTeams(0, i)%>" id="request_<%=MeetTeams(0, i)%>"></td>
							</tr>
						<%Next%>
					</table>
					</form>
				</div>
				<div class="col-sm-6">
					<p>Dear Coach Doe:</p>
			
					<p>You are receiving this email because your team does not yet have a roster on file with GSE
					and you are scheduled to compete in the " & sMeetName & " on " & dMeetDate & ".  It is
					imperative that we have a roster on file for your team at least 72 hours prior to this event.  As
					well, at least 24 hours prior to this event you will need to submit a meet line-up indicating which
					athletes you expect to be running.</p>
			
					<p>There are 2 ways you can submit your roster to us: </p>
			
					<ol>
						<li>
							Go to www.gopherstateevents.com, log in, and hand-enter your roster one-at-a-time.  Again, more specific "
							instructions for this can be found on the site.<br>
					
							User ID: user_id<br>
							Password: pwd<br>
						</li>
						<li>
							Attach the roster to an email addressed to bob.schneider@gopherstateevents.com.  Make sure that the "
							attachment is an MS Excel file and contains first name, last name, gender, and grade.
						</li>
					</ol>

					<p>Thank you in advance for taking care of this in a timely fashion to ensure that all is ready to go
					on race day.  You may call or email me for assistance regarding this process.</p>
			
					<p>Sincerely;<br><br>
					Bob Schneider<br>
					GSE (Gopher State Events) www.gopherstateevents.com<br>
					612-720-8427 </p>
				</div>
			</div>
		</div>
	</div>
	<!--#include file = "../../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
