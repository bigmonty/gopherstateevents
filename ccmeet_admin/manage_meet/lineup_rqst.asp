<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisMeet, lTeamID
Dim i, j
Dim sMeetName, sMsg, sUrgent
Dim dMeetDate, dShutdown
Dim MeetTeams(), RosterTeams(), RosterReady(), SendTo()
Dim cdoMessage, cdoConfig
Dim bRosterExists

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate, WhenShutdown FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
dShutdown = rs(2).Value
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
    sUrgent = "n"
    If Request.Form.Item("urgent") = "on" Then sUrgent = "y"

	i = 0
	ReDim SendTo(5, 0)
	For j = 0 To UBound(MeetTeams, 2) - 1
        sql = "SELECT t.TeamsID, t.TeamName, c.LastName, c.Email, c.UserID, c.Password FROM Teams t INNER JOIN Coaches c "
	    sql = sql & "ON t.CoachesID = c.CoachesID  WHERE t.TeamsID = " & MeetTeams(0, j)
	    Set rs = conn.Execute(sql)
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
	    Set rs = Nothing
    Next

	For i = 0 to UBound(SendTo, 2) - 1
		If Not SendTo(3, i) = vbNullString Then
			sMsg = vbCrLf
			sMsg = sMsg & "Dear Coach " & SendTo(2, i) & ": " & vbCrLf & vbCrLf
	
            If sUrgent = "n" Then
			    sMsg = sMsg & "You are receiving this email because " & SendTo(1, i) & " is participating in the "
			    sMsg = sMsg & sMeetName & " on " & dMeetDate & ".   Please click the link below to submit your line-up, using "
			    sMsg = sMsg & "the login information below. Please ensure that this is done prior to " & dShutdown & ". " & vbCrLf & vbCrLf
			Else
			    sMsg = sMsg & "URGENT:  You are receiving this note because " & SendTo(1, i) & " is participating in the "
			    sMsg = sMsg & sMeetName & " on " & dMeetDate & " and does not yet have a line-up in.   NOTE:  A LINE-UP is different than a ROSTER.  "
                sMsg = sMsg & "Your ROSTER is all the kids on your team.  Your LINE-UP is who is running which race in this meet.  "
                sMsg = sMsg & "Please click the link below to submit your line-up, using "
			    sMsg = sMsg & "the login information below. THIS MUST BE DONE PRIOR TO " & dShutdown & ". " & vbCrLf & vbCrLf

                sMsg = sMsg & "You may make a limited number of changes to your line-up on site but NO ROSTER ADDITIONS CAN BE MADE THEN!  Please "
                sMsg = sMsg & "ensure that anyone who MIGHT run in this meet is on your roster by the above-indicated deadline. " & vbCrLf & vbCrLf
            End If
            				
			sMsg = sMsg & "https://www.gopherstateevents.com/misc/login.asp" & vbCrLf & vbCrLf
						
			sMsg = sMsg & "User ID: " & SendTo(4, i) & " " & vbCrLf
			sMsg = sMsg & "Password: " & SendTo(5, i) & " " & vbCrLf & vbCrLf
			
			sMsg = sMsg & "Thank you in advance for taking care of this in a timely fashion.  You may call or email me for assistance "
			sMsg =sMsg & "regarding this process. " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
			sMsg = sMSg & "Bob Schneider " & vbCrLf
			sMsg = sMSg & "GSE (Gopher State Events) " & " " & vbCrLf 
			sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
			sMsg = sMsg & "612-720-8427 " & vbCrLf
	
			'write these to db
			sql = "INSERT INTO LineupRqst (TeamsID, DateSent, MeetID) VALUES (" & CLng(SendTo(0, i)) & ", '" & Date 
			sql = sql & "', " & lThisMeet & ")"
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
				.Subject = "Meet Line-up Request: " & SendTo(1, i)
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
			Set cdoConfig = Nothing
		End If
	Next
End If

'identify which teams have meet rosters uploaded
i = 0
ReDim RosterTeams(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DISTINCT r.TeamsID FROM Roster r INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
sql = sql & "WHERE ir.MeetsID = " & lThisMeet & " ORDER BY r.TeamsID"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	RosterTeams(i) = rs(0).Value
	i = i + 1
	ReDim Preserve RosterTeams(i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Function DatesSent(lTeamID)
	DatesSent = vbNullString
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT DateSent FROM LineupRqst WHERE TeamsID = " & lTeamID & " AND MeetID = " & lThisMeet 
	sql = sql & " ORDER BY DateSent"
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
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>
			
			<h4 class="h4">CCMeet Line-Up Request: <%=sMeetName%></h4>
			
			<div class="row">
				<div class="col-sm-6">
					<form class="form" name="request_lineup" method="post" action="lineup_rqst.asp?meet_id=<%=lThisMeet%>">
					<table class="table table-striped">
						<tr>
							<td colspan="4">	
								<input type="hidden" name="request_from" id="request_from" value="request_from">
								<input type="submit" class="form-control" name="submit" id="submit" value="Request Lineup(s)">
							</td>
						</tr>
						<tr>
							<td colspan="4">	
								<input type="checkbox" name="send_all" id="send_all">&nbsp;Send All
								<input type="checkbox" name="urgent" id="urgent">&nbsp;Send as Urgent
							</td>
						</tr>
						<tr>
							<th>Exists</th>
							<th>Team</th>
							<th>Date Sent</th>
							<th>Send</th>
						</tr>
						<%For i = 0 to UBound(MeetTeams, 2) - 1%>
							<%bRosterExists = False%>
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
								<td><%=MeetTeams(1, i)%></td>
								<td><%=DatesSent(MeetTeams(0, i))%></td>
								<td>
									<input type="checkbox" name="request_<%=MeetTeams(0, i)%>" id="request_<%=MeetTeams(0, i)%>"
									<%If bRosterExists = True Then%>
										disabled
									<%End If%>
									/>
								</td>
							</tr>
						<%Next%>
					</table>
					</form>
				</div>
				<div class="col-sm-6">
					<p>Dear Coach Doe:</p>
			
					<h5>Non-urgent version:</h5>
					<p>You are receiving this email because Some Team is participating in the <%=sMeetName%> on <%=dMeetDate%>.   Please click the link
					below to submit your line-up using the login information indicated. Please ensure that this is done prior to <%=dShutdown%>.</p>
					
					<h5>Urgent Version:</h5>	
					<p>URGENT:  You are receiving this note because Some Team is participating in the
					<%=sMeetName%> on <%=dMeetDate%> and does not yet have a line-up in.   NOTE:  A LINE-UP is different than a ROSTER.
					Your ROSTER is all the kids on your team.  Your LINE-UP is who is running which race in this meet.
					Please click the link below to submit your line-up, using
					the login information below. THIS MUST BE DONE PRIOR TO <%=dShutdown%>.</p>

					<p>You may make a limited number of changes to your line-up on site but NO ROSTER ADDITIONS CAN BE MADE THEN.  Please
					ensure that anyone who MIGHT run in this meet is on your roster by the above-indicated deadline.</p>

					<p><a href="https://www.gopherstateevents.com/misc/login.asp">Login</a></p>
								
					<p>User ID: my_user_id <br>
					Password: my_password</p>
					
					<p>Thank you in advance for taking care of this in a timely fashion.  You may call or email me for assistance regarding this process.</p>
			
					<p>Sincerely;</p>
					<p>Bob Schneider<br>
					GSE (Gopher State Events)<br> 
					www.gopherstateevents.com<br>
					612-720-8427</p>
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
