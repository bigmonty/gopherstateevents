<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet, lTeamID
Dim sMsg, sMeetName
Dim MeetTeams(), AlreadySent(), SendTo()
Dim cdoMessage, cdoConfig
Dim dMeetDate
Dim bAlreadySent

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

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

If Request.Form.Item("send_instr") = "send_instr" Then
	i = 0
	ReDim SendTo(5, 0)
	sql = "SELECT t.TeamsID, t.TeamName, c.LastName, c.Email, c.UserID, c.Password FROM Teams t INNER JOIN Coaches c "
	sql = sql & "ON t.CoachesID = c.CoachesID ORDER BY t.TeamsID"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
        If InMeet(rs(0).Value) = "y" Then
            If Request.Form.Item("send_all") = "on" Then
                If Request.Form.Item("no_resend") = "on" Then
                    If CDate(LastSend(rs(0).Value)) <= CDate("8/1/" & Year(Date)) Then 'hasnt been sent this year so send
			            SendTo(0, i) = rs(0).Value
			            SendTo(1, i) = rs(1).Value
			            SendTo(2, i) = rs(2).Value
			            SendTo(3, i) = rs(3).Value
			            SendTo(4, i) = rs(4).Value
			            SendTo(5, i) = rs(5).Value
			            i = i + 1
			            ReDim Preserve SendTo(5, i)
                    End If
                Else
			        SendTo(0, i) = rs(0).Value
			        SendTo(1, i) = rs(1).Value
			        SendTo(2, i) = rs(2).Value
			        SendTo(3, i) = rs(3).Value
			        SendTo(4, i) = rs(4).Value
			        SendTo(5, i) = rs(5).Value
			        i = i + 1
			        ReDim Preserve SendTo(5, i)
                End If
		    Else
                If Request.Form.Item("send_to_" & rs(0).Value) = "on" Then
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
        End If
		rs.MoveNext
	Loop
	Set rs = Nothing

	For i = 0 to UBound(SendTo, 2) - 1
		If Not SendTo(3, i) & "" = "" Then
			sMsg = vbCrLf
			sMsg = sMsg & "Dear Coach " & SendTo(2, i) & ": " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "You are receiving this email because your team is scheduled to participate in the " & sMeetName 
			sMsg = sMsg & " on " & dMeetDate & " which is being timed and scored by GSE (Gopher State Events)." & vbCrLf & vbCrLf
	
			sMsg = sMsg & "An account has been created for you at www.gopherstateevents.com.  Your login "
			sMsg = sMsg & "information for this account is listed below.  This account allows you access to meetinfo, "
			sMsg = sMsg & "course maps (if available to us), a map to the site, the ability "
			sMsg = sMsg & "to upload your roster to us, the ability to assign kids to races (your meet line-up), pre-filled meet sheets, and more. " 
            sMsg = sMsg & "YOU WILL ALSO FIND INFORMATION ON HOW TO HAVE US UPLOAD YOUR ROSTER FOR YOU! " & vbCrLf & vbCrLf
			
			sMsg = sMsg & "User ID: " & SendTo(4, i) & " " & vbCrLf
			sMsg = sMsg & "Password: " & SendTo(5, i) & " " & vbCrLf & vbCrLf
			
			sMsg = sMsg & "AN IMPORTANT NOTE ABOUT SUBMITTING YOUR ROSTER:  If you have a roster on file from a previous year, "
			sMsg = sMsg & "you will NOT need to submit a new one.  Your roster still exists AND YOUR PARTICIPANT'S GRADES "
			sMsg = sMsg & "HAVE BEEN INCREMENTED BY ONE.  You will simply need to delete those participants who are no "
			sMsg = sMsg & "longer on your team, add any new members, and adjust any grades that should not have been "
			sMsg = sMsg & "increased.  " & vbCrLf & vbCrLf
			
			sMsg = sMsg & "Thank you for in advance.  You may call or email me for assistance regarding this process.  " & vbCrLf & vbCrLf
	
			sMsg = sMsg & "Sincerely;" & vbCrLf & vbCrLf
			sMsg = sMsg & "Bob Schneider " & vbCrLf
			sMsg = sMsg & "GSE (Gopher State Events) " & vbCrLf
			sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
			sMsg = sMsg & "612-720-8427 " & vbCrLf
	
			'write these to db
			sql = "INSERT INTO TeamInstr (TeamsID, DateSent) VALUES (" & CLng(SendTo(0, i)) & ", '" & Date & "')"
			Set rs = conn.Execute(sql)
			Set rs = Nothing

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
			 
			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = SendTo(3, i)
                If i = 0 Then .BCC = "bob.schneider@gopherstateevents.com"
'				.BCC = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
				.Subject = sMeetName
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
			Set cdoConfig = Nothing
		End If
	Next
End If

'identify which teams have had this information sent to them this year
i = 0
ReDim AlreadySent(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, DateSent FROM TeamInstr ORDER BY TeamsID"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If Year(rs(1).Value) = Year(Date) Then
		AlreadySent(0, i) = rs(0).Value
		AlreadySent(1, i) = rs(1).Value
		i = i + 1
		ReDim Preserve AlreadySent(1, i)
	End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function LastSend(lThisTeam)
    Dim rs2, sql2

    LastSend = "8/1/" & Year(Date)      'set date to rule out sends from previous years

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT DateSent FROM TeamInstr WHERE TeamsID = " & lThisTeam & " AND DateSent > '8/1/" & Year(Date) & "' ORDER BY DateSent DESC"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then LastSend = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function InMeet(lThisTeam)
    Dim x

    InMeet = "n"

    For x = 0 To UBound(MeetTeams, 2) - 1
        If CLng(lThisTeam) = CLng(meetTeams(0, x)) Then
            InMeet = "y"
            Exit For
        End If
    Next
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country Team Meet Instructions Sheet</title>

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
			
			<h4 class="h4">CCMeet Team Instructions: <%=sMeetName%></h4>
				
			<div class="row">
				<div class="col-sm-6">		
					<form class="form" name="send_instr" method="post" action="team_instr.asp?meet_id=<%=lThisMeet%>">
					<table class="table table-striped">
						<tr>
							<td colspan="4">	
								<input type="hidden" name="send_instr" id="send_instr" value="send_instr">
								<input class="form-control" type="submit" name="submit" id="submit" value="Send Instructions">
							</td>
						</tr>
						<tr>
							<td colspan="4">	
								<input type="checkbox" name="send_all" id="send_all">&nbsp;Send All
								<input type="checkbox" name="no_resend" id="no_resend" checked>&nbsp;No Resend
							</td>
						</tr>
						<tr>
							<td>No.</td>
							<td>Team</td>
							<td>Date Sent</td>
							<td>Send</td>
						</tr>
						<%For i = 0 to UBound(MeetTeams, 2) - 1%>
							<%bAlreadySent = False%>
								<tr>
									<td><%=i + 1%>)</td>
									<td><%=MeetTeams(1, i)%></td>
									<td>
										<%For j = 0 to UBound(AlreadySent, 2) - 1%>
											<%If CLng(AlreadySent(0, j)) = CLng(MeetTeams(0, i)) Then%>
												<%=AlreadySent(1, j)%><br>
												<%bAlreadySent = True%>
											<%End If%>
										<%Next%>
									</td>
									<td><input type="checkbox" name="send_to_<%=MeetTeams(0, i)%>" id="send_to_<%=MeetTeams(0, i)%>"></td>
								</tr>
						<%Next%>
					</table>
					</form>
				</div>
				<div class="col-sm-6">
					<h4 class="h4">Sample Text</h4>
					
					<p>Dear Coach Jones:</p>

					<p>You are receiving this email because your team is scheduled to participate in the <%=sMeetName%> 
					on <%=dMeetDate%> which is being timed and scored by GSE (Gopher State Events).  Attached
					to this email you will find the guidelines that we use to manage meet.</p>
		
					<p>An account has been created for you at www.gopherstateevents.com.  Your login
					information for this account is listed below.  This account allows you access to meetinfo,
					course maps (if available to us), a map to the site, the ability
					to upload your roster to us, the ability to assign kids to races (your meet line-up), pre-filled meet sheets, and more.
					YOU WILL ALSO FIND INFORMATION ON HOW TO HAVE US UPLOAD YOUR ROSTER FOR YOU!</p>

					<p>User ID: My ID<br>
					Password: My Pwd</p>
				
					<p>AN IMPORTANT NOTE ABOUT SUBMITTING YOUR ROSTER:  If you have a roster on file from a previous year,
					you will NOT need to submit a new one.  Your roster still exists AND YOUR PARTICIPANT'S GRADES
					HAVE BEEN INCREMENTED BY ONE.  You will simply need to delete those participants who are no
					longer on your team, add any new members, and adjust any grades that should not have been
					increased.</p>
				
					<p>Thank you for in advance.  You may call or email me for assistance regarding this process.</p>
			
					<p>Sincerely;</p>
					
					<p>Bob Schneider <br>
					GSE (Gopher State Events)<br>
					www.gopherstateevents.com<br>
					612-720-8427<br>
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
