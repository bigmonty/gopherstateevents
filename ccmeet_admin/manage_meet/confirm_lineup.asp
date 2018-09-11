<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisMeet, lTeamID
Dim i, j, k
Dim sMeetName, sGradeYear, sMsg, sSport, sGender
Dim MeetTeams(), RosterTeams(), RosterReady(), SendTo(), RaceArr()
Dim cdoMessage, cdoConfig
Dim dMeetDate, dShutdown
Dim bRosterExists, bHasEntries

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate, WhenShutdown, Sport FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
dShutdown = rs(2).Value
sSport  = rs(3).Value
Set rs = Nothing
 
'get year for roster grades
If Month(dMeetDate) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

'get meet teams array
i = 0
ReDim MeetTeams(2, 0)
sql = "SELECT mt.TeamsID, t.TeamName, t.Gender FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
    MeetTeams(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve MeetTeams(2, i)
	rs.MoveNext
Loop
Set rs = Nothing
    	
'get races in this meet
i = 0
ReDim RaceArr(2, 0)
sql = "SELECT RacesID, RaceDesc, Gender FROM Races WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RaceArr(0, i) = rs(0).Value
	RaceArr(1, i) = Replace(rs(1).Value, "''", "'")
    RaceArr(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve RaceArr(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("request_from") = "request_from" Then
			
%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

	i = 0
	ReDim SendTo(6, 0)
    For j = 0 to UBound(MeetTeams, 2) - 1
	    sql = "SELECT t.TeamsID, t.TeamName, c.LastName, c.Email, UserID, Password, t.Gender FROM Teams t INNER JOIN "
	    sql = sql & "Coaches c ON t.CoachesID = c.CoachesID  WHERE t.TeamsID = " & MeetTeams(0, j)
	    Set rs = conn.Execute(sql)
        If Request.Form.Item("send_all") = "on" Then
			SendTo(0, i) = rs(0).Value
			SendTo(1, i) = rs(1).Value & " (" & rs(6).Value & ")"
			SendTo(2, i) = rs(2).Value
			SendTo(3, i) = rs(3).Value
			SendTo(4, i) = rs(4).Value
			SendTo(5, i) = rs(5).Value
            SendTo(6, i) = rs(6).Value
			i = i + 1
			ReDim Preserve SendTo(6, i)
        Else
		    If Request.Form.Item("request_" & rs(0).Value) = "on" Then
			    SendTo(0, i) = rs(0).Value
			    SendTo(1, i) = rs(1).Value & " (" & rs(6).Value & ")"
			    SendTo(2, i) = rs(2).Value
			    SendTo(3, i) = rs(3).Value
			    SendTo(4, i) = rs(4).Value
			    SendTo(5, i) = rs(5).Value
                SendTo(6, i) = rs(6).Value
			    i = i + 1
			    ReDim Preserve SendTo(6, i)
		    End If
        End If
	    Set rs = Nothing
    Next

	For i = 0 to UBound(SendTo, 2) - 1
		If Not SendTo(3, i) & "" = "" Then
			sMsg = vbCrLf
			sMsg = sMsg & "Dear Coach " & SendTo(2, i) & ": " & vbCrLf & vbCrLf
	
            If CDate(Date) > CDate(dShutdown) Then
			    sMsg = sMsg & "Below is the meet line-up for the " & sMeetName & " for " & SendTo(1, i) & ".  If you want to make any changes please "
                sMsg = sMsg & "contact me at bob.schneider@gopherstateevents.com.  Subject to meet management policy, you can "
                sMsg = sMsg & "also make a limited number of changes on site, however please avoid this if possible. " & vbCrLf & vbCrLf
            Else
			    sMsg = sMsg & "Below is the meet line-up for the " & sMeetName & " for " & SendTo(1, i) & ".  You can make changes yourself via your "
                sMsg = sMsg & " Gopher State Events login until " & dShutdown & ".  If you want to make any changes after that please "
                sMsg = sMsg & "contact me at bob.schneider@gopherstateevents.com.  Subject to meet management policy, you can "
                sMsg = sMsg & "also make a limited number of changes on site, however please avoid this if possible. " & vbCrLf & vbCrLf
            End If	

			sMsg = sMsg & "Meet Information can be found at https://www.gopherstateevents.com/events/ccmeet_info.asp?meet_id=" & lThisMeet & " " & vbCrLf & vbCrLf

            sMsg = sMsg & "RESULTS will be posted online as the meet progresses.  Coaches and team staff will get an email indicating when the results "
            sMsg = sMsg & "for the each race go online, including a listing of your team's individual finishers and a direct link to the results page "
            sMsg = sMsg & "online.  PLEASE CHECK THESE RESULTS EMAILS FOR ACCURACY!  Hard copies will NOT be printed!" & vbCrLf & vbCrLf
            
            If sSport = "Nordic Ski" Then 
                sMsg = sMsg & "FINISH LINE PIX will be online later that evening and can be found on the results page once they are processed.  Notification "
                sMsg = sMsg & "will be sent to the team coach and team staff when they are available.  " & vbCrLf & vbCrLf
            End If

			sMsg = sMsg & "NOTE:  You can print out a meet sheet to record your team's individual and team results during the meet. This can be done "
			sMsg = sMsg & "from your login at https://www.gopherstateevents.com/." & vbCrLf & vbCrLf
				
			sMsg = sMsg & "FINALLY:  Please return the envelope with any unused bib numbers to the timing tent PRIOR TO THE START OF THE MEET.  If someone is "
			sMsg = sMsg & "not  participating we want to ensure that their bib number does not get accidentally read by our RFID readers, thus skewing the meet results." & vbCrLf & vbCrLf
				
			sMsg = sMsg & "Please print this email and bring it with you to the meet.  There will NOT be an additional copy in your packet. " & vbCrLf & vbCrLf

			sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
			sMsg = sMSg & "Bob Schneider " & vbCrLf
			sMsg = sMSg & "GSE (Gopher State Events) " & " " & vbCrLf 
			sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
			sMsg = sMsg & "612-720-8427 " & vbCrLf & vbCrLf

            Select Case SendTo(6, i)
                Case "M"
                    sGender = "Male"
                Case "F"
                    sGender = "Female"
            End Select
			
			'get lineup for each race in this meet for this team
			For j = 0 to UBound(RaceArr, 2) - 1
                'only include the correct gender in each email
                If RaceArr(2, j) = sGender Then
                    bHasEntries = False
				    Set rs = Server.CreateObject("ADODB.Recordset")
				    sql = "SELECT r.LastName, r.FirstName, g.Grade" & sGradeYear & ", ir.Bib FROM Roster r "
				    sql = sql & "INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID "
				    sql = sql & "WHERE ir.RacesID = " & RaceArr(0, j) & " AND r.TeamsID = " & SendTo(0, i) & " ORDER BY r.LastName, r.FirstName"
				    rs.Open sql, conn, 1, 2
					
                    sMsg = sMsg & RaceArr(1, j) & ": " & vbCrLf

				    If rs.RecordCount > 0 Then
					    k = 1
					    Do While Not rs.EOF
						    sMsg = sMsg & k & ") " & rs(3).Value & " " & rs(0).Value & ", " & rs(1).Value & " (" & rs(2).Value & ") " & vbCrLf
						    k = k + 1
						    rs.MoveNext
					    Loop
					    sMsg = sMsg & " " & vbCrLf

                        bHasEntries = True
				    End If
				    rs.Close
				    Set rs = Nothing 

                    If bHasEntries = False Then
                        sMsg = sMsg & "You have entered no one in this race. " & vbCrLf
                        sMsg = sMsg & " " & vbCrLf
                    End If
                End If
			Next
			 
			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = SendTo(3, i)
'				.To = "bob.schneider@gopherstateevents.com"
				If i = 0 Then .BCC = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
				.Subject = "Line-up Confirmation: " & sMeetName
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
        End If
	Next

    Set cdoConfig = Nothing
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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Line-Up Confirmation</title>
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
			
			<h4 class="h4">CCMeet Line-Up Confirmation: <%=sMeetName%></h4>
			
			<div class="row">
				<div class="col-sm-4">
					<form class="form" name="confirm_lineup" method="post" action="confirm_lineup.asp?meet_id=<%=lThisMeet%>">
					<table class="table table-striped">
						<tr>
							<td colspan="3">	
								<input type="hidden" name="request_from" id="request_from" value="request_from">
								<input type="submit" class="form-control" name="submit" id="submit" value="Send Confirmation(s)">
							</td>
						</tr>
						<tr>
							<td colspan="3"><input type="checkbox" name="send_all" id="send_all">&nbsp;Send All</td>
						</tr>
						<tr>
							<th>Exists</th>
							<th>Team</th>
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
								<td><input type="checkbox" name="request_<%=MeetTeams(0, i)%>" id="request_<%=MeetTeams(0, i)%>"></td>
							</tr>
						<%Next%>
					</table>
					</form>
				</div>
				<div class="col-sm-8">
					<p>Dear Coach Jones:</p>
			
					<%If CDate(Date) > CDate(dShutdown) Then%>
						<p>Below is the meet line-up for the <%=sMeetName%> for Some School.  If you want to make any changes please
						contact me at bob.schneider@gopherstateevents.com.  Subject to meet management policy, you can
						also make a limited number of on site, however please avoid this if possible.</p>
					<%Else%>
						<p>Below is the meet line-up for <%=sMeetName%>.  You can make changes yourself via your
						Gopher State Events login until <%=dShutdown%>.  If you want to make any changes after that please
						contact me at bob.schneider@gopherstateevents.com.  Subject to meet management policy, you can
						also make a limited number of changes on site, however please avoid this if possible.</p>
					<%End If%>

					<p>Meet Information can be found <a href="https://www.gopherstateevents.com/events/ccmeet_info.asp?meet_id=<%=lThisMeet%>">
					here</a>.</p>
					
					<p>Results will be posted online as the meet progresses.  Coaches and team followers will get an email indicating when the results
					for the each race go online, including a listing of your team's individual finishers and a direct link to the results page 
					online.  PLEASE CHECK THESE RESULTS EMAILS FOR ACCURACY!  Hard copies will NOT be printed.</p>
				
					<%If sSport = "Nordic Ski" Then%>
						<p>FINISH LINE PIX will be online later that evening and can be found on the results page once they are processed.  Notification 
						will be sent to the team coach and team followsers when they are available.</p>
					<%End If%>		
								
					<p>NOTE:  You can print out a meet sheet to record your team's individual and team results during the meet. This can be done from your 
					login at http://www.gopherstateevents.com/.</p>
					
					<p>FINALLY:  Please return the envelope with any unused bib numbers to the timing tent PRIOR TO THE START OF THE MEET.  If someone is
					not participating we want to ensure that their bib number does not get accidentally read by our RFID readers, thus skewing the meet results.</<p>
					
					<p>Please print this email and bring it with you to the meet.  There will NOT be an additional copy in your packet.</p>
		
					<p>
						Sincerely;<br>
						Bob Schneider<br>
						GSE (Gopher State Events)<br>
						www.gopherstateevents.com<br>
						612-720-8427
					</p>
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
