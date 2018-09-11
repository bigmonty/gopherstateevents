<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lThisMeet, lCoachesID
Dim i, j, x
Dim cdoMessage, cdoConfig
Dim sMsg, sMeetName, sSuppMsg, sSport, sCoachEmail, sOtherRecips, sAllRecips, sTeamName
Dim dMeetDate
Dim Races(), Meets(), Finishers(), SendHistory()

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")
If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

ReDim Races(1, 0)
ReDim MeetTeams(1, 0)

i = 0
ReDim Meets(2, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Meets(0, i) = rs(0).Value
	Meets(1, i) = rs(1).Value
	Meets(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve Meets(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

If CLng(lThisMeet) > 0 Then Call MeetInfo()

%>
<!--#include file = "../../../includes/cdo_connect.asp" -->
<%

If Request.Form.Item("get_meet") = "get_meet" Then
	lThisMeet = Request.Form.Item("meets")

    If CStr(lThisMeet) = vbNullString Then
        lThisMeet = 0
    Else
        Call MeetInfo()
    End If
ElseIf Request.Form.Item("submit_send") = "submit_send" Then
    sSuppMsg = Request.Form.Item("supp_msg")

    For j = 0 To UBound(Races, 2) - 1
        If Request.Form.Item("race_" & Races(0, j)) = "on" Then
	        sql = "SELECT mt.TeamsID FROM MeetTeams mt INNER JOIN Races r ON mt.MeetsID = r.MeetsID WHERE r.RacesID = " & Races(0, j)
	        Set rs = conn.Execute(sql)
	        Do While Not rs.EOF
                'see if this team is checked to receive results
		        If Request.Form.Item("team_" & rs(0).Value) = "on" Then
                    'get this team's race finishers
                    Call GetFinishers(rs(0).Value, Races(0, j))

                    'if they have some finishers proceed
                    If UBound(Finishers, 2) > 0 Then 
                        Call SendResults(rs(0).Value, Races(0, j), Races(1, j))
'                        Call SendIndResults(Races(0, j), Races(1, j))
                    End If
                End If
		        rs.MoveNext
	        Loop
	        Set rs = Nothing
        End If
    Next
End If

i = 0
ReDim SendHistory(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceName, TeamName, WhenSent, Recipients FROM ResultsSent WHERE MeetsID = " & lThisMeet & " ORDER BY RaceName, WhenSent"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    SendHistory(0, i) = rs(0).Value
    SendHistory(1, i) = Replace(rs(1).Value, "''", "'")
    SendHistory(2, i) = rs(2).Value
    SendHistory(3, i) = rs(3).Value
    i = i + 1
    ReDim Preserve SendHistory(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetFinishers(lThisTeam, lThisRace)
    x = 0
    ReDim Finishers(6, 0)  
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT ir.Bib, r.FirstName, r.LastName, ir.RaceTime, ir.RosterID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
    sql2 = sql2 & "WHERE r.TeamsID = " & lThisTeam & " AND ir.RacesID = " & lThisRace & " ORDER BY ir.Place"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        Finishers(0, x) = rs2(0).Value
        Finishers(1, x) = Replace(rs2(2).Value, "''", "'") & ", " & Replace(rs2(1).Value, "''", "'")
        Finishers(2, x) = rs2(3).Value
        Finishers(3, x) = rs2(4).Value
        'Finishers(4, x) = email
        'Finishers(5, x) = cell
        'Finishers(6, x) = provider
        x = x + 1
        ReDim Preserve Finishers(6, x)
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    For x = 0 To UBound(Finishers, 2) - 1
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT Email, CellPhone, CellProvider FROM PerfTrkr WHERE RosterID = " & Finishers(3, x)
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then
            Finishers(4, x) = rs2(0).Value
            Finishers(5, x) = rs2(1).Value
            Finishers(6, x) = rs2(2).Value
        End If
        rs2.Close
       Set rs2 = Nothing
    Next
End Sub

Private Sub SendIndResults(lThisRace, sThisRace)
    Dim x, y
    Dim sMyFollowers
    Dim MobileData()

    For x = 0 To UBound(Finishers, 2) - 1   'finishers for this team and this race
        'email results...get followers with email addresses
        sMyFollowers = vbNullString
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT f.Email FROM PTFollowers f INNER JOIN PerfTrkr p "
        sql = sql & "ON p.PerfTrkrID = f.PerfTrkrID WHERE p.RosterID = " & Finishers(3, x) & " AND f.ResultsNotif='y'"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Not rs(0).Value & "" = "" Then sMyFollowers = sMyFollowers & rs(0).Value & ";"
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If Not Finishers(4, x) & "" = "" Then   
            'email message
	        sMsg = vbNullString
            sMsg = vbCrLf
	        sMsg = sMsg & "My Results for " & sMeetName & " on " & dMeetDate & " " & vbCrLf
	        sMsg = sMsg & "Race: " & sThisRace & vbCrLf & vbCrLf

            sMsg = sMsg & "Name: " & Finishers(1, x) & " " & vbCrLf
            sMsg = sMsg & "Time: " & Finishers(2, x) & " " & vbCrLf

            sMsg = sMsg & "Note:  Finishing place is omitted in races with wave, interval, or chip start.  This is "
	        sMsg = sMsg & "because athletes who started later could change the individual finish order of the event.  "
            sMsg = sMsg & "You will find overall meet results for this race here: "
	        sMsg = sMsg & "http://www.gopherstateevents.com/results/cc_rslts/cc_rslts.asp?meet_id=" & lThisMeet & "&race_id=" & lThisRace & "&sport=" & sSport
            sMsg = sMsg  & vbCrLf & vbCrLf

	        sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
	        sMsg = sMSg & "Bob Schneider" & vbCrLf
	        sMsg = sMSg & "GSE (Gopher State Events) " & " " & vbCrLf 
	        sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
	        sMsg = sMsg & "612-720-8427 " & vbCrLf
			 
	        Set cdoMessage = CreateObject("CDO.Message")
	        With cdoMessage
		        Set .Configuration = cdoConfig
'                .To = "bob.schneider@gopherstateevents.com"
                .To = Finishers(4, x)
                If Len(sMyFollowers) > 0 Then .CC = sMyFollowers
                .BCC = "bob.schneider@gopherstateevents.com"
		        .From = "bob.schneider@gopherstateevents.com"
		        .Subject = Finishers(1, x) & "'s Results for " & sMeetName & "-" & sThisRace
		        .TextBody = sMsg
		        .Send
	        End With
	        Set cdoMessage = Nothing

	        'write these to db
	        sql = "INSERT INTO ResultsSentInd (MeetsID, RaceName, PartName, WhenSent, OtherRecips) VALUES (" & lThisMeet & ", '" & sThisRace & "', '" 
            sql = sql & Finishers(1, x) & "', '" & Now() & "', '" & sMyFollowers & "')"
	        Set rs = conn.Execute(sql)
	        Set rs = Nothing
        End If

        'check for sms send for participant
        If Not Finishers(5, x) & "" = "" Then
            If Not Finishers(6, x) & "" = "" Then
                'mobile message
                sMsg = vbNullString
                sMsg = sMsg & "My Results for " & sMeetName & " on " & dMeetDate & " " & vbCrLf
                sMsg = sMsg & "Name: " & Finishers(1, x) & " " & vbCrLf
                sMsg = sMsg & "Race: " & sThisRace & vbCrLf
                sMsg = sMsg & "Time: " & Finishers(2, x) & " " & vbCrLf
                sMsg = sMsg & "Results at http://www.gopherstateevents.com/results/cc_rslts/cc_rslts.asp?meet_id=" & lThisMeet & "&race_id=" & lThisRace & "&sport=" & sSport

                Set cdoMessage = Server.CreateObject("CDO.Message")
                Set cdoMessage.Configuration = cdoConfig
		        With cdoMessage
                    .From = "bob.schneider@gopherstateevents.com"
			        .To = Finishers(5, x) & GetSendURL(Finishers(6, x))
'                    .To = "bob.schneider@gopherstateevents.com"
			        .TextBody = sMsg
			        .Send
		        End With
	            Set cdoMessage = Nothing

		        'insert into email sent
		        sql = "INSERT INTO RsltsSmsSent (RosterID, RaceID, WhenSent, MobileNum, Provider) VALUES (" & Finishers(3, x) & ", " & lThisRace & ", '" 
                sql = sql & Now() & "', '" & Finishers(5, x) & "', " & Finishers(6, x) & ")"
		        Set rs = conn.Execute(sql)
		        Set rs = Nothing
            End If
        End If

        'send to followers
        'get followers with mobile data
        y = 0
        ReDim MobileData(1, 0)
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT f.CellPhone, f.CellProvider FROM PTFollowers f INNER JOIN PerfTrkr p ON p.PerfTrkrID = f.PerfTrkrID WHERE p.RosterID = " 
        sql = sql & Finishers(3, x)
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            MobileData(0, y) = rs(0).Value
            MobileData(1, y) = rs(1).Value
            y = y + 1
            ReDim Preserve MobileData(1, y)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        For y = 0 To UBound(MobileData, 2) - 1
             If Not MobileData(0, y) & "" = "" Then
                If Not MobileData(1, y) & "" = "" Then
                    Set cdoMessage = Server.CreateObject("CDO.Message")
                    Set cdoMessage.Configuration = cdoConfig
		            With cdoMessage
                        .From = "bob.schneider@gopherstateevents.com"
			            .To = MobileData(0, y) & GetSendURL(MobileData(1, y))
'                        .To = "bob.schneider@gopherstateevents.com"
			            .TextBody = sMsg
			            .Send
		            End With
	                Set cdoMessage = Nothing

		            'insert into email sent
		            sql = "INSERT INTO RsltsSmsSent (RosterID, RaceID, WhenSent, MobileNum, Provider) VALUES (" & Finishers(3, x) & ", " & lThisRace & ", '" 
                    sql = sql & Now() & "', '" & MobileData(0, y) & "', " & MobileData(1, y) & ")"
		            Set rs = conn.Execute(sql)
		            Set rs = Nothing
                End If
            End If
       Next
    Next
End Sub

Private Function GetSendURL(lProviderID)
	If Not CStr(lProviderID) & "" = ""  Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT SendURL FROM CellProviders WHERE CellProvidersID = " & lProviderID
		rs.Open sql, conn2, 1, 2
		If rs.RecordCount > 0 Then GetSendURL = rs(0).Value
		Set rs = Nothing
	End If
End Function

Private Sub SendResults(lThisTeam, lThisRace, sThisRace)
    'get head coach
    sql2 = "SELECT c.Email, t.TeamName, c.CoachesID FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID WHERE t.TeamsID = " & lThisTeam
    Set rs2 = conn.Execute(sql2)
    If rs2(0).Value & "" = "" Then
        sCoachEmail = "bob.schneider@gopherstateevents.com"
    Else
        sCoachEmail = rs2(0).Value
    End If
    sTeamName = rs2(1).Value
    lCoachesID = rs2(2).Value
    Set rs2 = Nothing

    'get other recips
    sOtherRecips = vbNullString
    sql2 = "SELECT Email FROM TeamStaff WHERE CoachesID = " & lCoachesID & " AND SendTo = 'y'"
    Set rs2 = conn.Execute(sql2)
    Do While Not rs2.EOF
        sOtherRecips = sOtherRecips & rs2(0).Value & ";"
        rs2.MoveNext
    Loop
    Set rs2 = Nothing

    If Not sOtherRecips & "" = "" Then sOtherRecips = Left(sOtherRecips, Len(sOtherRecips) - 1)

	sMsg = vbCrLf
	sMsg = sMsg & "Results for " & sMeetName & " on " & dMeetDate & " " & vbCrLf
	sMsg = sMsg & "Race: " & sThisRace & vbCrLf & vbCrLf
				
    If Not sSuppMsg = vbNullString Then sMsg = sMsg & sSuppMsg & " " & vbCrLf & vbCrLf

    sMsg = sMsg & "You will find meet results for this race here: "
	sMsg = sMsg & "http://www.gopherstateevents.com/results/cc_rslts/cc_rslts.asp?meet_id=" & lThisMeet & "&race_id=" & lThisRace & "&sport=" & sSport
    sMsg = sMsg  & vbCrLf & vbCrLf

    sMsg = sMsg & "Below are your team's race participants and their UNOFFICIAL times.  A time of '00:00.000' indicates they did not finish or were "
    sMsg = sMsg & "not read by our electronic system.  PLEASE NOTIFY US ASAP IF ANYTHING LOOKS INCORRECT! " & vbCrLf
    sMsg = sMsg & "BIB NAME TIME " & vbCrLf
    For i = 0 To UBound(Finishers, 2) - 1
        sMsg = sMsg & Finishers(0, i) & "  " & Finishers(1, i) & "  " & Finishers(2, i) & vbCrLf
    Next
    sMsg = sMsg & vbCrLf

    sMsg = sMsg & "Please consider Gopher State Events for all of your meet and fitness event management needs. " & vbCrLf & vbCrLf

	sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
	sMsg = sMSg & "Bob Schneider" & vbCrLf
	sMsg = sMSg & "GSE (Gopher State Events) " & " " & vbCrLf 
	sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
	sMsg = sMsg & "612-720-8427 " & vbCrLf
	
    If sOtherRecips = vbNullString Then
        sAllRecips = sCoachEmail
    Else
        sAllRecips = sCoachEmail & ";" & sOtherRecips
    End If

	'write these to db
	sql2 = "INSERT INTO ResultsSent (MeetsID, RaceName, TeamName, WhenSent, Recipients) VALUES (" & lThisMeet & ", '" & sThisRace & "', '" & sTeamName 
    sql2 = sql2 & "', '" & Now() & "', '" & sAllRecips & "')"
	Set rs2 = conn.Execute(sql2)
	Set rs2 = Nothing
			 
	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
'        .To = "bob.schneider@gopherstateevents.com"
        .To = sCoachEmail
        If Len(sOtherRecips) > 0 Then .CC = sOtherRecips
'        .BCC = "bob.schneider@gopherstateevents.com;"
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "Meet Results for " & sMeetName & " (" & Races(1,  j) & ")"
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
End Sub

Private Sub MeetInfo()
    'get meet name
    sql = "SELECT MeetName, MeetDate, Sport FROM Meets WHERE MeetsID = " & lThisMeet
    Set rs = conn.Execute(sql)
    sMeetName = Replace(rs(0).Value, "''", "'")
    dMeetDate = rs(1).Value
    sSport = rs(2).Value
    Set rs = Nothing

    If sSport = "Nordic Ski" Then
        sSport = "nordic"
    Else
        sSport = "cc"
    End If

    'get races in this meet
    i = 0
    sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisMeet & " ORDER BY ViewOrder, RaceTime"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    Races(0, i) = rs(0).Value
	    Races(1, i) = Replace(rs(1).Value, "''", "'")
	    i = i + 1
	    ReDim Preserve Races(1, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing

    'get meet teams array
    i = 0
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
End Sub

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE CC-Nordic Results Email</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	
	<div class="row">
		<%If Session("role") = "admin" Then%>
            <!--#include file = "../../../includes/admin_menu.asp" -->
        <%Else%>
		    <!--#include file = "../../../staff/staff_menu.asp" -->
        <%End If%>

		<div class="col-sm-10">
		    <%If Session("role") = "admin" Then%>
				<!--#include file = "../manage_meet_nav.asp" -->
            <%End If%>
			
			<h4 class="h4">Email Meet Results For <%=sMeetName%></h4>

            <div class="row">
                <form class="form-inline" name="get_meet" method="post" action="email_results.asp">
                <label for="meets">Select Meet:</label>
                <select class="form-control" name="meets" id="meets" onchange="this.form.submit1.click();">
                    <option value="">&nbsp;</option>
                    <%For i = 0 to UBound(Meets, 2) - 1%>
                        <%If CLng(lThisMeet) = 0 Then%>
                            <option value="<%=Meets(0, i)%>"><%=Meets(1, i)%> (<%=Meets(2, i)%>)</option>
                        <%Else%>
                            <%If CLng(Meets(0, i)) = CLng(lThisMeet) Then%>
                                <option value="<%=Meets(0, i)%>" selected><%=Meets(1, i)%> (<%=Meets(2, i)%>)</option>
                            <%Else%>
                                <option value="<%=Meets(0, i)%>"><%=Meets(1, i)%> (<%=Meets(2, i)%>)</option>
                            <%End If%>
                        <%End If%>
                    <%Next%>
                </select>
                <input type="hidden" name="get_meet" id="get_meet" value="get_meet">
                <input type="submit" class="form-control" name="submit1" id="submit1" value="Get Meet">
                </form>
           </div>			

			<%If CLng(lThisMeet) > 0 Then%>
                <div class="row">
                    <div class="col-sm-6">
                        <h4 class="h4">Send Results Email</h4>
                        <form class="form" name="request_lineup" method="post" action="email_results.asp?meet_id=<%=lThisMeet%>">
                        <table class="table">
                            <tr>
                                <td style="text-align:center" colspan="2">	
                                    <input type="hidden" name="submit_send" id="submit_send" value="submit_send">
                                    <input type="submit" name="submit2" id="submit2" value="Send Results">
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">	
                                    <textarea class="form-control" name="supp_msg" id="supp_msg" rows="3"></textarea>
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <table class="table table-striped">
                                        <tr><th>No.</th><th>Race</th><th>Send</th></tr>
                                        <%For i = 0 To UBound(Races, 2) - 1%>
                                            <tr>
                                                <td><%=i + 1%>)</td>
                                                <td><%=Races(1, i)%></td>
                                                <td><input type="checkbox" name="race_<%=Races(0, i)%>" id="race_<%=Races(0, i)%>"></td>
                                            </tr>
                                        <%Next%>
                                    </table>
                                </td>
                                <td valign="top">
                                    <table class="table table-striped">
                                        <tr><th>No</th><th>Team</th><th>Send</th></tr>
                                        <%For i = 0 to UBound(MeetTeams, 2) - 1%>
                                            <tr>
                                                <td><%=i + 1%>)</td>
                                                <td><%=MeetTeams(1, i)%></td>
                                                <td>
                                                    <input type="checkbox" name="team_<%=MeetTeams(0, i)%>" id="team_<%=MeetTeams(0, i)%>" checked>
                                                </td>
                                            </tr>
                                        <%Next%>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        </form>
                    </div>
                    <div class="col-sm-6">
                        <h4 class="h4">Results Email History</h4>

                        <table class="table table-striped bg-success">
                            <tr><th>No.</th><th>Race</th><th>Team</th><th>Sent</th></tr>
                            <%For i = 0 To UBound(SendHistory, 2) - 1%>
                                <tr>
                                    <td><%=i + 1%></td>
                                    <td><%=SendHistory(0, i)%></td>
                                    <td><%=SendHistory(1, i)%></td>
                                    <td><%=SendHistory(2, i)%></td>
                                </tr>
                            <%Next%>
                        </table>
                    </div>
                </div>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
