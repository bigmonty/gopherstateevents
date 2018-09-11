<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, rs, conn, rs2, sql2, conn2
Dim i, j
Dim lEventID, lEventDirID, lEmailRsltsID
Dim iEventType
Dim sMyEmail, sEventName, sPageToSend, sEventDirEmail, sSuppMsg, sSendWhat
Dim RaceArr(), FinishersArr(), RaceRslts(), SendTo()
Dim cdoMessage, cdoConfig, objEmail, xmlhttp, EmailContents
Dim dWhenSent
Dim bFound, bSendThis

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

lEventID = Request.QueryString("event_id")
If lEventID = vbNullString Then lEventID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Dim Events
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
    
If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_ind") = "submit_ind" Then
    Call EventInfo()

    sSendWhat = Request.Form.Item("ind_send_what")
    If sSendWhat = vbNullString Then sSendWhat = "all"

	sql = "SELECT EventName FROM Events WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
	Set rs = Nothing
	
	For i = 0 to UBound(RaceArr, 2) - 1
        j = 0
        ReDim SendTo(0)
        Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT ir.ParticipantID, p.Email FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID "
        sql = sql & "WHERE ir.RaceID = " & RaceArr(0, i) & " AND p.Email IS NOT NULL"
		rs.Open sql, conn, 1, 2
		Do While Not rs.EOF
            If Request.Form.Item("send_" & rs(0).Value) = "on" Then
                SendTo(j) = rs(0).Value
                j = j + 1
                ReDim Preserve SendTo(j)
            End If
            rs.MoveNext
        Loop
        rs.Close
  		Set rs = Nothing

		For j = 0 To UBound(SendTo) - 1
			'get email address
			sMyEmail = vbNullString

            If j = 0 Then
	            'add this event to the emailrslts table
                sSuppMsg = Request.Form.Item("supp_msg_selected")
                If Not sSuppMsg & "" = "" Then sSuppMsg = Replace(sSuppMsg, "'", "''")
	            sql = "INSERT INTO EmailRslts(EventID, WhenSent, SuppMsg, SendType) VALUES (" & lEventID & ", '" & Now() & "', '" & sSuppMsg  
                sql = sql & "', 'selected')"
	            Set rs = conn.Execute(sql)
	            Set rs = Nothing

                'get email rslts id so the supplemental message can be retrieved
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT EmailRsltsID FROM EmailRslts WHERE EventID = " & lEventID & " ORDER BY EmailRsltsID DESC"
                rs.Open sql, conn, 1, 2
                lEmailRsltsID = rs(0).Value
                rs.Close
                Set rs = Nothing
            End If
    			
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT Email FROM Participant WHERE ParticipantID = " & SendTo(j)
			rs.Open sql, conn, 1, 2
            If Not rs(0).Value & "" = "" Then If DontSend(rs(0).Value) = False Then sMyEmail = rs(0).Value
			rs.Close
			Set rs = Nothing
			
            If sSendWhat = "all" or sSendWhat = "email" Then
			    If Not sMyEmail = vbNullString Then
				    If ValidEmail(sMyEmail) = True Then
                        sPageToSend = "http://www.gopherstateevents.com/perf_center/my_results.asp?event_id=" & lEventID & "&race_id=" & RaceArr(0, i) & "&part_id=" & SendTo(j) & "&email_rslts_id=" & lEmailRsltsID

                        Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	                        xmlhttp.open "GET", sPageToSend, false
	                        xmlhttp.send ""
	                        EmailContents = xmlhttp.responseText
                        Set xmlhttp = nothing

			            Set cdoMessage = CreateObject("CDO.Message")
			            With cdoMessage
				            Set .Configuration = cdoConfig
				            .To = sMyEmail
'                            .To = "bob.schneider@gopherstateevents.com"
				            .From = "bob.schneider@gopherstateevents.com"
                            If j = 0 Then 
                                If Request.Form.Item("event_dir") = "on" Then
                                    .BCC = "bob.schneider@gopherstateevents.com;" & sEventDirEmail
                                Else
                                    .BCC = "bob.schneider@gopherstateevents.com;"
                                End If
                            End If
   				            .Subject = "My Results for " & sEventName
				            .HTMLBody = EmailContents
				            .Send
			            End With
			            Set cdoMessage = Nothing

					    'insert into email sent
					    sql = "INSERT INTO ResultsSent (ParticipantID, RaceID, WhenSent) VALUES (" & SendTo(j) & ", "
					    sql = sql & RaceArr(0, i) & ", '" & Now() & "')"
					    Set rs = conn.Execute(sql)
					    Set rs = Nothing
                    End If
				End If
			End If

            If sSendWhat = "all" Or sSendWhat = "sms" Then Call SendText(SendTo(j), RaceArr(0, i), RaceArr(1, i))
		Next
	Next
ElseIf Request.Form.Item("submit_all") = "submit_all" Then
    Call EventInfo()

    sSendWhat = Request.Form.Item("send_what")
    If sSendWhat = vbNullString Then sSendWhat = "all"

	For i = 0 to UBound(RaceArr, 2) - 1
        j = 0
        ReDim FinishersArr(1, 0)
        Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT ir.ParticipantID, pr.Bib FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID WHERE ir.RaceID = " 
        sql = sql & RaceArr(0, i) 
		rs.Open sql, conn, 1, 2
		Do While Not rs.EOF
            FinishersArr(0, j) = rs(0).Value
            FinishersArr(1, j) = rs(1).Value
            j = j + 1
            ReDim Preserve FinishersArr(1, j)
            rs.MoveNext
        Loop
        rs.Close
  		Set rs = Nothing

		For j = 0 To UBound(FinishersArr, 2) - 1
			bSendThis = True

            If j = 0 Then
	            'add this event to the emailrslts table
                sSuppMsg = Request.Form.Item("supp_msg_batch")
                If Not sSuppMsg & "" = "" Then sSuppMsg = Replace(sSuppMsg, "'", "''")
	            sql = "INSERT INTO EmailRslts(EventID, WhenSent, SuppMsg, SendType) VALUES (" & lEventID & ", '" & Now() & "', '" & sSuppMsg  & "', 'batch')"
	            Set rs = conn.Execute(sql)
	            Set rs = Nothing

                'get email rslts id so the supplemental message can be retrieved
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT EmailRsltsID FROM EmailRslts WHERE EventID = " & lEventID & " ORDER BY EmailRsltsID DESC"
                rs.Open sql, conn, 1, 2
                lEmailRsltsID = rs(0).Value
                rs.Close
                Set rs = Nothing
            End If
			
			If Request.Form.Item("no_resend") = "on" Then
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT ResultsSentID FROM ResultsSent WHERE ParticipantID = " & FinishersArr(0, j) & " AND RaceID = " & RaceArr(0, i)
                sql = sql & "AND Bib = "  & FinishersArr(1, j)
				rs.Open sql, conn, 1, 2
				If rs.RecordCount > 0 Then bSendThis = False
				rs.Close
				Set rs = Nothing
			End If
			
			If bSendThis = True Then
				'get email address
				sMyEmail = vbNullString
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT Email FROM Participant WHERE ParticipantID = " & FinishersArr(0, j)
				rs.Open sql, conn, 1, 2
                If Not rs(0).Value & "" = "" Then 
                    If DontSend(rs(0).Value) = False Then sMyEmail = rs(0).Value
                End If
				rs.Close
				Set rs = Nothing

                If sSendWhat = "all" Or sSendWhat = "email" Then
				    If Not sMyEmail = vbNullString Then
					    If ValidEmail(sMyEmail) = True Then

                            sPageToSend = "http://www.gopherstateevents.com/perf_center/my_results.asp?event_id=" & lEventID & "&race_id=" & RaceArr(0, i) & "&part_id=" & FinishersArr(0, j) & "&email_rslts_id=" & lEmailRsltsID

                            Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	                            xmlhttp.open "GET", sPageToSend, false
	                            xmlhttp.send ""
	                            EmailContents = xmlhttp.responseText
                            Set xmlhttp = nothing

			                Set cdoMessage = CreateObject("CDO.Message")
			                With cdoMessage
				                Set .Configuration = cdoConfig
				                .To = sMyEmail
    '                            .To = "bob.schneider@gopherstateevents.com"
				                .From = "bob.schneider@gopherstateevents.com"
                                If j = 0 Then 
                                   If Request.Form.Item("evnt_dir_send") = "on" Then
                                        .BCC = "bob.schneider@gopherstateevents.com;" & sEventDirEmail
                                    Else
                                        .BCC = "bob.schneider@gopherstateevents.com;"
                                    End If
                                End If
   				                .Subject = "My Results for " & sEventName
				                .HTMLBody = EmailContents
				                .Send
			                End With
			                Set cdoMessage = Nothing

						    'insert into email sent
						    sql = "INSERT INTO ResultsSent (ParticipantID, RaceID, WhenSent, Bib) VALUES (" & FinishersArr(0, j) & ", "
						    sql = sql & RaceArr(0, i) & ", '" & Now() & "', " & FinishersArr(1, j) & ")"
						    Set rs = conn.Execute(sql)
						    Set rs = Nothing
					    End If
                    End If
				End If

                If sSendWhat = "all" or sSendWhat = "sms" Then Call SendText(FinishersArr(0, j), RaceArr(0, i), RaceArr(1, i))
			End If
		Next
	Next
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

If CLng(lEventID) > 0 Then
    Call EventInfo()

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT WhenSent FROM EmailRslts WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then dWhenSent = rs(0).Value
    rs.Close
    Set rs = Nothing
End If

%>
<!--#include file = "../../includes/valid_email.asp" -->
<%

Private Sub SendText(lThisPart, lThisRace, sRaceName)
    Dim sMobileNumber, sMsg, sMyTime
    Dim iMyBib
    Dim lCellProvider

    lCellProvider = 0

    'check for entry in table
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MobileNumber, CellProvider FROM MobileSettings WHERE PartID = " & lThisPart
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sMobileNumber = rs(0).Value
        lCellProvider = rs(1).Value
    End If
    rs.Close
    Set rs = Nothing

    'send if exists
    If CLng(lCellProvider) > 0 Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT FnlScnds FROM IndResults WHERE RaceID = " & lThisRace & " AND ParticipantID = " & lThisPart
		rs.Open sql, conn, 1, 2
        sMyTime = ConvertToMInutes(rs(0).Value)
		rs.Close
		Set rs = Nothing

		sql = "SELECT Bib FROM PartRace WHERE RaceID = " & lThisRace & " AND ParticipantID = " & lThisPart
		Set rs = conn.Execute(sql)
        iMyBib = rs(0).Value
		Set rs = Nothing

        sMsg = "UNOFFICIAL TIME: " & sMyTime 
        sMsg = sMsg & " Results @ www.gopherstateevents.com/results/fitness_events/results.asp?event_id=" & lEventID & "&race_id=" & lThisRace 

        Set cdoMessage = Server.CreateObject("CDO.Message")
        Set cdoMessage.Configuration = cdoConfig
		With cdoMessage
            .From = "bob.schneider@gopherstateevents.com"
			.To = sMobileNumber & GetSendURL(lCellProvider)
'            .To = "bob.schneider@gopherstateevents.com"
			.TextBody = sMsg
			.Send
		End With
	    Set cdoMessage = Nothing

		'insert into email sent
		sql = "INSERT INTO RsltsSmsSent (PartID, RaceID, WhenSent, Bib) VALUES (" & lThisPart & ", " & RaceArr(0, i) & ", '" & Now() & "', " & iMyBib & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    End If
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

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
   
    'accomodate a '0' value
    If CSng(sglScnds) <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(CSng(sglScnds) \ 60)
    sSecondPart = CStr(((CSng(sglScnds) / 60) - (CSng(sglScnds) \ 60)) * 60)
    
    'add leading zero to seconds if necessary
    If CSng(sSecondPart) < 10 Then
        sSecondPart = "0" & sSecondPart
    End If
    
    'make sure there are exactly two decimal places
    If Len(sSecondPart) < 5 Then
        If Len(sSecondPart) = 2 Then
            sSecondPart = sSecondPart & ".00"
        ElseIf Len(sSecondPart) = 4 Then
            sSecondPart = sSecondPart & "0"
        End If
    Else
        sSecondPart = Left(sSecondPart, 5)
    End If
    
    'do the conversion
    If CInt(sMinutePart) <= 60 Then
        ConvertToMinutes = sMinutePart & ":" & sSecondPart
    Else
        sHourPart = CStr(CSng(sMinutePart) \ 60)
        sMinutePart = CStr(CSng(sMinutePart) Mod 60)

        If Len(sMinutePart) = 1 Then
            sMinutePart = "0" & sMinutePart
        End If

        ConvertToMinutes = sHourPart & ":" & sMinutePart & ":" & sSecondPart
    End If
End Function

Function DontSend(sThisEmail) 
	Dim rs2, sql2

    DontSend = False

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Email FROM DontSend WHERE Email = '" & sThisEmail & "' AND (DontSend = 'all' OR DontSend = 'results')"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then DontSend = True
    rs2.Close
    Set rs2 = Nothing
End Function

Private Sub EventInfo()
    sql = "SELECT EventName, EventDirID, EventType FROM Events WHERE EventID = " & lEventID
    Set rs = conn.Execute(sql)
    sEventName = Replace(rs(0).Value, "''", "'")
    lEventDirID = rs(1).Value
    iEventType = rs(2).Value
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Email FROM EventDir WHERE EventDirID = " & lEventDirID
    rs.Open sql, conn, 1, 2
    sEventDirEmail = rs(0).Value
    rs.Close
    Set rs = Nothing

    i = 0
    ReDim RaceArr(1, 0)
    sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    RaceArr(0, i) = rs(0).Value
        RaceArr(1, i) = Replace(rs(1).Value, "''", "'")
	    i = i + 1
	    ReDim Preserve RaceArr(1, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub GetRaceResults(lThisRace)
	Dim x

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT pr.AgeGrp, pr.RaceID, pr.Age, p.Gender FROM PartRace pr INNER JOIN Participant p ON pr.ParticipantID = p.ParticipantID "
    sql = sql & "WHERE pr.RaceID = " & lThisRace & " AND pr.AgeGrp IS NULL"
	rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If rs(0).Value & "" = "" Then 
            rs(0).Value = GetAgeGrp(rs(3).Value, rs(2).Value, rs(1).Value)
            rs.Update
        End If
        rs.MoveNext
    Loop
	rs.Close
	Set rs = Nothing
	
	x = 0
	ReDim RaceRslts(6, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, p.Gender, ir.FnlTime, p.Email, pr.Bib FROM Participant p INNER JOIN IndResults ir "
	sql = sql & "ON p.ParticipantID = ir.ParticipantID INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID WHERE ir.RaceID = " & lThisRace
    sql = sql & " AND pr.RaceID = " & lThisRace & " ORDER BY p.LastName, p.FirstName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		RaceRslts(0, x) = rs(0).Value
		RaceRslts(1, x) = rs(6).Value & "-" & Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
		RaceRslts(2, x) = rs(3).Value
        RaceRslts(3, x) = rs(4).Value
		RaceRslts(4, x) = GetMySend(rs(0).Value, lThisRace)
		RaceRslts(5, x) = rs(5).Value
        RaceRslts(6, x) = GetMyPhone(rs(0).Value)

        x = x + 1
		ReDim Preserve RaceRslts(6, x)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End Sub

Private Function GetMyPhone(lThisPart)
	GetMyPhone = vbNullString
	
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT MobileNumber FROM MobileSettings WHERE PartID = " & lThisPart
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetMyPhone = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetMySend(lThisPart, lThisRaceID)
	GetMySend = vbNullString
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT WhenSent FROM ResultsSent WHERE RaceID = " & lThisRaceID & " AND ParticipantID = " & lThisPart & " ORDER BY WhenSent DESC"
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetMySend = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function

Public Function GetAgeGrp(sMF, iAge, lThisRaceID)
    Dim iBegAge, iEndAge
    
    iBegAge = 0
    
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sMF & "' AND RaceID = " & lThisRaceID & " ORDER BY EndAge DESC"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        If CInt(iAge) <= CInt(rs2(0).Value) Then
            iEndAge = rs2(0).Value
        Else
            iBegAge = CInt(rs2(0).Value) + 1
            Exit Do
        End If
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If iBegAge = 0 Then
        GetAgeGrp = iEndAge & " and Under"
    Else
        If iEndAge = 110 Then
            GetAgeGrp = CInt(iBegAge) & " and Over"
        Else
            GetAgeGrp = CInt(iBegAge) & " - " & iEndAge
        End If
    End If
End Function

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Email Results</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<%If Session("role") = "admin" Then%>
            <!--#include file = "../../includes/admin_menu.asp" -->
        <%Else%>
		    <!--#include file = "../../staff/staff_menu.asp" -->
        <%End If%>

		<div class="col-md-10">
			<h3 class="h3">RaceWare Email Results for <%=sEventName%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="email_results.asp?event_id=<%=lEventID%>">
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
				
            <%If CLng(lEventID) > 0 Then%>
		        <%If Session("role") = "admin" Then%>
                    <!--#include file = "../../includes/event_nav.asp" -->
                    <!--#include file = "email_nav.asp" -->
                <%End If%>

			    <%If CStr(dWhenSent) = vbNullString Then%>
				    <p>Results for this event have not been sent.</p>
			    <%Else%>
				    <p>Results for this event were last sent on <%=dWhenSent%>.</p>
			    <%End If%>

			    <h4 class="h4">Batch Send</h4>
				
			    <div style="background-color:#ececd8;margin-bottom: 10px;">
				    <form name="send_all" Method="Post" action="email_results.asp?event_id=<%=lEventID%>">
                    <input type="checkbox" name="evnt_dir_send" id="evnt_dir_send">&nbsp;Send To Event Director&nbsp;&nbsp;
				    <input type="checkbox" name="no_resend" id="no_resend" checked>&nbsp;No Resend

                    <div class="bg-warning">
                        <input type="radio" name="send_what" id="ind_send_what" value="all" checked>&nbsp;All &nbsp;&nbsp;
                        <input type="radio" name="send_what" id="ind_send_what" value="email">&nbsp;Email Only &nbsp;&nbsp;
                        <input type="radio" name="send_what" id="ind_send_what" value="sms">&nbsp;SMS Only
                    </div>

                    <textarea name="supp_msg_batch" id="supp_msg_batch" rows="3" cols="100" style="font-size: 1.1em"></textarea>
                    <br>
				    <input type="hidden" name="submit_all" id="submit_all" value="submit_all">
				    <input type="submit" name="submit2" id="submit2" value="Send All">
				    </form>
			    </div>

                <hr>

			    <h4 class="h4">Send Record and Individual Send</h4>

			    <form name="send_ind" method="Post" action="email_results.asp?event_id=<%=lEventID%>">
			    <div>
				    <input type="hidden" name="submit_ind" id="submit_ind" value="submit_ind">
				    <input type="submit" name="submit2" id="submit2" value="Send Selected"><br>
                    <input type="checkbox" name="event_dir" id="event_dir">&nbsp;Send To Event Director

                    <div class="bg-warning">
                        <input type="radio" name="ind_send_what" id="ind_send_what" value="all" checked>&nbsp;All &nbsp;&nbsp;
                        <input type="radio" name="ind_send_what" id="ind_send_what" value="email">&nbsp;Email Only &nbsp;&nbsp;
                        <input type="radio" name="ind_send_what" id="ind_send_what" value="sms">&nbsp;SMS Only
                    </div>

                    <textarea name="supp_msg_selected" id="supp_msg_selected" rows="3" cols="100" style="font-size: 1.1em"></textarea>
			    </div>

			    <%For i = 0 To UBound(RaceArr, 2) - 1%>
				    <%Call GetRaceResults(RaceArr(0, i))%>
					
				    <h4 style="background: none;color: #000;"><%=RaceArr(1, i)%>&nbsp;(<%=UBound(RaceRslts, 2)%> Finishers)</h4>
					
				    <table class="table table-striped table-condensed">
					    <tr>	
						    <th>Pl.</th>
						    <th>Bib-Name</th>
						    <th>M/F</th>
						    <th>Time</th>
                            <th>Email</th>
                            <th>Phone</th>
						    <th>Last Send</th>
						    <th>Send</th>
					    </tr>
					    <%For j = 0 To UBound(RaceRslts, 2) - 1%>
							<tr>	
								<td><%=j +1%>)</td>
								<td><%=RaceRslts(1, j)%></td>
								<td><%=RaceRslts(2, j)%></td>
								<td><%=RaceRslts(3, j)%></td>
								<td><a href="mailto:<%=RaceRslts(5, j)%>"><%=RaceRslts(5, j)%></a></td>
                                <td><%=RaceRslts(6, j)%></td>
                                <td><%=RaceRslts(4, j)%></td>
								<td style="text-align:center;">
									<input type="checkbox" name="send_<%=RaceRslts(0, j)%>" id="send_<%=RaceRslts(0, j)%>">
								</td>
							</tr>
					    <%Next%>
				    </table>
			    <%Next%>
			    </form>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
