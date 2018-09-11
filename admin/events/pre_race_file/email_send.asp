<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, rs, conn, rs2, sql2
Dim i, j, k
Dim lEventID, lEventDirID
Dim sMyEmail, sSuppMsg, sEventName
Dim RaceArr(), PartArr(), MissingParts()
Dim cdoMessage, cdoConfig, objEmail, xmlhttp, EmailContents, sPageToSend, sEventDirEmail
Dim dWhenSent
Dim bFound
Dim bSendThis

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

lEventID = Request.QueryString("event_id")

ReDim MissingParts(0)

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

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

Call EventInfo()

sSuppMsg = Request.Form.Item("supp_msg_all")
If Not sSuppMsg = vbNullString Then sSuppMsg = Replace(sSuppMsg, chr(34), "")
	
For i = 0 to UBound(RaceArr, 2) - 1
j = 0
ReDim PartArr(0)
sql = "SELECT ParticipantID FROM PartRace WHERE RaceID = " & RaceArr(0, i)
Set rs = Conn.Execute(sql)
Do While Not rs.EOF
	PartArr(j) = rs(0).Value
	j = j + 1
	ReDim Preserve PartArr(j)
	rs.MoveNext
Loop
Set rs = Nothing
		
k = 0

For j = 0 to UBound(PartArr) - 1
	bSendThis = True
			
	If Request.Form.Item("no_resend") = "on" Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT ParticipantID, RaceID FROM PreRaceRecips WHERE ParticipantID = " & PartArr(j) & " AND RaceID = " & RaceArr(0, i)
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then 	bSendThis = False
		rs.Close
		Set rs = Nothing
	End If
	
	If bSendThis = True Then
		'get email address
		sMyEmail = vbNullString
				
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT Email FROM Participant WHERE ParticipantID = " & PartArr(j)
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then
            If Not rs(0).Value & "" = "" Then  sMyEmail = rs(0).Value
'                        If DontSend(rs(0).Value) = False Then sMyEmail = rs(0).Value
'                    End If
        Else
            MissingParts(k) = PartArr(j)
            k = k + 1
            ReDim Preserve MissingParts(k)
		End If
		rs.Close
		Set rs = Nothing

		If Not sMyEmail & "" = "" Then
			If ValidEmail(sMyEmail) = True Then
                sPageToSend = "http://www.gopherstateevents.com/misc/pre_race.asp?event_id=" & lEventID & "&race_id=" & RaceArr(0, i) & "&part_id=" & PartArr(j) & "&supp_msg=" & sSuppMsg

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
'                            .To = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
                If j = 0 Then .BCC = "bob.schneider@gopherstateevents.com;" & sEventDirEmail
   				.Subject = "Information for " & sEventName
				.HTMLBody = EmailContents
				.Send
			End With
			Set cdoMessage = Nothing
            Set cdoConfig = Nothing

					    'insert into email sent
					    sql = "INSERT INTO PreRaceRecips (ParticipantID, RaceID, WhenSent) VALUES (" & PartArr(j) & ", " & RaceArr(0, i) & ", '" & Now() & "')"
					    Set rs = conn.Execute(sql)
					    Set rs = Nothing
					End If
				End If
			End If
		Next
	Next
	
	'add this event to the emailrslts table
	sql = "INSERT INTO PreRaceSent(EventID, DateSent) VALUES (" & lEventID & ", '" & Now() & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

If CLng(lEventID) > 0 Then Call EventInfo()

Private Sub EventInfo()
    sql = "SELECT EventName, EventDirID FROM Events WHERE EventID = " & lEventID
    Set rs = conn.Execute(sql)
    sEventName = Replace(rs(0).Value, "''", "'")
    lEventDirID = rs(1).Value
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT DateSent FROM PreRaceSent WHERE EventID = " & lEventID
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
	    RaceArr(1, i) = Replace(rs(1).Value, "''", "'")
	    i = i + 1
	    ReDim Preserve RaceArr(1, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

%>
<!--#include file = "../../includes/valid_email.asp" -->
<%

Function DontSend(sThisEmail) 
	Dim rs2, sql2

    DontSend = False

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Email FROM DontSend WHERE Email = '" & sThisEmail & "' AND (DontSend = 'all' OR DontSend = 'pre-race')"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then DontSend = True
    rs2.Close
    Set rs2 = Nothing
End Function

Private Sub GetRaceInfo(lThisRace)
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
	ReDim PartArr(4, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, p.Email, pr.AgeGrp FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID "
    sql = sql & "WHERE pr.RaceID = " & lThisRace & " ORDER BY p.LastName, p.FirstName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		PartArr(0, x) = rs(0).Value
		PartArr(1, x) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
		PartArr(2, x) = rs(3).Value
		PartArr(3, x) = GetMySend(rs(0).Value, lThisRace)
        partArr(4, x) = rs(4).Value
        x = x + 1
		ReDim Preserve PartArr(4, x)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End Sub

Private Function GetMySend(lThisPart, lThisRaceID)
	GetMySend = vbNullString
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT WhenSent FROM PreRaceRecips WHERE RaceID = " & lThisRaceID & " AND ParticipantID = " & lThisPart & " ORDER BY WhenSent DESC"
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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Pre-Race Email Send</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">GSE Pre-Race Email Utility for <%=sEventName%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="email_pre-race.asp?event_id=<%=lEventID%>">
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

			<%If Not Clng(lEventID) = 0 Then%>
                <!--#include file = "../../includes/event_nav.asp" -->
                <!--#include file = "email_nav.asp" -->
            			
			    <%If CStr(dWhenSent) = vbNullString Then%>
				    <p>A pre-race email for this event has not been sent.</p>
			    <%Else%>
				    <p>Pre-race emails for this event were last sent on <%=dWhenSent%>.</p>
			    <%End If%>
				
			    <%If UBound(MissingParts) > 0 Then%>
                    <div style="background-color:#ececd8;">
				        <h4 class="h4">Missing Participants:</h4>
                        <ul>
                            <%For i = 0 To UBound(MissingParts) - 1%>
                                <li><%=MissingParts(i)%></li>
                            <%Next%>
                        </ul>
			        </div>
                <%End If%>				
				
			    <div class="bg-info" style="text-align: center;padding:5px;"">
				    <form name="sample_email" Method="Post" action="email_pre-race.asp?event_id=<%=lEventID%>">
                    <h4 style="text-align: left;">Send Sample</h4>
                    <p style="text-align: left;">
                        Supplemental Message: 
                        <br>
                        <textarea name="supp_msg_sample" id="supp_msg_sample" rows="3" cols="100"></textarea>
                        <br>
                        Send To: 
                        <br>
                        <textarea name="sample_recips" id="sample_recips" rows="2" cols="100">bob.schneider@gopherstateevents.com</textarea>
                    </p>
				    <input type="hidden" name="submit_sample" id="submit_sample" value="submit_sample">
				    <input type="submit" name="submit1" id="submit1" value="Send Sample">
				    </form>
			    </div>
				
			    <div class="bg-warning" style="text-align: center;padding:5px;"">
				    <form name="send_all" Method="Post" action="email_pre-race.asp?event_id=<%=lEventID%>">
                    <h4 style="text-align: left;">Send To All</h4>
                    <p style="text-align: left;">
                        Supplemental Message: 
                        <br>
                        <textarea name="supp_msg_all" id="supp_msg_all" rows="3" cols="100"></textarea>
                    </p>
                    <br>
				    <input type="checkbox" name="no_resend" id="no_resend">&nbsp;No Resend
				    <input type="hidden" name="submit_all" id="submit_all" value="submit_all">
				    <input type="submit" name="submit2" id="submit2" value="Send All">
				    </form>
			    </div>
				
			    <div class="bg-danger" style="text-align: center;padding:5px;">
				    <h4 style="text-align: left;"><a href="javascript:pop('pre_race_file/email_file.asp?event_id=<%=lEventID%>',600,700)">Send From File</a></h4>
			    </div>

                <div class="bg-success" style="padding:5px;">
			        <form name="send_ind" method="Post" action="email_pre-race.asp?event_id=<%=lEventID%>">
                    <h4 class="h4">Send To Selected</h4>
			        <div style="text-align:center;margin-top: 10px;">
                       <p style="text-align: left;">
                            Supplemental Message: 
                            <br>
                            <textarea name="supp_msg_ind" id="supp_msg_ind" rows="3" cols="100"></textarea>
                        </p>
				        <input type="hidden" name="submit_ind" id="submit_ind" value="submit_ind">
				        <input type="submit" name="submit4" id="submit4" value="Send Selected">
 			        </div>
			        <%For i = 0 To UBound(RaceArr, 2) - 1%>
				        <%Call GetRaceInfo(RaceArr(0, i))%>
					
                        <h4 style="padding-top: 10px;"><%=RaceArr(1, i)%></h4>
				        <table>
					        <tr>	
                                <th>No.</th>
						        <th>Name</th>
                                <th>Age Grp</th>
                                <th>Email</th>
						        <th>Last Send</th>
						        <th>Send</th>
					        </tr>
					        <%For j = 0 To UBound(PartArr, 2) - 1%>
						        <%If j mod 2 = 0 Then%>
							        <tr>	
								        <td class="alt"><%=j +1%>)</td>
								        <td class="alt"><%=PartArr(1, j)%></td>
                                        <td class="alt"><%=PartArr(4, j)%></td>
                                        <td class="alt"><a href="mailto:<%=PartArr(2, j)%>"><%=PartArr(2, j)%></a></td>
                                        <td class="alt"><%=PartArr(3, j)%></td>
								        <td class="alt" style="text-align:center;">
									        <input type="checkbox" name="send_<%=PartArr(0, j)%>" id="send_<%=PartArr(0, j)%>">
								        </td>
							        </tr>
						        <%Else%>
							        <tr>	
								        <td><%=j +1%>)</td>
								        <td><%=PartArr(1, j)%></td>
                                        <td class="alt"><%=PartArr(4, j)%></td>
                                        <td class="alt"><a href="mailto:<%=PartArr(2, j)%>"><%=PartArr(2, j)%></a></td>
                                        <td><%=PartArr(3, j)%></td>
								        <td style="text-align:center;">
									        <input type="checkbox" name="send_<%=PartArr(0, j)%>" id="send_<%=PartArr(0, j)%>">
								        </td>
							        </tr>
						        <%End If%>
					        <%Next%>
				        </table>
			        <%Next%>
			        </form>
                </div>
            <%End If%>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
