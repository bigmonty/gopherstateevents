<%@ Language=VBScript%>

<%
Option Explicit

Dim sql, rs, conn
Dim lMsgLogID, lMsgRecip, lCoachID
Dim i, j, k
Dim strEmail(), BadEmail(), SendTo(), strCoachEmail
Dim sSubject, sMsg, sSendHow, lSender, sTeamIDs
Dim cdoMessage, cdoConfig

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

Server.ScriptTimeout = 1200

sSubject = Request.Form.Item("subject")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 30
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Session("role") = "coach" Then
    lCoachID = Session("my_id")
Else
    lCoachID = Session("team_coach_id")
End If

'get team staff
Set rs = Server.CreateObject("ADODB.Recordset")
If Session("role") = "coach" Then
    sql = "SELECT TeamsID FROM Teams WHERE CoachesID = " & Session("my_id")
Else
    sql = "SELECT TeamsID FROM Teams WHERE CoachesID = " & Session("team_coach_id")
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sTeamIDs = sTeamIDs & rs(0).Value & ", "
	rs.MoveNext
Loop
rs.Close
Set rs=Nothing

sTeamIDs = Left(sTeamIDs, Len(sTeamIDs) - 2)

'get head coach email
Set rs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT Email FROM Coaches WHERE CoachesID = " & lCoachID
rs.Open sql, conn, 1, 2
strCoachEmail = rs(0).Value
Set rs = Nothing

'get participant names for the email
i = 0	
ReDim strEmail(2, 0)
Set rs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT RosterID, Email FROM Roster WHERE TeamsID IN (" & sTeamIDs & ") AND Archive = 'n' ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If Request.Form.Item("all_parts") = "on" Then
		If Request.Form.Item("p_e_" & rs(0).Value) = "y" Then
			If Not rs(1).Value & "" = "" Then
				strEmail(0, i) = rs(0).Value
				strEmail(1, i) = rs(1).Value & "; "
                strEmail(2, i) = "Participant"
				i = i + 1
				ReDim Preserve strEmail(2, i)
			End If
		End If
	Else
		If Request.Form.Item("part_" & rs(0).Value) = "on" Then
			If Request.Form.Item("p_e_" & rs(0).Value) = "y" Then
				If Not rs(1).Value & "" = "" Then
				    strEmail(0, i) = rs(0).Value
				    strEmail(1, i) = rs(1).Value & "; "
                    strEmail(2, i) = "Participant"
				    i = i + 1
				    ReDim Preserve strEmail(2, i)
				End If
			End If
		End If
	End If
	rs.MoveNext
Loop
rs.Close
Set rs =Nothing

'get staff names for the email
Set rs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamStaffID, Email FROM TeamStaff WHERE CoachesID = " & lCoachID & " ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If Request.Form.Item("all_staff") = "on" Then
		If Request.Form.Item("s_e_" & rs(0).Value) = "y" Then
			If Not rs(1).Value & "" = "" Then
				strEmail(0, i) = rs(0).Value
				strEmail(1, i) = rs(1).Value & "; "
                strEmail(2, i) = "Team Staff"
				i = i + 1
				ReDim Preserve strEmail(2, i)
			End If
		End If
	Else
		If Request.Form.Item("staff_" & rs(0).Value) = "on" Then
			If Request.Form.Item("s_e_" & rs(0).Value) = "y" Then
				If Not rs(1).Value & "" = "" Then
				    strEmail(0, i) = rs(0).Value
				    strEmail(1, i) = rs(1).Value & "; "
                    strEmail(2, i) = "Team Staff"
				    i = i + 1
				    ReDim Preserve strEmail(2, i)
				End If
			End If
		End If
	End If
	rs.MoveNext
Loop
rs.Close
Set rs =Nothing

'get team contact names for the email
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamContactsID, Email FROM TeamContacts WHERE TeamsID IN (" & sTeamIDs & ") ORDER BY ContactName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If Request.Form.Item("all_contacts") = "on" Then
		If Request.Form.Item("c_e_" & rs(0).Value) = "y" Then
			If Not rs(1).Value & "" = "" Then
				strEmail(0, i) = rs(0).Value
				strEmail(1, i) = rs(1).Value & "; "
                strEmail(2, i) = "Team Contact"
				i = i + 1
				ReDim Preserve strEmail(2, i)
			End If
		End If
	Else
		If Request.Form.Item("contact_" & rs(0).Value) = "on" Then
			If Request.Form.Item("c_e_" & rs(0).Value) = "y" Then
				If Not rs(1).Value & "" = "" Then
				    strEmail(0, i) = rs(0).Value
				    strEmail(1, i) = rs(1).Value & "; "
                    strEmail(2, i) = "Team Contact"
				    i = i + 1
				    ReDim Preserve strEmail(2, i)
				End If
			End If
		End If
	End If
	rs.MoveNext
Loop
rs.Close
Set rs =Nothing

j = 0
k = 0
ReDim BadEmail(0)
ReDim SendTo(2, 0)
For i = 0 to UBound(strEmail, 2) - 1
	strEmail(1, i) = Trim(strEmail(1, i))
	
	If ValidEmail(strEmail(1, i)) = True Then
		SendTo(0, k) = strEmail(0, i)
		SendTo(1, k) = strEmail(1, i)
        SendTo(2, k) = strEmail(2, i)
		k = k + 1
		ReDim Preserve SendTo(2, k)
    Else
		BadEmail(j) = strEmail(1, i)
		j = j + 1
		ReDim Preserve BadEmail(j)
	End If
Next

'get sender into subject/body
sSubject = sSubject & " (from " & Session("my_name") & ")"
sMsg = "From: " &  Session("my_name") & vbCrLf & vbCrLf & Request.Form.Item("message")
lSender = Session("my_id")

%>
<!--#include file = "../../../../includes/cdo_connect.asp" -->
<%

For j = 0 To UBound(SendTo, 2) - 1
    If j = 0 Then
        'log this send
        sql = "INSERT INTO MsgLog (MsgType, SenderID, SenderRole, WhenSent, Msg, Email, Subject) VALUES ('Email', " & lSender & ", '" 
        sql = sql & Session("role") & "', '" & Now & "', '" & Replace(Request.Form.Item("Message"), "'", "''") & "', '" & Session("my_email") 
        sql = sql & "', '" & Left(sSubject, 50) & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        'get message id
        sql = "SELECT MsgLogID FROM MsgLog WHERE SenderID = " & Session("my_id") & " AND SenderRole = '" & Session("role") & "' ORDER BY MsgLogID DESC"
        Set rs = conn.Execute(sql)
        lMsgLogID = rs(0).Value
        Set rs = Nothing 
    End If

    'insert recipient
    sql = "INSERT INTO MsgRecips (MsgLogID, Recipient, Email, Role) VALUES (" & lMsgLogID & ", " & SendTo(0, j) & ", '" & SendTo(1, j) & "', '" 
    sql = sql & SendTo(2, j) & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    Set cdoMessage = CreateObject("CDO.Message")
    With cdoMessage
	    Set .Configuration = cdoConfig

        .To = SendTo(1, j)
	    If j = 0 Then .CC = strCoachEmail

	    If Len(Session("my_email")) > 0 Then
		    .From = Session("my_email")
	    Else
		    .From = "bob.schneider@gopherstateevents.com"		'in case the sender does not have an email entered for their account
	    End If

        .Subject = sSubject
	    .TextBody = sMsg
	    .Send
    End With
    Set cdoMessage = Nothing
Next

Dim sSentBy, sLogMsg
sSentBy = Session("my_name")

sLogMsg = "An email was successfully sent from Gopher State Events.  The communcation was sent by " & sSentBy & " (" & Session("role") & "). "
sLogMsg = sLogMsg & "Sender's email address is " & Session("my_email") & ". The following are the specifics of the message: " & vbCrLf & vbCrLf

sLogMsg = sLogMsg & "When Sent: " & Now & "; " & vbCrLf
sLogMsg = sLogMsg & "Num Recips: " & UBound(SendTo, 2) & "; " & vbCrLf
sLogMsg = sLogMsg & "Msg Size: " & Len(sMsg) & "; " & vbCrLf
sLogMsg = sLogMsg & "Sent From: Gopher State Events"

Set cdoMessage = CreateObject("CDO.Message")
With cdoMessage
	Set .Configuration = cdoConfig
    .From = "bob.schneider@gopherstateevents.com"
    .To = "bob.schneider@gopherstateevents.com"
    .Subject = "Gopher State Events Email Sent"
    .TextBody = sLogMsg
	.Send
End With
Set cdoMessage = Nothing
Set cdoConfig = Nothing

'log this send to me
sql = "INSERT INTO SentMail (WhenSent, SentBy, Role, NumRecips, MsgSize, SentFrom) VALUES ('" & Now & "', '" & Session("my_email")
sql = sql & "', '" & Session("role") & "', " & UBound(SendTo, 2) & ", " & Len(sMsg) & ", 'Gopher State Events')"
Set rs = conn.Execute(sql)
Set rs = Nothing

%>
<!--#include file = "../../../../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>Gopher State Events&reg; Send Group Email</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<!--#include file = "communications_nav.asp" --
			<h4 class="h4">Gopher State Events<sup>&reg;</sup> Email Send Report</h4>
			<%If UBound(BadEmail) > 0 Then%>
				<p>The following were determined to be invalid addresses and were not sent:</p>
				<ul class="list-group">
					<%For i = 0 to UBound(BadEmail) - 1%>
						<li class="list-group-item"><%=BadEmail(i)%></li>
					<%Next%>
				</ul>
			<%Else%>
				<p>Your email was successfully sent.</p>
			<%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
