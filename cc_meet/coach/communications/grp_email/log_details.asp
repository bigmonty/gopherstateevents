<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lMsgLogID, lCoachID, lSender
Dim sSubject, sMsg, sRole, sEmail, strCoachEmail, sTeamIDs
Dim dWhenSent
Dim EmailParts(), EmailStaff(), Contacts(), Recipients(), strEmail()
Dim bFound
Dim cdoMessage, cdoConfig

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

lMsgLogID = Request.QueryString ("msg_log_id")

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

'get recipients
i = 0
ReDim Recipients(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MsgRecipsID, Recipient, Role FROM MsgRecips WHERE MsgLogID = " & lMsgLogID & " ORDER BY Role"
rs.Open sql, conn, 1, 2
Do While Not rs.eof
	Recipients(0, i) = rs(0).Value
	Recipients(1, i) = rs(1).Value
    Recipients(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve Recipients(2, i)
	rs.MoveNext
Loop
Set rs=Nothing

'get message details
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SenderID, SenderRole, WhenSent, Msg, Email, Subject FROM MsgLog WHERE MsgLogID = " & lMsgLogID
rs.Open sql, conn, 1, 2
lSender = rs(0).Value
sRole = rs(1).Value
dWhenSent = rs(2).Value
sMsg = Replace(rs(3).Value, "''", "'")
sEmail = rs(4).Value
sSubject = Replace(rs(5).Value, "''", "'")
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_this") = "submit_this" Then
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
    sql = "SELECT RosterID, Email FROM Rsoter WHERE TeamsID IN (" & sTeamIDs & ") AND Archive = 'n' ORDER BY LastName, FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
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
	    rs.MoveNext
    Loop
    rs.Close
    Set rs =Nothing

    'get staff names for the email
    Set rs=Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamStaffID, Email FROM TeamSTaff WHERE CoachesID = " & lCoachID & " ORDER BY LastName, FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
		If Request.Form.Item("staff_" & rs(0).Value) = "on" Then
			If Request.Form.Item("s_e_" & rs(0).Value) = "y" Then
				If Not rs(1).Value & "" = "" Then
				    strEmail(0, i) = rs(0).Value
				    strEmail(1, i) = rs(1).Value & "; "
                    strEmail(2, i) = "Staff"
				    i = i + 1
				    ReDim Preserve strEmail(2, i)
				End If
			End If
		End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs =Nothing

    'get team contact names for the email
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamContactsID, Email FROM TeamContacts WHERE TeamsID IN (" & sTeamsID & ") ORDER BY ContactName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
		If Request.Form.Item("contact_" & rs(0).Value) = "on" Then
			If Request.Form.Item("c_e_" & rs(0).Value) = "y" Then
				If Not rs(1).Value & "" = "" Then
				    strEmail(0, i) = rs(0).Value
				    strEmail(1, i) = rs(1).Value & "; "
                    strEmail(2, i) = "Contact"
				    i = i + 1
				    ReDim Preserve strEmail(2, i)
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
		    SendTo(0, i) = strEmail(0, i)
		    SendTo(1, i) = strEmail(1, i)
            SendTo(2, i) = strEmail(2, i)
		    k = k + 1
		    ReDim Preserve SendTo(2, k)
        Else
		    BadEmail(j) = strEmail(1, i)
		    j = j + 1
		    ReDim Preserve BadEmail(j)
	    End If
    Next

    If Not sFiles = vbNullString Then 
        Files = Split(sFiles, ", " )
    Else
        ReDim Files(0)
    End If

    %>
    <!--#include file = "../../../../includes/cdo_connect.asp" -->
    <%

    For j = 0 To UBound(SendTo, 2) - 1
        'insert first recipient
        sql = "INSERT INTO MsgRecips (MsgLogID, Recipient, Email, Role) VALUES (" & lMsgLogID & ", " & SendTo(0, j) & ", '" & SendTo(1, j) & "', '" 
        sql = sql & SendTo(2, j) & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        Set cdoMessage = CreateObject("CDO.Message")
        With cdoMessage
	        Set .Configuration = cdoConfig
            .To = SendTo(1, j)
		    .From = sEmail
            .Subject = sSubject
	        .TextBody = sMsg
	        .Send
        End With
        Set cdoMessage = Nothing
    Next

    Set cdoConfig = Nothing
End If

'first get participant emails
bFound = False
i=0
ReDim EmailParts(1,0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RosterID, FirstName, LastName FROM Roster WHERE TeamsID IN (" & sTeamIDs & ") AND Email <> '' AND Archive = 'n' ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    For j = 0 To UBound(Recipients, 2) - 1
        If CLng(Recipients(1, j)) = CLng(rs(0).Value) Then
            bFound = True
            Exit For
        End If
    Next

    If bFound = False Then
	    EmailParts(0, i) = rs(0).Value
	    EmailParts(1, i) = rs(1).Value & " " & rs(2).Value
	    i = i+1
	    ReDim Preserve EmailParts(1, i)
    Else
        bFound = False
    End If
	rs.MoveNext
Loop
Set rs=Nothing

'get staff members
bFound = False
i=0
ReDim EmailStaff(1,0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamStaffID, FirstName, LastName FROM TeamStaff WHERE CoachesID = " & lCoachID & " AND Email <> '' ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.eof
    For j = 0 To UBound(Recipients, 2) - 1
        If CLng(Recipients(1, j)) = CLng(rs(0).Value) Then
            bFound = True
            Exit For
        End If
    Next

    If bFound = False Then
	    EmailStaff(0, i) = rs(0).Value
	    EmailStaff(1, i) = rs(1).Value & " " & rs(2).Value
	    i = i+1
	    ReDim Preserve EmailStaff(1, i)
    Else
        bFound = False
    End If
	rs.MoveNext
Loop
Set rs=Nothing

'get team contacts
bFound = False
i=0
ReDim Contacts(1,0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamContactsID, ContactName FROM TeamContacts WHERE TeamsID IN (" & sTeamIDs & ") AND Email <> '' ORDER BY ContactName"
rs.Open sql, conn, 1, 2
Do While Not rs.eof
    For j = 0 To UBound(Recipients, 2) - 1
        If CLng(Recipients(1, j)) = CLng(rs(0).Value) Then
            bFound = True
            Exit For
        End If
    Next

    If bFound = False Then
	    Contacts(0, i) = rs(0).Value
	    Contacts(1, i) = Replace(rs(1).Value, "''", "'")
	    i = i+1
	    ReDim Preserve Contacts(1, i)
    Else
        bFound = False
    End If
	rs.MoveNext
Loop
Set rs=Nothing

Private Function GetRecipName(lRecipID, sRecipRole)
    Set rs = Server.CreateObject("ADODB.Recordset")
    Select Case sRecipRole
        Case "Participant"
            sql = "SELECT LastName, FirstName FROM Roster WHERE RosterID = " & lRecipID
            rs.Open sql, conn, 1, 2
            GetRecipName = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        Case "Team Staff"
            sql = "SELECT LastName, FirstName FROM TeamStaff WHERE TeamStaffID = " & lRecipID
            rs.Open sql, conn, 1, 2
            GetRecipName = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        Case "Team Contact"
            sql = "SELECT ContactName FROM TeamContacts WHERE TeamContactsID = " & lRecipID
            rs.Open sql, conn, 1, 2
            GetRecipName = Replace(rs(0).Value, "''", "'")
        Case Else
            sql = "SELECT LastName, FirstName FROM Coaches WHERE CoachesID = " & lRecipID
            rs.Open sql, conn, 1, 2
            GetRecipName = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    End Select
    rs.Close
    Set rs = Nothing
End Function

Private Function GetSenderName(lThisSender, sThisRole)
    Set rs = Server.CreateObject("ADODB.Recordset")
    Select Case sThisRole
        Case "coach"
            sql = "SELECT LastName, FirstName FROM Coaches WHERE CoachesID = " & lThisSender
            rs.Open sql, conn, 1, 2
            GetSenderName = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        Case Else
            sql = "SELECT ContactName FROM TeamContacts WHERE TeamContactsID = " & lThisSender
            rs.Open sql, conn, 1, 2
            GetSenderName = Replace(rs(0).Value, "''", "'")
    End Select
    rs.Close
    Set rs = Nothing
End Function

%>
<!--#include file = "../../../../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>Gopher State Events&reg; Group Email Log Details</title>
<!--#include file = "../../../../includes/js.asp" --> 
</head>

<body>
<div class="container">
	<h4 style="width: 935px;">Gopher State Events<sup>&reg;</sup> Group Email Message Details</h4>	

	<form name="re-send" method="post" action="log_details.asp?msg_log_id=<%=lMsgLogID%>">
    <div class="col-sm-4">
        <h4 class="h4">Available Recipients</h4>
		<%If UBound(EmailStaff, 2) > 0 Then%>
			<h5 class="h5">Support Staff:</h5>
			<ul class="list-group">
				<%If UBound(EmailStaff, 2) > 1 Then%>
					<li class="list-group-item"><input type="checkbox" name="all_staff" id="all_staff" title="all_staff">All Staff</li>
				<%End If%>
				<%For i = 0 to UBound(EmailStaff, 2) - 1%>
					<li class="list-group-item">
						<input type="hidden" name="s_e_<%=EmailStaff(0, i)%>" id="s_e_<%=EmailStaff(0, i)%>" value="y">
						<input type="checkbox" name="staff_<%=EmailStaff(0, i)%>" id="staff_<%=EmailStaff(0, i)%>"><%=EmailStaff(1, i)%>
					</li>
				<%Next%>
			</ul>
		<%End If%>

		<%If UBound(Contacts, 2) > 0 Then%>
			<h5 class="h5">Team Contacts:</h5>
			<ul class="list-group">
				<%If UBound(Contacts, 2) > 1 Then%>
					<li class="list-group-item"><input type="checkbox" name="all_contacts" id="all_contacts" title="all_contacts"> All Contacts</li>
				<%End If%>
				<%For i = 0 to UBound(Contacts, 2) - 1%>
					<li class="list-group-item">
						<input type="hidden" name="c_e_<%=Contacts(0, i)%>" id="c_e_<%=Contacts(0, i)%>" value="y">
						<input type="checkbox" name="contact_<%=Contacts(0, i)%>" id="contact_<%=Contacts(0, i)%>"><%=Contacts(1, i)%>
					</li>
				<%Next%>
			</ul>
		<%End If%>
    </div>
    <div class="col-sm-4">
        <h5 class="h5">Previous Recipients</h5>

		<%If UBound(EmailParts, 2) > 0 Then%>
			<label>Participants:</label>
			<ul class="list-group">
				<%If UBound(EmailParts, 2) > 1 Then%>
					<li class="list-group-item"><input type="checkbox" name="all_parts" id="all_parts" title="all_parts"> All Participants</li>
				<%End If%>
				<%For i = 0 to UBound(EmailParts, 2) - 1%>
					<li class="list-group-item">
						<input type="hidden" name="p_e_<%=EmailParts(0, i)%>" id="p_e_<%=EmailParts(0, i)%>" value="y">
						<input type="checkbox" name="part_<%=EmailParts(0, i)%>" id="part_<%=EmailParts(0, i)%>"><%=EmailParts(1, i)%>
					</li>
				<%Next%>
			</ul>
		<%End If%>

		<%If UBound(Recipients, 2) > 0 Then%>
            <label>Team Staff:</label>
			<ul class="list-group">
				<%For i = 0 to UBound(Recipients, 2) - 1%>
					<%If Recipients(2, i) = "Team Staff" Then%>
                        <li class="list-group-item"><%=GetRecipName(Recipients(1, i), Recipients(2, i))%></li>
                    <%End If%>
				<%Next%>
			</ul>

            <label>Participants:</label>
			<ul>
				<%For i = 0 to UBound(Recipients, 2) - 1%>
					<%If Recipients(2, i) = "Participant" Then%>
                        <li class="list-group-item"><%=GetRecipName(Recipients(1, i), Recipients(2, i))%></li>
                    <%End If%>
				<%Next%>
			</ul>

            <label>Team Contacts:</label>
			<ul>
				<%For i = 0 to UBound(Recipients, 2) - 1%>
					<%If Recipients(2, i) = "Team Contact" Then%>
                        <li><%=GetRecipName(Recipients(1, i), Recipients(2, i))%></li>
                    <%End If%>
				<%Next%>
			</ul>
		<%End If%>
    </div>
    <div class="col-sm-4">
		<table class="table">
			<tr>
				<th>Sender:</th>
				<td><a href="mailto:<%=sEmail%>"><%=GetSenderName(lSender, sRole)%> (<%=sRole%>)</a></td>
			</tr>
			<tr>
				<th>Subject:</th>
				<td><%=sSubject%></td>
			</tr>
			<tr>
				<th>When Sent:</th>
				<td><%=dWhenSent%></td>
            </tr>
			<tr>
				<th>Msg:</th>
				<td><%=sMsg%></td>
			</tr>
			<tr>
				<td>
                    <input type="hidden" name="submit_this" id="submit_this" value="submit_this">
                    <input type="submit" class="form-control" name="submit1" id="submit1" value="Re-Send to Selected">
                </td>
			</tr>
		</table>
    </div>
	</form>
</div>
<%
conn.Close
Set conn=Nothing
%>
</body>
</html>
