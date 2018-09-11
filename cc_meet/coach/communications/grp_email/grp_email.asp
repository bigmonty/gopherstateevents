<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lCoachID
Dim EmailParts(), EmailStaff(), Contacts()
Dim sTeamIDs

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

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

'first get participant emails
i=0
ReDim EmailParts(1,0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RosterID, FirstName, LastName FROM Roster WHERE TeamsID IN (" & sTeamIDs & ") AND Archive = 'n' AND Email <> '' ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.eof
	EmailParts(0, i) = rs(0).Value
	EmailParts(1, i) = rs(1).Value & " " & rs(2).Value
	i = i+1
	ReDim Preserve EmailParts(1, i)
	rs.MoveNext
Loop
Set rs=Nothing

'get staff members
i=0
ReDim EmailStaff(1,0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamStaffID, FirstName, LastName FROM TeamStaff WHERE CoachesID = " & lCoachID & " AND Email <> '' ORDER BY LastName, FirstName" 
rs.Open sql, conn, 1, 2
Do While Not rs.eof
	EmailStaff(0, i) = rs(0).Value
	EmailStaff(1, i) = rs(1).Value & " " & rs(2).Value
	i = i+1
	ReDim Preserve EmailStaff(1, i)
	rs.MoveNext
Loop
Set rs=Nothing

'get team contacts
i=0
ReDim Contacts(1,0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamContactsID, ContactName FROM TeamContacts WHERE TeamsID IN (" & sTeamIDs & ") AND Email <> '' ORDER BY ContactName"
rs.Open sql, conn, 1, 2
Do While Not rs.eof
	Contacts(0, i) = rs(0).Value
	Contacts(1, i) = Replace(rs(1).Value, "''", "'")
	i = i+1
	ReDim Preserve Contacts(1, i)
	rs.MoveNext
Loop
Set rs=Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>Gopher State Events&reg; Cross-Country/Nordic Ski Group Email</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<!--#include file = "communications_nav.asp" -->
			<h4 class="h4">Gopher State Events<sup>&reg;</sup> Cross-Country/Nordic Ski Group Email Feature</h4>	

			<p>
				NOTE:  If an error is generated, use the "Back" key to return to the page but DO NOT resend the message from that screen.
				Open the <a href="email_log.asp" style="font-weight: bold;">Message Log</a> page and work from that screen.
			</p>

			<form role="form" class="form-horizontal" name="grp_email" method="post" action="send_email.asp">
			<div class="row">
				<div class="col-sm-3">
					<%If UBound(EmailStaff, 2) > 0 Then%>
						<h5 class="h5">Support Staff:</h5>
						<ul class="list-group">
							<%If UBound(EmailStaff, 2) > 1 Then%>
								<li class="list-group-item">
									<input type="checkbox" name="all_staff" id="all_staff" title="all_staff">
									<span style="font-weight: bold;">All Staff</span>
								</li>
							<%End If%>
							<%For i = 0 to UBound(EmailStaff, 2) - 1%>
								<li class="list-group-item">
									<input type="hidden" name="s_e_<%=EmailStaff(0, i)%>" id="s_e_<%=EmailStaff(0, i)%>" value="y">
									<input type="checkbox" name="staff_<%=EmailStaff(0, i)%>" id="staff_<%=EmailStaff(0, i)%>"><%=EmailStaff(1, i)%>
								</li>
							<%Next%>
						</ul>
					<%Else%>
						<p>None of your support staff have an email address on file.</p>
					<%End If%>

					<%If UBound(Contacts, 2) > 0 Then%>
						<h4 class="h4">Team Contacts:</h4>
						<ul class="list-group">
							<%If UBound(Contacts, 2) > 1 Then%>
								<li class="list-group-item">
									<input type="checkbox" name="all_contacts" id="all_contacts" title="all_contacts">
									<span style="font-weight: bold;">All Contacts</span><br><hr style="margin:5px 0 5px 0;">
								</li>
							<%End If%>
							<%For i = 0 to UBound(Contacts, 2) - 1%>
								<li class="list-group-item">
									<input type="hidden" name="c_e_<%=Contacts(0, i)%>" id="c_e_<%=Contacts(0, i)%>" value="y">
									<input type="checkbox" name="contact_<%=Contacts(0, i)%>" id="contact_<%=Contacts(0, i)%>"><%=Contacts(1, i)%>
								</li>
							<%Next%>
						</ul>
					<%Else%>
						<p>None of your team contacts have an email address on file.</p>
					<%End If%>
				</div>
				<div class="col-sm-3">
					<%If UBound(EmailParts, 2) > 0 Then%>
						<h4 class="h4">Participants:</h4>
						<ul class="list-group">
							<%If UBound(EmailParts, 2) > 1 Then%>
								<li class="list-group-item">
									<input type="checkbox" name="all_parts" id="all_parts" title="all_parts">
									<span style="font-weight: bold;">All Participants</span><br><hr style="margin:5px 0 5px 0;">
								</li>
							<%End If%>
							<%For i = 0 to UBound(EmailParts, 2) - 1%>
								<li class="list-group-item">
									<input type="hidden" name="p_e_<%=EmailParts(0, i)%>" id="p_e_<%=EmailParts(0, i)%>" value="y">
									<input type="checkbox" name="part_<%=EmailParts(0, i)%>" id="part_<%=EmailParts(0, i)%>"><%=EmailParts(1, i)%>
								</li>
							<%Next%>
						</ul>
					<%Else%>
						<p>None of your team participants have an email address on file.</p>
					<%End If%>
				</div>
				<div class="col-sm-6">
					<div class="form-group row">
						<label for="subject" class="control-label col-sm-4">Subject:</label>
						<div class="col-sm-8">
							<input type="text" class="form-control" name="subject" id="subject">
						</div>
					</div>
					<div class="form-group row">
						<label for="message" class="control-label col-sm-4">Msg:</label>
						<div class="col-sm-8">
							<textarea class="form-control" name="message" id="message" rows="18"></textarea>
						</div>
					</div>
					<div class="form-group">
						<input type="submit" class="form-control" name="submit" id="submit" value="Send Email">
					</div>
				</div>					
				</form>
			</div>
		</div>
	</div>
</div>
<!--#include file = "../../../../includes/footer.asp" --> 
<%
conn.Close
Set conn=Nothing
%>
</body>
</html>
