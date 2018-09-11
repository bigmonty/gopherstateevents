<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lMyProvider, lTeamID, lCoachID
Dim sMyNumber, sErrMsg, sTeamIDs
Dim TextParts(), TextStaff(), TextContacts()

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

'get my wireless
Set rs=Server.CreateObject("ADODB.Recordset")
If Session("role") = "coach" Then
	sql = "SELECT CellProvidersID, CellPhone FROM Coaches WHERE CoachesID = " & Session("my_id")
Else
    sql = "SELECT Provider, MobilePhone FROM TeamStaff WHERE TeamStaffID = " & Session("my_id")
End If
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
	lMyProvider = rs(0).Value
	sMyNumber = rs(1).Value
End If
rs.Close
Set rs = Nothing

If CStr(lMyProvider) & "" = "" Or CStr(sMyNumber) & "" = "" Then
	sErrMsg = "I am sorry.  You may not send a text message until your wireless information is in our database.  You may enter this information on "
	sErrmsg = sErrMsg & "your profile page."
End If

If sErrMsg = vbNullString Then
	'get participant text info
	i=0
	ReDim TextParts(1,0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RosterID, FirstName, LastName, CellProvidersID, CellPhone FROM Roster WHERE TeamsID IN (" & sTeamIDs  & ") AND Archive = 'n' "
    sql = sql & "ORDER BY LastName, FirstName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.eof
		If CLng(rs(3).Value) > 0 Then
            If Not Trim(rs(4).Value) & "" = "" Then
                TextParts(0, i) = rs(0).Value
		        TextParts(1, i) = Replace(rs(1).Value, "''","") & " " & Replace(rs(2).Value, "''","")
		        i = i+1
		        ReDim Preserve TextParts(1, i)
            End If
        End If
		rs.MoveNext
	Loop
	Set rs=Nothing
	
	'get staff
	i=0
	ReDim TextStaff(1,0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT TeamStaffID, FirstName, LastName, Provider, MobilePhone FROM TeamStaff WHERE CoachesID = " & lCoachID  & " ORDER BY LastName, FirstName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.eof
		If CLng(rs(3).Value) > 0 Then
            If Not Trim(rs(4).Value) & "" = "" Then
                TextStaff(0, i) = rs(0).Value
		        TextStaff(1, i) = Replace(rs(1).Value, "''","") & " " & Replace(rs(2).Value, "''","")
		        i = i+1
		        ReDim Preserve TextStaff(1, i)
            End If
        End If
		rs.MoveNext
	Loop
	Set rs=Nothing
	
	'get team contacts
	i=0
	ReDim TextContacts(1,0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT TeamContactsID, ContactName, CellProvidersID, CellPhone FROM TeamContacts WHERE TeamsID IN (" & sTeamIDs  & ") ORDER BY ContactName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		If CLng(rs(2).Value) > 0 Then
			If Not Trim(rs(3).Value) & "" = "" Then
				TextContacts(0, i) = rs(0).Value
				TextContacts(1, i) = Replace(rs(1).Value, "''", "'")
				i = i + 1
				ReDim Preserve TextContacts(1, i)
			End If
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>Gopher State Events&reg; Text Message Utility</title>

<SCRIPT LANGUAGE="JavaScript">
// function parameters are: field - the string field, count - the field for remaining characters number and max - the maximum number of characters 
function CountLeft(field, count, max) {
// if the length of the string in the input field is greater than the max value, trim it 
if (field.value.length > max)
field.value = field.value.substring(0, max);
else
// calculate the remaining characters 
count.value = max - field.value.length;
}
</script>
</head>

<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<h4 class="h4">GSE Message Utility!</h4>
			<p>(Only those team members with updated phone and provider information are listed here.)</p>
					
			<%If Not Session("role") = "coach" Then%>
				<p class="bg-danger">Note: All text messages will automatically be sent to the team's coach if 
				their mobile device is enabled!</p>
			<%End If%>
					
			<%If sErrMsg = vbNullString Then%>
				<form class="form" name="txt_msg" method="post" action="send_txt.asp">
				<div class="row">
					<div class="col-sm-3">
						<h5 class="h5">Support Staff:</h5>
						<%If UBound(TextStaff, 2) > 0 Then%>
							<ul class="list-group">
								<%If UBound(TextStaff, 2) > 1 Then%>
									<li class="list-group-item"><input type="checkbox" name="all_staff" id="all_staff">All Staff</li>
								<%End If%>
								<%For i = 0 to UBound(TextStaff, 2) - 1%>
									<li class="list-group-item">
										<input type="hidden" name="staff_exists_<%=TextStaff(0, i)%>" id="staff_exists_<%=TextStaff(0, i)%>" value="y">
										<input type="checkbox" name="staff_<%=TextStaff(0, i)%>" id="staff_<%=TextStaff(0, i)%>"><%=TextStaff(1, i)%>
									</li>
								<%Next%>
							</ul>
						<%Else%>
							<p>None of your staff have their provider and/or cell number on file with Gopher State Events.</p>
						<%End If%>
										
						<h5 class="h5">Team Contacts:</h5>
						<%If UBound(TextContacts, 2) > 0 Then%>
							<ul class="list-group">
								<%If UBound(TextContacts, 2) > 1 Then%>
									<li class="list-group-item"><input type="checkbox" name="all_contacts" id="all_contacts">All Contacts</li>
								<%End If%>
								<%For i = 0 to UBound(TextContacts, 2) - 1%>
									<li class="list-group-item">
										<input type="hidden" name="contact_exists_<%=TextContacts(0, i)%>" id="contact_exists_<%=TextContacts(0, i)%>" value="y">
										<input type="checkbox" name="contacts_<%=TextContacts(0, i)%>" id="contacts_<%=TextContacts(0, i)%>"><%=TextContacts(1, i)%>
									</li>
								<%Next%>
							</ul>
						<%Else%>
							<p>None of your team contacts have their provider and/or cell number on file with Gopher State Events.</p>
						<%End If%>
					</div>
					<div class="col-sm-3">
						<h5 class="h5">Participants:</h5>
						<%If UBound(TextParts, 2) > 0 Then%>
							<ul class="list-group">
								<%If UBound(TextParts, 2) > 1 Then%>
									<li class="list-group-item"><input type="checkbox" name="all_parts" id="all_parts"> All Participants</li>
								<%End If%>
								<%For i = 0 to UBound(TextParts, 2) - 1%>
									<li class="list-group-item">
										<input type="hidden" name="part_exists_<%=TextParts(0, i)%>" id="part_exists_<%=TextParts(0, i)%>" value="y">
										<input type="checkbox" name="part_<%=TextParts(0, i)%>" id="part_<%=TextParts(0, i)%>"><%=TextParts(1, i)%>
									</li>
								<%Next%>
							</ul>
						<%Else%>
							<p>None of your team members have their provider and/or cell number on file with Gopher State Events.</p>
						<%End If%>
					</div>
					<div class="col-sm-6">
						<p>
							Mobile providers have character limits.  If your text is too long it will be truncated.  The most common limit is 160 characters.
						</p>
										
						<h5 class="h5">Message:</h5>
						<textarea class="form-control" name="message" id="message" rows="5" onKeyDown="CountLeft(this.form.message,this.form.left,160);" 
							onKeyUp="CountLeft(this.form.message,this.form.left,160);" style="font-size: 1.2em;"></textarea>
						<input readonly type="text" name="left" size="3" maxlength="3" value="160" style="border:none;">characters left
						<input type="submit" class="form-control" name="submit" id="submit" value="Send">
					</div>	
				</div>				
				</form>
			<%Else%>
				<p><%=sErrMsg%></p>
			<%End If%>
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
