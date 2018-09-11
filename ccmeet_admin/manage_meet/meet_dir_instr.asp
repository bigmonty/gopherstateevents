<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisMeet, sMeetName, dMeetDate, sFirstName, sLastName, sEmail, sUserID, sPassword
Dim i, j
Dim cdoMessage, cdoConfig
Dim sMsg
Dim lMeetDirID
Dim bAlreadySent, dDateSent

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet data
sql = "SELECT md.MeetDirID, md.FirstName, md.LastName, md.Email, md.UserID, md.Password, m.MeetName, m.MeetDate "
sql = sql & "FROM MeetDir md INNER JOIN Meets m ON md.MeetDirID = m.MeetDirID WHERE m.MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
lMeetDirID = rs(0).Value
sFirstName = Replace(rs(1).Value, "''", "'")
sLastName = Replace(rs(2).Value, "''", "'")
sEmail = rs(3).Value
sUserID = rs(4).Value
sPassword = rs(5).Value
sMeetName = Replace(rs(6).Value, "''", "'")
dMeetDate = rs(7).Value
Set rs = Nothing

If Request.Form.Item("send_instr") = "send_instr" Then
	sMsg = vbCrLf
	sMsg = sMsg & "Dear Meet Director " & sLastName & ": " & vbCrLf & vbCrLf
	
	sMsg = sMsg & "You are receiving this email because you are hosting the " & sMeetName 
	sMsg = sMsg & " on " & dMeetDate & " which you have contracted GSE (Gopher State Events) to score.  Attached "
	sMsg = sMsg & "to this email you will find a sketch of what we need from you to make your meet  a success for you, "
	sMsg = sMsg & "us, and most importantly, the teams that are participating. " & vbCrLf & vbCrLf
	
	sMsg = sMsg & "When you have GSE manage your meet your workload diminishes dramatically but there are a couple "
	sMsg = sMsg & "of things that are very important for you to know and understand.  Please read the attachment carefully "
	sMSg = sMSg & "and contact us if you have any questions. " & vbCrLf & vbCrLf
			
	sMsg = sMsg & "An account has been created for you on our site (www.gopherstateevents.com) to help you manage this meet.  Please "
	sMSg = sMSg & "log on to the site at your convenience and take a look around.  Your login information is listed below.  Note "
	sMsg = sMsg & "that if you are also a coach that is using the system, your login information is the same for both roles.  The "
	sMsg = sMsg & "difference is that to log in as a meet director you will enter your information in the 'For Meet Directors Only' "
	sMsg = sMsg & "area on the home page. "& vbCrLf & vbCrLf
			
	sMsg = sMsg & "User ID: " & sUserID & vbCrLf
	sMsg = sMsg & "Password: " & sPassword & vbCrLf & vbCrLf
			
	sMsg = sMsg & "Thank you in advance for looking over the process detailed on the attachnent to ensure that all is ready "
	sMsg = sMsg &  "to go on race day.  You may call or email me for assistance regarding this process. " & vbCrLf & vbCrLf
	
	sMsg = sMsg & "Sincerely ;" & vbCrLf & vbCrLf
	sMsg = sMSg & "Bob Schneider " & vbCrLf
	sMsg = sMSg & "GSE (Gopher State Events)  " & vbCrLf
	sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
	sMsg = sMsg & "612-720-8427 " & vbCrLf
	
	'write these to db
	sql = "INSERT INTO MeetDirInstr (MeetDirID, DateSent) VALUES (" & lMeetDirID & ", '" & Date & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
			 
	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
'		.To = "bob.schneider@gopherstateevents.com"
		.To = sEmail
		.BCC = "bob.schneider@gopherstateevents.com"
		.From = "bob.schneider@gopherstateevents.com"
		.AddAttachment "c:\inetpub\h51web\gopherstateevents\admin\ccmeet_admin\grp_email\attachments\meet_dir_instr.doc"
		.Subject = "GSE Meet Director Instructions"
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
End If

'identify which teams have had this information sent to them this year
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DateSent FROM MeetDirInstr WHERE MeetDirID = " & lMeetDirID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then 	
	bAlreadySent = True
	dDateSent = rs(0).Value
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country Meet Director Instructions Sheet</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    

			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>
			
			<h4 class="h4">CCMeet Meet Dir Instructions:&nbsp;<%=sMeetName%></h4>
			
			<form class="form-inline" name="send_instr" method="post" action="meet_dir_instr.asp?meet_id=<%=lThisMeet%>">
			<%=sFirstName%>&nbsp;<%=sLastName%>&nbsp;&nbsp;&nbsp;&nbsp;
			<%If Not dDateSent = vbNullString Then%>
				<%=dDateSent%>&nbsp;&nbsp;&nbsp;&nbsp;
				<%bAlreadySent = True%>
			<%End If%>
			<input type="checkbox" name="send_to" id="send_to">&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="hidden" name="send_instr" id="send_instr" value="send_instr">
			<input type="submit" name="submit" id="submit" value="Send Instructions">
			</form>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
