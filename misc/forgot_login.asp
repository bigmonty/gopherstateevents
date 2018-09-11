<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim sEmail
Dim sPageStatus
Dim sUserID, sPassword, sRole
Dim sMsg
Dim cdoMessage, cdoConfig

If Request.Form.Item("get_login") = "get_login" Then
	sEmail = Request.Form.Item("email")
    sRole = Request.Form.Item("role")

	sPageStatus = "not_found"
	
	Response.Buffer = True		'Turn buffering on
	Response.Expires = -1		'Page expires immediately
												
	Set conn = Server.CreateObject("ADODB.Connection")
	
	Select Case sRole
		Case "staff"
			conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
			sql = "SELECT UserID, Password FROM Staff WHERE Email = '" & sEmail & "'"
		Case "coach"
			conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
			sql = "SELECT UserID, Password FROM Coaches WHERE Email = '" & sEmail & "'"
		Case "meet_dir"
			conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
			sql = "SELECT UserID, Password FROM MeetDir WHERE Email = '" & sEmail & "'"
		Case "event_dir"
			conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
			sql = "SELECT UserID, Password FROM EventDir WHERE Email = '" & sEmail & "'"
		Case "team_staff"
			conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
			sql = "SELECT UserName, Password FROM TeamStaff WHERE Email = '" & sEmail & "'"
	End Select
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		sUserID = rs(0).value
		sPassword = rs(1).Value
		sPageStatus = "found"
	End If
    rs.Close
	Set rs = Nothing
	
	conn.Close
	Set conn = Nothing
	
	If sPageStatus = "found" Then
		sMsg = vbCrLf
		sMsg = sMsg & "You are receiving this email because a request for login information for your Gopher State Events "
		sMsg = sMsg & "(www.gopherstateevents.com) account was requested sent to this address.  If you did not make this request, please "
		sMsg = sMsg & "notify us immediately at 612.720.8427 or by sending an email to bob.schneider@gopherstateevents.com." & vbCrLf & vbCrLf
		
		sMsg = sMsg & "Here is your login information: " & vbCrLf
		sMsg = sMsg & "Your UserID is: " & sUserID & vbCrLf
		sMsg = sMsg & "Your Password is: " & sPassword & vbCrLf & vbCrLf
		
		sMsg = sMsg & "Sincerely~" & vbCrLf
		sMsg = sMsg & "Bob Schneider" & vbCrLf
		sMsg = sMsg & "Gopher State Events, LLC"
		
		'send login
%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = sEmail
			.CC = "bob.schneider@gopherstateevents.com"
			.From = "bob.schneider@gopherstateevents.com"
			.Subject = "GSE Forgot Login"
			.TextBody = sMsg
			.Send
		End With
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing
	End If
End If
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; Forgot Login</title>
<meta name="description" content="Gopher State Events 'Forgot Login' utility.">
<!--#include file = "../includes/js.asp" -->
<script>
function chkFlds() {
if (document.login.email.value == '' ||
    document.login.role.value == '') 
{
 	alert('Both fields are required!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body>
<div class="container">
	<h3 class="h3">Forgot GSE Sign-in</h3>
	
	<%Select Case sPageStatus%>
		<%Case "found"%>
			<p>The login information has been sent to address requested.  Please notify us via 
			<a href="mailto:bob.schneider@gopherstateevents.com">email</a> if you do not receive it.</p>
		<%Case "not_found"%>
			<p>We are sorry but there was no match in our database for this email address.  Please contact us via 
			<a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by phone at 612.720.8427 to resolve this 
			discrepancy.</p>
		<%Case Else%>
			<p>By clicking the button below, your login information will be sent to the email address your supplied if this email address 
			exists in our database FOR THE ROLE INDICIATED.  If it does not exist, you will receive a message indicating that.</p>  
			
			<p>If no email address matches the one you give us you will need to contact us via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> 
            or telephone at 612-720-8427 and your login will be given to you after verifying your identity.  If you request your login 
            information via email, please include your name, your event/meet and your school as it pertains to your account.  Also, 
            please indicate the role you are requesting an account for.</p>

			<form class="form-inline" name="login" method="post" action="forgot_login.asp" onSubmit="return chkFlds();">
			<label for="email">Email address:</label>
			<input type="text" class="form-control" name="email" id="email">
			<label for="role">GSE Role:</label>
		    <select class="form-control" name="role" id="role">
                <option value="">&nbsp;</option>
                <option value="coach">CC/Nordic Coach</option>
                <option value="staff">GSE Staff</option>
                <option value="meet_dir">Meet Director</option>
                <option value="event_dir">Event Director</option>
				<option value="team_staff">Team Staff</option>
            </select>
			<input type="hidden" name="get_login" id="get_login" value="get_login">
			<input type="submit" class="form-control" name="submit" id="submit" value="Get Login">
			</form>                            			
	<%End Select%>
</div>
</body>
</html>
