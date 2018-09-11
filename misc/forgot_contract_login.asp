<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim sEmail
Dim sPageStatus
Dim sUserID, sPassword
Dim sMsg
Dim cdoMessage, cdoConfig

If Request.Form.Item("get_login") = "get_login" Then
	sEmail = Request.Form.Item("email")
	
	Response.Buffer = True		'Turn buffering on
	Response.Expires = -1		'Page expires immediately
												
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

    sPageStatus = "not_found"

	sql = "SELECT UserID, Password FROM EventDir WHERE Email = '" & sEmail & "'"
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
		sMsg = sMsg & "Hangar51 Software/GSE/CCLog"
		
		'send login
%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = sEmail
			.BCC = "bob.schneider@gopherstateevents.com"
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
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>GSE Forgot Contract Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="Forgot contract page login information for GSE (Gopher State Events).">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">




<style type="text/css">
<!--
p, input{
	font-size:0.85em;
	}
-->
</style>

<script>
function chkFlds() {
if (document.login.email.value == '') 
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
<div style="margin:10px;background-color:#fff;padding: 5px;background-image: url('/graphics/gse_ad.jpg');background-repeat: no-repeat;">
	<hr style="margin:135px 10px 0 10px;">

	<h1 style="margin:0 10px 10px 10px;padding:5px;font-size:1.1em;background-color: #ececd8;">Forgot GSE Sign-in</h1>
	
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
			exists in our database.  If it does not exist, you will receive a message indicating that.</p>  
			
			<p>If no email address matches the one you give us you will need to contact us via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> 
            or telephone at 612-720-8427 and your login will be given to you after verifying your identity.  If you request your login 
            information via email, please include your name, your event/meet and your school as it pertains to your account.  Also, 
            please indicate the role you are requesting an account for.</p>

			<form name="login" method="post" action="forgot_contract_login.asp" onSubmit="return chkFlds();">
			<span style="font-weight:bold;font-size:0.9em;">Email address:</span>
			<input type="text" name="email" id="email" size="30">
            <br>
			<input type="hidden" name="get_login" id="get_login" value="get_login">
			<input type="submit" name="submit" id="submit" value="Get Login">
			</form>                            			
	<%End Select%>
</div>
</body>
</html>
