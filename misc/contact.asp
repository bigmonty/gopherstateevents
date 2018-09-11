<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sSubject, sYourName, sEmail, sBody, sMsg
Dim cdoMessage, cdoConfig
Dim bMsgSent

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%
    
If Request.form.Item("submit_this") = "submit_this" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'contact_vira' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_contact_vira") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_contact_vira") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
		
		sYourName = CleanInput(Trim(Request.Form.Item("your_name")))
		If sHackMsg = vbNullString Then sEmail = CleanInput(Trim(Request.Form.Item("email")))
		If sHackMsg = vbNullString Then sSubject = CleanInput(Trim(Request.Form.Item("subject")))
		If sHackMsg = vbNullString Then sBody = CleanInput(Trim(Request.Form.Item("body")))
		
		If sHackMsg = vbNullString Then
			sMsg = "A Message from GSE Contact Us" & vbCrLf & vbCrLf
			sMsg = sMsg & "From: " & sYourName & vbCrLf
			sMsg = sMsg & "Email Address: " & sEmail & vbCrLf & vbCrLf
			sMsg = sMsg & "Message: " &vbCrLf 
			sMsg = sMsg & sBody

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = "bob.schneider@gopherstateevents.com"
				.From = "" & sEmail & "<" & sEmail & ">"
				.Subject = sSubject
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
			
			bMsgSent = True
			
            If Not sBody = vbNullString Then sBody = Replace(sBody, "'", "''")

			sql = "INSERT INTO ContactLog (FromName, FromEmail, Subject, Message, WhenSent) VALUES ('" & sYourName
			sql = sql & "', '" & sEmail & "', '" & sSubject & "', '" & Left(sBody, 2000) & "', '" & Now() & "')"
			Set rs = conn.Execute(sql)
			Set rs = Nothing
		End If
	End If
End If

'log this user if they are just entering the site
If Session("access_contact_vira") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'contact_vira')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'contact'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_contact_vira")
End If

%>
<!--#include file = "../includes/clean_input.asp" -->
<%
    
Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; Contact Us</title>
<meta name="description" content="Contact Gopher State Events (GSE), a conventional timing service for fitness events, cross-country, and nordic skiing offered by H51 Software, LLC in Minnetonka, MN.">
<!--#include file = "../includes/js.asp" -->

<script>
function chkFlds() {
if (document.contact_us.your_name.value == '' ||
    document.contact_us.subject.value == '' ||
    document.contact_us.body.value == '' ||
    document.contact_us.email.value == '') 
{
 	alert('All fields are required!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body onload="javascript:contact_us.your_name.focus()">
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-md-10">
			<h1 class="h1">Contact Gopher State Events</h1>

            <div class="col-md-6">
				<h4 class="h4">Gopher State Events, LLC</h4>
                <div>
				    601 Carlson Pkwy<br>
				    Suite 1050<br>
				    Minnetonka, MN 55305<br>
				    <a href="tel:+1-612-720-8427">612-720-8427</a>
                </div>
                <iframe class="embed-responsive-item" src="http://www.mapquest.com/embed?hk=1g30zQm" 
                marginwidth="0" marginheight="0" frameborder="0" scrolling="no"></iframe>						
            </div>
            <div class="col-md-6">
			    <%If bMsgSent = True Then%>
				    <p>Your email has been sent and it will be responded to in less than 24 hours.  Thank you for your interest in GSE!</p>
			    <%Else%>
				    <form class="form-horizontal" name="contact_us" method="post" action="contact.asp" onSubmit="return chkFlds();">
				    <div class="form-group">
					    <label for="your_name" class="control-label col-xs-4">Your Name:</label>
				        <div class="col-xs-8">
                            <input class="form-control" type="text" name="your_name" id="your_name" maxlength="50">
                        </div>
				    </div>
				    <div class="form-group">
					    <label for="email" class="control-label col-xs-4">Email Address:</label>
				        <div class="col-xs-8">
                            <input class="form-control" type="text" name="email" id="email" maxlength="50">
                        </div>
				    </div>
				    <div class="form-group">
					    <label for="subject" class="control-label col-xs-4">Subject:</label>
				        <div class="col-xs-8">
                            <input class="form-control" type="text" name="subject" id="subject" maxlength="50">
                        </div>
				    </div>
                    <div class="form-group">
                        <label for="body" class="control-label col-xs-4">Message:</label>
				        <div class="col-xs-8">
                            <textarea class="form-control" name="body" id="body" rows="10"></textarea>
                        </div>
                    </div>
				    <div class="form-group">
					    <input class="form-control" type="hidden" name="submit_this" id="submit_this" value="submit_this">
					    <input class="form-control" type="submit" name="submit1" id="submit1" value="Send">
				    </div>
				    </form>
			    <%End If%>
            </div>
  		</div>
	</div>
	<!--#include file = "../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
