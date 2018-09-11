<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; Privacy Policy</title>
<meta name="description" content="About Gopher State Events (GSE) timing service for fitness events, cross-country, and nordic skiing in Minnetonka, MN.">
<!--#include file = "../includes/js.asp" --> 
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-md-10">
			<h1 class="h1">Gopher State Events Privacy Policy</h1>
	
	        <p>Your privacy is very important to us. This privacy policy describes the information we gather and what we do with that information.</p>

	        <ul style="margin-left:15px;">
		        <li>We will never sell, trade, loan or give your information (including email addresses) to anyone else. </li>
		        <li>We collect only information needed to provide the fitness management, timing, and event management services that we are contracted to  provide</li>
		        <li>By our choice, we do not have access to any of your personal financial information.  Any of our services that require online
		        payment processing are done through a secure third-party.</li>
		        <li>We will not subject our users to pop-up advertising or any advertising that detracts from the enjoyable and efficient use of our
		        sites</li>
		        <li>You will NEVER receive any outside advertising via e-mail resulting from your use of this site or races we manage.</li>
		        <li>This site will never install any tracking utilities, spyware, or malware on your system.</li>
                <li>We may occassionally send you emails (your race results, pre-race informational emails, information on races that we have in-house
                that we think you may be interested in, etc.  You can opt out of these emails very easily.</li>
	        </ul>
        </div>
    </div>
	<!--#include file = "../includes/footer.asp" -->
</div>
<%
		
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
