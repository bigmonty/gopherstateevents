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
<title>About GSE&copy;</title>
<meta name="description" content="Gopher State Events (GSE) is a chip timing service for fitness events, cross-country, and nordic skiing in Minnetonka, MN.">
<!--#include file = "../includes/js.asp" --> 

<style type="text/css">
	.page_list{
		margin-left:10px;
	}
	
	.page_list li{
		margin-top:5px;
		padding:0;
		font-size:0.85em;	
	}
	
	p{
		font-size:0.9em;
	}
	
	h5{
		margin:15px 0 0 0;
		padding:0;
	}
</style>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-md-10">
			<h1 class="h1">About GSE (Gopher State Events)</h1>
			
			<iframe style="height: 200px; width: 340px;float: left;margin: 10px;" src="http://www.mapquest.com/embed?hk=1g30zQm" 
                            marginwidth="0" marginheight="0" frameborder="0" scrolling="no"></iframe>
			
			<p>Gopher State Events, LLC is a combination web-service/on-site results processing/event management utility for fitness 
            events, mountain bike, cross-country, and nordic ski events as well as duathlons, triathlons, and distance swimming and specialty events.  We 
            are located in Minnetonka, MN and use state-of-the-art disposable and permanent rfid (chip) technology.</p>

            <p>GSE processes results on site, posts results online and sends email results as the event progresses, posts finish line pictures
            and videos to the web, and is willing to customize our services to accomodate the needs of our customers.  Feel free to 
            <a href="../contact_us/contact_us.asp">contact us</a> directly for a more detailed explanation of our services and pricing.</p>

            <br>

            <h3 class="h3">The GSE Mission</h3>
            <p>
                To make fitness event management, including related school sports, a more enjoyable, healthy, and social experience for the participants,
                to support the worthy causes that these events contribute to, to make the work easier for the event director, and to do this at a fee that 
                is affordable and margin-based, rather than driven by what we think the market will bear.
            </p>
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
