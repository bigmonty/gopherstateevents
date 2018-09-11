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
<title>About GSE's&copy; Staff</title>
<meta name="description" content="About Gopher State Events (GSE) staff">
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
			<h1 class="h1">About GSE Staff</h1>

			<p>Gopher State Events, LLC was founded by Bob Schneider and Bob Bakken ("the Bobs") in late 2012.  GSE "assumed" Schneider's old company,
            Virtual Race Assistant, tapped into Bakken's logistics and equipment expertise, and they were off-n-running (pun intended).</p>
				
			<p><img src="/graphics/bob_boston.jpg" alt="Bob Schneider" width="100" style="float:left;margin:10px;">Bob Schneider is a lifelong runner and educator.  He currently teaches math at Edina (MN USA) High School.  
			While he no longer runs due to a knee injury suffered coaching high school basketball, running has been one of the 
			defining characteristics of his life.</p>
				
			<p><img src="/graphics/staff/bob_s.jpg" alt="Bob Schneider-GSE" width="150" style="float:right;margin:10px;">Schneider is the "IT Guy".  He
            has developed a software business (online at <a href="http://www.h51software.net">http://www.h51software.net</a>).  H51 creates 
            educational software and services.  He has also developed an online service for managing high school, college/university, and club cross-country and track & field teams
			(called <a href="http://www.etraxc.com" onclick="openThis(this.href,1024,768);return false;">eTRaXC</a>).  eTRaXC is a comprehensive 
			online coach-directed running team manager.</p>

            <p>GSE makes it a point to hire timers that are committed and loyal professionals who are committed to their work.  Many are teachers and
            school administrators.  We believe it is important to put your race in the hands of people who understand the importance of your work and
            theirs.</p>

            <br>

            <h4 class="h4">Our Crew!</h4>
            <table style="margin-left: 20px;">
                <tr>
                    <td style="padding: 10px;text-align: center;" valign="top"><img src="/graphics/staff/arne.jpg" alt="Arne" width="230"><br><span style="font-weight:bold;">Arne</span></td>
                    <td style="padding: 10px;text-align: center;" valign="top"><img src="/graphics/staff/jim_kurt.jpg" alt="Kurt & Jim" width="230"><br><span style="font-weight:bold;">Kurt & Jim</span></td>
                    <td style="padding: 10px;text-align: center;" valign="top"><img src="/graphics/staff/karen.jpg" alt="Karen" width="230"><br><span style="font-weight:bold;">Karen</span></td>
                </tr>
            </table>
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
