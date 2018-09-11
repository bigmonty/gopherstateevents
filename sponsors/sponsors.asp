<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sShowBlurb

sShowBlurb = Request.QueryString("show_blurb")
If sShowBlurb = vbNullString Then sShowBlurb = "y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events Sponsors</title>
<!--#include file = "../includes/js.asp" -->
<style type="text/css">
<!--
    p, input{
	    font-size:0.85em;
	}
    
    #grid {
		margin-top:10px;
		padding-bottom:10px;
	}
    
	#grid div {
		padding:10px;
		float:left;
        width: 170px;
        text-align: center;
	}
	#grid div img{
        width: 170px;
	}
-->
</style>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-md-10">
 			<h1 class="h1">Gopher State Events Partners & Sponsors</h1>
            <%If sShowBlurb = "y" Then%>
                <div style="text-align: right;padding: 0 10px 0 0;margin: 0;">
                    <a href="sponsors.asp?show_blurb=n" style="font-size: 0.8em;">Hide Message</a>
                </div>

                <p>Gopher State Events, LLC is a "for profit" company, but one that is committed to the changing landscape of fitness events, social connectivity, and 
                healthy lifestyles.  We attempt to foster this mission from our school-sponsored events (cross-country running and Nordic skiing)
                through the host of fitness offerings that we are associated with.</p>

                <p>We strive to keep our management prices down and continue to offer services that support our mission.  Free finish line pictures and videos
                are posted shortly after all of our events.  Individual results emails are sent during the event, often within minutes of finishing, 
                posting results online during the event, connecting events through event series, partnering with H51 Software, LLC on eTRaXC/my-eTRaXC, 
                our Performance Tracker, and My GSE History utilities are just some of the ways we are trying to achieve our goals.</p>

                <p>Below are groups, companies, organizations, and other entities that have joined with us in trying make the path to fitness and well-being
                a little easier for all.</p>

                <p>Interested in partnering up with the fastest growing timing companies in the area?  <a href="/misc/vira_contact.asp" 
                    style="font-weight: bold;">Contact Us</a> for more information!</p>
            <%Else%>
                <div style="text-align: right;padding: 0 10px 0 0;margin: 0;">
                    <a href="sponsors.asp?show_blurb=y" style="font-size: 0.8em;">Show Message</a>
                </div>
            <%End If%>

            <div id="grid">
		        <div id="grid-far-left">
                     <a href="http://www.tempoevents.com" onclick="openThis(this.href,1024,768);return false;">
                     <img src="graphics/tempo_banner.png" alt="Tempo Events">
                     </a>
                    <br><br>
                     <a href="http://www.busybodypromotions.com" onclick="openThis(this.href,1024,768);return false;">
                     <img src="/graphics/bbplogo.png" alt="Busy Body Promotions">
                     </a>
                </div>
		        <div id="grid-left">
                     <a href="http://www.etraxc.com" onclick="openThis(this.href,1024,768);return false;">
                     <img src="/sponsors/graphics/etraxc_sponsor.png" alt="eTRaXC">
                     </a>
                    <br><br>
                     <a href="http://www.pro-tree.com" onclick="openThis(this.href,1024,768);return false;">
                     <img src="/graphics/pro-tree.png" alt="Pro-Tree Outdoor Services" style="width:125px;">
                     </a>
                 </div>
		        <div id="grid-middle">
                    <a href="http://www.liquidweb.com/?RID=bobbabuoy" onclick="openThis(this.href,1024,768);return false;"><img src="http://rgfx.liquidweb.com/banners/120x240.jpg" 
                        alt="Liquid Web Fully Managed Web Hosting" border=0></a>
                </div>
		        <div id="grid-right">
					<a href="http://www.roadid.com/Common/default.aspx" onclick="openThis(this.href,1024,768);return false;" style="text-decoration:none;">
						<img src="/graphics/road_id.gif" alt="RoadID"></a>
                    <br><br>
					<a href="http://excelsiorbrew.com/" onclick="openThis(this.href,1024,768);return false;" style="text-decoration:none;">
						<img src="/events/logos/excelsior_brew.png" alt="Excelsior Brewing"></a>
                </div>
		        <div id="grid-far-right">
					<a href="http://threaddesignsinc.com/" onclick="openThis(this.href,1024,768);return false;" style="text-decoration:none;">
						<img src="/graphics/thread_designs.png" alt="Thread Designs, Inc"></a>
                    <br><br>
                     <a href="http://www.rosemountata.com" onclick="openThis(this.href,1024,768);return false;">
                     <img src="/sponsors/graphics/gse_partner_image.png" alt="Rosemount ATA Martial Arts Academy" style="width:125px;">
                     </a>                
                </div>
            </div>
  		</div>
	</div>
	<!--#include file = "../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>