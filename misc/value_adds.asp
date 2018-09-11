<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events Value-Added Services</title>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-sm-10">
 			<h1 class="h1">Gopher State Events Value Added Services</h1>

            <p>
                The following services are offered on an a la carte basis to our event directors.  We offer these services in this manner to keep our prices,
                and your costs, down by not asking event directors to support services that they do not use.  Please 
                <a href="mailto:bob.schneider@gopherstateevents.com">contact us</a> regarding pricing on services you feel might enhance your event.
            </p>

            <div class="list-group">
                <div  class="list-group-item list-group-item-success">
                    <h4 class="h4">Finish Line Pix Sponsorship</h4>
                    <div class="row">
                       <div class="col-sm-4">
                            <img src="/graphics/Camera-icon.png" alt="Finish Line Pix" class="img-responsive">
                        </div>
                        <div class="col-sm-8">
                            Our goal is to photograph every finisher in every race and we put them online later in the day for free download by your 
                            participants.  The image has our logo and yours splashed across the bottom or the top.  For a small fee (about the cost of 
                            a single entry fee) you can replace our logo with any logo you choose.  A nice treat for your main sponsor?
                            <br>
                            <a href="mailto:bob.schneider@gopherstateevents.com">Request Pricing</a>
                        </div>
                    </div>
                </div>
                <div  class="list-group-item list-group-item-danger">
                    <h4 class="h4">Digital Results Kiosk</h4>
                    <div class="row">
                        <div class="col-sm-4">
                            <img src="/graphics/digital_results.png" alt="Digital Results Kiosk" class="img-responsive">
                        </div>
                        <div class="col-sm-8">
                            Yes, we send results out via email and text messaging very quickly...sometimes within seconds.  And yes, we upload them to the
                            web with the same regularity.  But sometimes it is nice to just key in your bib number and get some immediate feedback.  Soon,
                            there will be nothing to key in.  Just walk up to the screen and your results will magically show up (assuming your bib is on
                            the front of your jersey).
                            <br>
                            For an example, enter bib number 3025 into
                            <a href="http://www.gopherstateevents.com/results/fitness_events/digital_results.asp?event_id=667" onclick="openThis(this.href,1024,768);return false;">this page</a>:
                            <br>
                            <a href="mailto:bob.schneider@gopherstateevents.com">Request Pricing</a>
                        </div>
                    </div>
                </div>
                <div  class="list-group-item list-group-item-info">
                    <h4 class="h4">Finish Line Truss</h4>
                    <div class="row">
                       <div class="col-sm-4">
                            <img src="/graphics/finish_truss.jpg" alt="Finish Line Truss" class="img-responsive">
                        </div>
                        <div class="col-sm-8">
                            Add a little grandeur, and some branding, to the finish line of your race.  Use your signage or ours.  We will hang a finish 
                            clock so approaching finishers can see their time as they approach.
                            <br>
                            <a href="mailto:bob.schneider@gopherstateevents.com">Request Pricing</a>
                        </div>
                    </div>
                </div>
                <div  class="list-group-item list-group-item-warning">
                    <h4 class="h4">Announcer</h4>
                    <div class="row">
                        <div class="col-sm-4">
                            <img src="/graphics/fulton_productions.png" alt="Announcers" class="img-responsive">
                        </div>
                        <div class="col-sm-8">
                            Need an announcer?  We've got just the guy.  Kyle Fulton at <a href="http://www.fultonproductions.com/" 
                            onclick="openThis(this.href,1024,768);return false;">Fulton Productions</a>.  Specializing in Energizing Music and Race Announcing. 
                            <br>
                            <a href="mailto:kyle@fultonproductions.com">Request Pricing</a>
                        </div>
                    </div>
                </div>
                <div  class="list-group-item list-group-item-success">
                    <h4 class="h4">Announcer's Portal</h4>
                    <div class="row">
                        <div class="col-sm-4">
                            <img src="/graphics/announcer.jpg" alt="Announcers Portal" class="img-responsive">
                        </div>
                        <div class="col-sm-8">
                            Tell your announcer who is approaching the finish line so they can share with your spectators which race they are competing in, 
                            their age, gender, and where they are from before they even finish.  Nothing to look up.  A nice addition to races that are 
                            looking to tap in to the social and personal side of fitness events.  Renders very well on an iPad, lap top, or smart phone 
                            (not included).  Here is what the display <a href="http://www.gseannouncer.com/default.asp?event_id=651&which=announcer" 
                            onclick="openThis(this.href,1024,768);return false;">looks like</a>.
                            <br>
                            <a href="mailto:bob.schneider@gopherstateevents.com">Request Pricing</a>
                        </div>
                    </div>
                </div>
                <div  class="list-group-item list-group-item-danger">
                    <h4 class="h4">Featured Event</h4>
                    <div class="row">
                        <div class="col-sm-4">
                            <img src="/graphics/newspaper.jpg" alt="Featured Event" class="img-responsive">
                        </div>
                        <div class="col-sm-8">
                            Put your event in front of our site visitors on multiple pages, including our home page. (There are five examples to the left.)
                            For roughly the cost of an entry fee it will appear randomly until race day.  Our algorithm gives priority to the most recent 
                            events and changes every day.  We can put a banner and block ad together for you or you can come up with your own.
                            <br>
                            <a href="mailto:bob.schneider@gopherstateevents.com">Request Pricing</a>
                        </div>
                    </div>
                </div>
            </div>
  		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>