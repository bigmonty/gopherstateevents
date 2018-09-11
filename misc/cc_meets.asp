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
<html lang="en">
<head>
<title>Gopher State Events Cross-Country Meet Timing</title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">
<meta name="description" content="Cross-country meet timing (RFID timing) by Gopher State Events (GSE).">

<style type="text/css">
    html{
	    height:100%;
    }

    a{
	    text-decoration:none;
	    color:#036;
	    padding:0;
	    margin:0;
    }
</style>
</head>

<body style="margin:10px;padding:10px;font-family:arial, serif;background-color:#ececd8;">
<table style="background-color: #fff;border-collapse: collapse;width: 600px;font-size:0.85em;">
    <tr>
        <td colspan="2" style="padding:0;margin:0;background-image:url('http://www.gopherstateevents.com/graphics/cc_meet_header.png');width:800px;height:135px;background-repeat: no-repeat;">
            &nbsp;
        </td>
   </tr>
    <tr>
        <td colspan="2">            
            <p style="font-style: italic;margin: 0;padding: 0;border-top: 1px solid #ccc;border-bottom: 1px solid #ccc;text-align: center;">Does this page 
                look a little disorganized?  
                <a href="http://www.gopherstateevents.com/misc/cc_meets.asp" style="font-weight: bold;">Click here to open it in a web browser</a>.
            </p>
        </td>
    </tr>
    <tr>
        <td style="width: 450px;padding:0 25px 25px 25px;font-size: 0.8em" valign="top">
            <h3>Please consider Gopher State Events (GSE) for your Cross-Country/Nordic Ski Meet Timing.</h3>

            <p>
                We believe that we offer more in the way of services per dollar spent than any other timing company in the Upper Midwest.  Our services are
                friendly and reliable.  Please view our features and cost breakdown below and 
                <a href="http://www.gopherstateevents.com/misc/vira_contact.asp" style="font-weight:bold;">contact us for more information.</a>
            </p>
            <table style="border-collapse;border-collapse;">
                <tr>
                    <td valign="top">
                        <h4 style="margin: 0;">Features:</h4>
                        <ul style="margin: 0;padding-left: 15px;">
                            <li>Finish line pictures online later that evening.</li>
                            <li>Finish line videos online within a couple of days.</li>
                            <li>Results online shortly after every race in the meet.</li>
                            <li>Coaches receive results via email while still on site.</li>
                            <li>Awards-ready results almost immediately.</li>
                            <li>Bibs and pins supplied free of charge.</li>
                        </ul>
                    </td>
                    <td valign="top">
                        <h4 style="margin: 0;">Costs:</h4>
                        <ul style="margin: 0;padding-left: 15px;">
                            <li>$250 Meet Fee</li>
                            <li>$0.50/Participant Fee</li>
                            <li>$0.60/Mile Round Trip</li>
                        </ul>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
		                <p>Sincerely~</p>
		                <p><a href="mailto:bob.schneider@gopherstateevents.com">Bob Schneider</a>
		                GSE&copy; (Gopher State Events)<br>
		                <a href="http://www.gopherstateevents.com/">www.gopherstateevents.com</a><br>
		                612.720.8427</p>
                    </td>
                </tr>
            </table>
        </td>
        <td style="width: 200px;text-align:center;padding: 0 10px 10px 10px;"valign="top">
           <div>		                
               <a href="http://www.gopherstateevents.com/calendar/calendar.asp" style="font-size:0.9em;margin-top:10px;">
               <img src="http://www.gopherstateevents.com/graphics/calendar.jpg" style="width: 195px;float: right;">
               </a>
            </div>

            <div style="text-align:center;background-color: #ececec;padding: 5px;">
                <h4 style="text-align:left;margin: 0;padding: 0;">Please Leave Me Alone!</h4>
                <p style="text-align:left;margin: 0;padding: 0;font-size: 0.8em;">
                    AT GSE we send pre-race, results. and promotional emails to those people who we think might benefit.  We understand that not 
                    everyone appreciates these types of notifications and we want to make it very easy to prevent receiving them if that is your wish.  
                    To get on the "Do Not Send" list simply visit <a href="http://www.gopherstateevents.com/misc/do_not_send.asp" 
                    style="font-weight: bold;color: #f00;">this page</a>, enter your email address, and click the button.  Make sure you use the email 
                    address that this was sent to and we will put that email address on our "Do Not Send" list.  NOTE:  This will prevent you from 
                    receiving ANY emails from GSE (pre-race informational, individual results, promotional, etc.)
                </p>
            </div>
             <div>
                <a href="https://twitter.com/gsetiming" class="twitter-follow-button" data-show-count="false">Follow @gsetiming</a>
                <br>
                <script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>
                <a href="http://www.facebook.com/GopherStateEvents" target="_TOP" title="Gopher State Events, LLC"><img src="http://badge.facebook.com/badge/451629081557890.2085.618718499.png" style="border: 0px;" /></a>
            </div>
         </td>
    </tr>
</table>

<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
