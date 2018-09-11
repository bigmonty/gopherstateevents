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
<title>GSE&copy; Service Offerings</title>
<meta name="description" content="Service offerings for Gopher State Events (GSE) timing service in Minnetonka, MN.">
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-md-10">
			<h1 class="h1">GSE Race Timing Fees & Services Provided</h1>

            <div style="margin: 10px 0 15px 0;padding: 5px;">
                <h4 class="h4">Our Price-Match Guarantee!</h4>

                <p style="font-size: 0.8em;">We guarantee to meet or beat any competitor's pricing to manage your event.  Just <a href="mailto:bob.schneider@gopherstateevents.com">send us</a>.
                last year's invoice or next year's estimate from a professional timing company and, subject to review, we will guarantee that price or 
                lower.  Please note that this offer is not subject to timing done by informal timing groups, groups using non-timing specific equipment, 
                or pricing based on unique pricing arrangements between event directors and a timing company.</p>
            </div>

            <h3 style="margin:10px;">Service Offerings</h3>

            <table>
                <tr>
                    <td style="width: 445px;" valign="top">
                        <h4 style="background-color: #ececd8;">Fitness Events</h4>

                        <ul>
                            <li>Course Certification</li>
                            <li>Course Measurement</li>
                            <li>Race Day Course Management</li>
                            <li>Volunteer Management</li>
                            <li>Pre-Race Publicity</li>
                            <li>Search Engine Optimization</li>
                            <li>Social Networking Promotion</li>
                        </ul>
                        <br>
                        <ul>
                            <li>Pre-Registration Input</li>
                            <li>Visible finish line clock</li>
                            <li>Conventional or RFID timing</li>
                            <li>Pre-race informational email</li>
                            <li>Map to site</li>
                            <li>Announcer</li>
                            <li>Race day data entry</li>
                            <li>Results email to all participants, usually within minutes of their finish</li>
                            <li>Awards-ready results on demand</li>
                            <li>Online results during the event</li>
                            <li>Free bibs upon request</li>
                            <li>Event records</li>
                            <li>Periodic results posting on site</li>
                            <li>Finish line pix and videos</li>
                            <li>Event Records</li>
                            <li>Virtual Series</li>
                        </ul>
                    </td>
                    <td style="width: 450px;" valign="top">
                        <h4 style="background-color: #ececd8;">CC Running/Nordic Ski</h4>

                        <ul>
                            <li>Online meet info</li>
                            <li>Secure meet director portal</li>
                            <li>Secure coach login</li>
                            <li>Visible finish line clock</li>
                            <li>Conventional or RFID Timing option</li>
                            <li>Map to site</li>
                            <li>Meet line-ups/start list emailed to all coaches</li>
                            <li>Race-day changes permitted</li>
                            <li>Awards-ready results almost immediately</li>
                            <li>Online results in minutes</li>
                            <li>Free bibs and pins</li>
                            <li>Event records</li>
                            <li>Finish line pix and videos online within 24 hours.</li>
                        </ul>
                        </td>
                </tr>
            </table>
			
		    <p>To discuss our process further, or to have us reserve a date for your event, contact us at <span style="font-weight:bold;">612.720.8427</span> or 
		    <a href="mailto:bob.schneider@gopherstateevents.com" style="font-weight:bold;">via email</a>.</p>
        </div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
