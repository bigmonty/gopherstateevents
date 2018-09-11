<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lThisMeet
Dim sMeetName, sLogo, sGalleryLink
Dim dMeetDate

lThisMeet = Request.QueryString("meet_id")
If CStr(lThisMeet) & "" = "" Then lThisMeet = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

ReDim RaceGallery(0)
	
If Not CLng(lThisMeet) = 0 Then
	sql = "SELECT MeetName, MeetDate, Logo FROM Meets WHERE MeetsID = " & lThisMeet
	Set rs = conn.Execute(sql)
	sMeetName = Replace(rs(0).Value, "''", "'")
	dMeetDate = rs(1).Value
    If Not rs(2).Value & "" = "" Then sLogo = rs(2).Value
   	Set rs = Nothing
	
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EmbedLink FROM RaceGallery WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sGalleryLink = rs(0).Value
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE&copy; CC-Nordic Pix Ready</title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">
<meta name="description" content="Pre-race information for a Gopher State Events (GSE) timed event.">

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
<table style="background-color: #fff;border-collapse: collapse;width: 550px;font-size:0.8em;">
    <tr>
        <td colspan="2" style="padding:0;margin:0;background-image:url('http://www.gopherstateevents.com/graphics/default_event_header.png');width:550px;height:140px;background-repeat: no-repeat;">
		    <h1 style="margin: 10px 0 0 250px;padding-top:0;color: #eb9c12;text-align: center;"><%=sMeetName%></h1>
		    <h3 style="padding:0;margin:0 0 0 250px;text-align: center;"><%=dMeetDate%></h3>

            <p style="font-style: italic;margin: 50px 0 0 0;padding: 0;border-top: 1px solid #ccc;border-bottom: 1px solid #ccc;text-align: center;">Does this page 
                look a little disorganized?  
             <a href="http://www.gopherstateevents.com/ccmeet_admin/manage_meet/pix-vids_notif.asp?meet_id=<%=lThisMeet%>"
                style="font-weight: bold;">Click here to open it in a web browser</a>.</p>
      </td>
   </tr>
    <tr>
        <td style="width: 410px;padding:0 10px 0 10px;background-color: #fff;" valign="top">
            <div style="margin: 0;padding: 5px;">
		        <p style="margin-top: 0;padding-top: 0;">Hello.  We apologize for the electronic intrusion but we thought you might like to know that
                finish line pictures for this meet are now online.  For your convenience, we are including a link to the gallery 
                    <a href="<%=sGalleryLink%>" style="font-weight: bold;color: #f00;">here</a>.</p>
                <p style="color: red;">(Note:  we occassionally miss a finisher but we try hard to get everyone.  If we missed someone please accept our 
                    apologies.</p>

		        <p style="margin-left: 10px;">Sincerely~</p>
		        <p style="margin-left: 10px;"><a href="mailto:bob.schneider@gopherstateevents.com">Bob Schneider</a><br>
		        GSE&copy; (Gopher State Events) &nbsp;<a href="http://www.gopherstateevents.com/">www.gopherstateevents.com</a><br>
		        612.720.8427</p>

                <h4 style="text-align:left;margin: 0;padding: 10px 0 0 10px;">Please Leave Me Alone!</h4>
                <p style="text-align:left;margin: 0;padding: 0 0 0 10px;font-size: 0.85em;">
                    AT GSE we send pre-race, results, pix-vids notification, and promotional emails to those people who we think might benefit.  We understand that not 
                    everyone appreciates these types of notifications and we want to make it very easy to prevent receiving them if that is your wish.  
                    To get on the "Do Not Send" list simply visit <a href="/misc/do_not_send.asp" 
                    style="font-weight: bold;color: #f00;">this page</a>, enter your email address, and click the button.  Make sure you use the email 
                    address that this was sent to and we will put that email address on our "Do Not Send" list.  NOTE:  This will prevent you from 
                    receiving ANY emails from GSE.
                </p>
            </div>
       </td>
        <td style="width: 170px;text-align:center;padding: 0;" valign="top">
            <p style="color: #039;text-align: left;font-size: 0.85em;margin: 0;padding: 0;">If we at Gopher State Events made your meet experience a more enjoyable one, 
            please invite other Meet directors to contact us and request a bid.</p>

            <%If Not sLogo & "" = "" Then%>
                <hr>
                <div style="margin: 0;padding: 0;"><img src="/events/logos/<%=sLogo%>" style="width: 150px;float: right;margin: 5px;" alt="Logo"></div>
                <hr>
            <%End If%>
            <table style="border-collapse: collapse;">
               <tr>
                    <td style="text-align: right;width: 300px;">
                        <a href="https://twitter.com/gsetiming" class="twitter-follow-button" data-show-count="false">Follow @gsetiming</a>
                        <br>
                        <script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>
                        <a href="http://www.facebook.com/GopherStateEvents" target="_TOP" title="Gopher State Events, LLC"><img src="http://badge.facebook.com/badge/451629081557890.2085.618718499.png" style="border: 0px;" /></a>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
