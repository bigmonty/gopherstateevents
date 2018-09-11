<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lEventID, lFeaturedEventsID
Dim sEventName, sLogo, sFeaturedURL, sBlockImage, sClickPage
Dim RaceGallery()
Dim dEventDate
Dim bFound

lEventID = Request.QueryString("event_id")
If CStr(lEventID) & "" = "" Then lEventID = 0

sClickPage = Request.ServerVariables("URL")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FeaturedEventsID, BlockImage, Views FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date 
sql = sql & "' AND '" & Date + 360 & "') AND Active = 'y' ORDER BY NewID()"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    lFeaturedEventsID = rs(0).Value
    sBlockImage = rs(1).Value
    rs(2).Value = CLng(rs(2).Value) + 1
    rs.Update
End If
rs.Close
Set rs = Nothing

ReDim RaceGallery(0)
ReDim RaceVids(0)
	
If Not CLng(lEventID) = 0 Then
	sql = "SELECT EventName, EventDate, Logo FROM Events WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
    If Not rs(2).Value & "" = "" Then sLogo = rs(2).Value
   	Set rs = Nothing
	
    i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT GalleryLink FROM RaceGallery WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        RaceGallery(i) = rs(0).Value
        i = i + 1
        ReDim Preserve RaceGallery(i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<!--the above three meta tags must come first-->
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">

<link rel="icon" href="favicon.ico" type="image/x-icon"> 
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon"> 

<title>GSE&copy; Pictures Ready Notification</title>
<meta name="description" content="Media availability notification for a Gopher State Events (GSE) timed event.">

<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap-theme.min.css">
<link rel="stylesheet" href="dist/css/bootstrap-submenu.min.css">

<script src="https://code.jquery.com/jquery-2.1.4.min.js" defer></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js" defer></script>
<script src="dist/js/bootstrap-submenu.min.js" defer></script>

<script src="/misc/scripts.js"></script>

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

<body>
<div class="bg-success">
<a href="http://www.gopherstateevents.com/misc/pix-vids_notif.asp?event_id=<%=lEventID%>"
    style="font-weight: bold;">View this page in a web browser</a>.
</div>

<div class="container">
    <div class="row">
        <div class="col-md-4">
            <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
        </div>
        <div class="col-md-8">
	        <h3 class="h3">Free Finish Line Pix for <%=sEventName%></h3>
	        <h4 class="h4"><%=dEventDate%></h4>
        </div>
    </div>

    <table class="table table-condensed">
        <tr>
            <td>
		        <p class="text-success">
                    Hello.  We apologize for the electronic intrusion but we thought you might like to know that
                    finish line pictures for this event are now online.  To view just click on the camera icon. Pictures should be listed in
                    order of finish so if you know about where you finished you can scroll through the list to find yourself.

                    (Note:  we occassionally miss a finisher but we try hard to get everyone.  If we missed you please accept our 
                    apologies.
                </p>

                <ul class="list-inline">
                    <%For i = 0 To UBound(RaceGallery) - 1%>
                        <li>
                            <a href="<%=RaceGallery(i)%>" onclick="openThis(this.href,1024,768);return false;">
                                <img src="http://www.gopherstateevents.com/graphics/Camera-icon.png" alt="Race Photos" class="img-responsive">
                            </a>
                        </li>
                    <%Next%>
                </ul>
                
		        <p>Sincerely~</p>
		        <p>
                    <a href="mailto:bob.schneider@gopherstateevents.com">Bob Schneider</a>
                   <br>
		            GSE&copy; (Gopher State Events) &nbsp;<a href="http://www.gopherstateevents.com/">www.gopherstateevents.com</a>
                    <br>
		            612.720.8427
                </p>

                <ul class="list-inline">
                    <li class="list-group-item">
                        <a href="https://twitter.com/gsetiming" onclick="openThis(this.href,1024,768);return false;">
                            <img src="http://www.gopherstateevents.com/graphics/social_media/fb.png" height="30" alt="Facebook">
                        </a>
                    </li>
                    <li class="list-group-item">
                        <a href="http://www.youtube.com/channel/UCs09DthS7jEZy5srWZEDJQw" onclick="openThis(this.href,1024,768);return false;">
                            <img src="http://www.gopherstateevents.com/graphics/social_media/youtube.png" alt="YouTube" height="30">
                        </a>
                    </li>
                    <li class="list-group-item">
                        <a href="http://plus.google.com/100097568010679842973?prsrc=3" rel="publisher" style="text-decoration:none;"
                            onclick="openThis(this.href,1024,768);return false;">
                            <img src="http://www.gopherstateevents.com/graphics/social_media/GooglePlus-512-Red.png" alt="Google+" height="30">
                        </a>
                    </li>
                    <li class="list-group-item">
                        <a href="http://www.linkedin.com/pub/bob-schneider/8/96a/876" onclick="openThis(this.href,1024,768);return false;">
                            <img src="http://www.gopherstateevents.com/graphics/social_media/LinkedIn-Logo.png" height="30" alt="View Bob Schneider's profile on LinkedIn">
                        </a>     
                    </li>
                    <li class="list-group-item">
                        <a href="https://twitter.com/gsetiming" class="twitter-follow-button" data-show-count="false" onclick="openThis(this.href,1024,768);return false;">
                            <img src="http://www.gopherstateevents.com/graphics/social_media/Twitter.png" alt="Follow @gsetiming" height="30">
                        </a>
                    </li>
                </ul>
    
                <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                <!-- GSE Banner Ad -->
                <ins class="adsbygoogle"
                        style="display:inline-block;width:728px;height:90px"
                        data-ad-client="ca-pub-1381996757332572"
                        data-ad-slot="1411231449"></ins>
                <script>
                (adsbygoogle = window.adsbygoogle || []).push({});
                </script>

                <div class="bg-warning">
                    <h5 class="h5">Please Leave Me Alone!</h5>
                    <p>
                        AT GSE we send pre-race, results. and promotional emails to those people who we think might benefit.  We understand that not 
                        everyone appreciates these types of notifications and we want to make it very easy to prevent receiving them if that is your wish.  
                        To get on the "Do Not Send" list simply visit <a href="http://www.gopherstateevents.com/misc/do_not_send.asp" 
                        style="font-weight: bold;color: #f00;">this page</a>, enter your email address, and click the button.  Make sure you use the email 
                        address that this was sent to and we will put that email address on our "Do Not Send" list.  NOTE:  This will prevent you from 
                        receiving ANY emails from GSE (pre-race informational, individual results, promotional, etc.)
                    </p>
                </div>
           </td>
            <td style="padding-top: 15px;">
                <%If Not sLogo & "" = "" Then%>
                    <img class="img-responsive" src="http://www.gopherstateevents.com/events/logos/<%=sLogo%>" alt="Logo">
                    <br>
                <%End If%>

                 <div class="bg-success">
                     If we at Gopher State Events made your race experience a more enjoyable one, 
                    please invite other event directors to contact us and request a bid.
                 </div>

                <h4 class="h4">Your Next Race?</h4>
                <a href="http://www.gopherstateevents.com/featured_events/featured_clicks.asp?featured_events_id=<%=lFeaturedEventsID%>&amp;click_page=<%=sClickPage%>" 
                    onclick="openThis(this.href,1024,768);return false;">
                    <img src="http://www.gopherstateevents.com/featured_events/images/<%=sBlockImage%>" alt="<%=sBlockImage%>" class="img-responsive">
                </a>
                <br>
                <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                <!-- GSE Vertical ad -->
                <ins class="adsbygoogle"
                        style="display:block"
                        data-ad-client="ca-pub-1381996757332572"
                        data-ad-slot="6120632641"
                        data-ad-format="auto"></ins>
                <script>
                (adsbygoogle = window.adsbygoogle || []).push({});
                </script>
            </td>
        </tr>
    </table>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
