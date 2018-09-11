<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lPartID, lEventID, lRaceID, lSeriesID, lFeaturedEventsID
Dim sEventName, sRaceName, sPartName, sMyGender, sMyAgeGrp, sWhichGndr, sMapLink, sStartTime, sSuppMsg, sLogo, sSeriesName, sPacketPickup, sWebsite
Dim sShowAge, sBlockImage, sClickPage
Dim iMyAge, iMyBib, iEdition, iYear, iEventType
Dim SeriesEvents()
Dim dEventDate
Dim bFound

lEventID = Request.QueryString("event_id")
lPartID = Request.QueryString("part_id")
lRaceID = Request.QueryString("race_id")
sSuppMsg = Request.QueryString("supp_msg")

If CStr(lEventID) & "" = "" Then lEventID = 0
If CStr(lPartID) & "" = "" Then lPartID = 0
If CStr(lRaceID) & "" = "" Then lRaceID = 0

sClickPage = Request.ServerVariables("URL")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FeaturedEventsID, BlockImage, Views FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date + 7
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
	
If Not CLng(lEventID) = 0 Then
    iEdition = 0
    iMyBib = 0
    iMyAge = 0

	sql = "SELECT e.EventName, e.EventDate, e.Edition, si.MapLink, e.Logo, e.PacketPickup, e.Website, e.EventType FROM Events e INNER JOIN SiteInfo si "
    sql = sql & "ON e.EventID = si.EventID WHERE e.EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
    iEdition = rs(2).Value
    sMapLink = rs(3).Value
    If Not rs(4).Value & "" = "" Then sLogo = rs(4).Value
    If Not rs(5).Value & "" = "" Then sPacketPickup = Replace(rs(5).Value, "''", "'")
    sWebsite = rs(6).Value
    iEventType = rs(7).Value
   	Set rs = Nothing

    iYear = Year(dEventDate)

    lSeriesID = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesID FROM SeriesEvents WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then lSeriesID = rs(0).Value
    rs.Close
    Set rs = Nothing		

    If Not CLng(lSeriesID) = 0 Then 
        'get series info
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
        rs.Open sql, conn, 1, 2
        sSeriesName = Replace(rs(0).Value, "''", "'")
        rs.Close
        Set rs = Nothing	
            
        'get series events	
        i = 0
        ReDim SeriesEvents(2, 0)
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventID, EventName, EventDate, Location FROM Events WHERE EventDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear 
        sql = sql & "' AND EventID <> " & lEventID & " ORDER BY EventDate"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If InSeries(rs(0).Value) = True Then
                SeriesEvents(0, i) = rs(0).Value
	            SeriesEvents(1, i) = Replace(rs(1).Value, "''", "'")
                SeriesEvents(2, i) = rs(2).Value
	            i = i + 1
	            ReDim Preserve SeriesEvents(2, i)
            End If
	        rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
		
	If Not CLng(lRaceID) = 0 Then
		sql = "SELECT RaceName, StartTime, ShowAge FROM RaceData WHERE RaceID = " & lRaceID
		Set rs = conn.Execute(sql)
		sRaceName = Replace(rs(0).Value, "''", "'")
        sStartTime = rs(1).Value
        sShowAge = rs(2).Value
		Set rs = Nothing
		
        Select Case CLng(lPartID)
            Case 1
                sPartName = "Some Participant"
                sMyGender = "M"
                sMyAgeGrp = "(appropriate age group here, if any)"
                iMyAge = "(their age here if displaying ages is allowed)"
                iMyBib = "1234"
		    Case Else
                If Not CLng(lPartID) = 0 Then
			        sql = "SELECT FirstName, LastName, Gender FROM Participant WHERE ParticipantID = " & lPartID
			        Set rs = conn.Execute(sql)
			        sPartName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
			        sMyGender = rs(2).Value
			        Set rs = Nothing
		
			        'get my age group
			        sql = "SELECT AgeGrp, Age, Bib FROM PartRace WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
			        Set rs = conn.Execute(sql)
			        sMyAgeGrp = rs(0).Value
                    iMyAge = rs(1).Value
                    iMyBib = rs(2).Value
			        Set rs = Nothing
                End If
		End Select
	End If
End If

Private Function InSeries(lEventID)
    InSeries = False

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT SeriesID FROM SeriesEvents WHERE EventID = " & lEventID & " AND SeriesID = " & lSeriesID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then InSeries = True
    rs2.Close
    Set rs2 = Nothing
End Function
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

<title>GSE&copy; Pre-Race Event Information</title>
<meta name="description" content="Pre-race information for a Gopher State Events (GSE) timed event.">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js" integrity="sha384-0mSbJDEHialfmuBBQP6A4Qrprq5OVfW37PRR3j5ELqxss1yVqOtnepnHVP9aJ7xS" crossorigin="anonymous"></script>
<script src="dist/js/bootstrap-submenu.min.js"></script>

<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap-theme.min.css">
<link rel="stylesheet" href="dist/css/bootstrap-submenu.min.css">

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
    <a href="http://www.gopherstateevents.com/misc/pre_race.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;part_id=<%=lPartID%>&amp;supp_msg=<%=sSuppMsg%>"
    style="font-weight: bold;">View this page in a web browser</a>.
</div>

<div class="container">
    <div class="row">
        <div class="col-md-6">
            <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
        </div>
        <div class="col-md-6">
	        <h3 class="h3"><%=sEventName%></h3>
	        <h4 class="h4"><%=dEventDate%></h4>
        </div>
    </div>

    <table class="table table-condensed">
        <tr>
            <td style="width: 75%;">
                <div class="row">
                    <div class="col-md-6">
		                <h5 class="h5">Dear <%=sPartName%>:</h5>
		
		                <p>
                            Your race information for <%=sEventName%> on <%=dEventDate%> is below.  Please check for accuracy.  This event
                            is being timed by <a href="http://www.gopherstateevents.com/" style="font-weight: bold;"
                            onclick="openThis(this.href,1024,768);return false;">Gopher State Events</a>.
                        </p>
	
                        <p class="bg-danger text-danger">
                            <strong>IMPORTANT NOTE:</strong> Please remove any RFID chips/tags from your shoes.  Some shoe companies (ASICS for one)
                            place an RFID tag under the insole that is easily removed.  You may replace it, or not, as you wish after the race.  Please
                            also ensure that you have no RFID tags on your shoes or clothing from previous events.
                        </p>
                        
			            <%If Not sSuppMsg = vbNullString Then%>
                            <div class="bg-warning">
                                <h4 class="h4">From the Event Director:</h4>
                                <%=sSuppMsg%>
                            </div>
                        <%End If%>

                        <h4 class="h4" style="margin-bottom: 0;padding-bottom: 0;">Your Info:</h4>
		                <ul class="list-group">
                            <li class="list-group-item">Race Registered For: <%=sRaceName%></li>
                            <li class="list-group-item">Race Time: <%=sStartTime%></li>
                            <li class="list-group-item">Gender: <%=sMyGender%></li>
			                <%If sShowAge = "y" Then%>
                                <%If iMyAge = "99" Then%>
                                    <li class="list-group-item">Age: na</li>
                                <%Else%>
                                    <li class="list-group-item">Age: <%=iMyAge%></li>
                                <%End If%>
                            <%End If%>
			                <%If Not sMyAgeGrp = "110 and Under" Then%>
                                <li class="list-group-item">Age Group: <%=sMyAgeGrp%></li>
                            <%End If%>
			                <%If CInt(iMyBib) > 0 Then%>
                                <li class="list-group-item">You have been assigned bib <%=iMyBib%>.</li>
                            <%End If%>
			                <%If Not sMapLink & "" = "" Then%>
                                <li class="list-group-item"><a href="<%=sMapLink%>" style="font-weight: bold;" onclick="openThis(this.href,1024,768);return false;">Event Location</a>.</li>
                            <%End If%>
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
                                <a href="http://plus.google.com/100097568010679842973?prsrc=3" rel="publisher" style="text-decoration:none;" onclick="openThis(this.href,1024,768);return false;">
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
                    </div>
                    <div class="col-md-6">
                        <div class="bg-warning" style="text-align:center;">
                            <a href="http://www.gopherstateevents.com/misc/timers_wanted.asp"><h4 class="h4">Timers Needed!</h4></a>
                        </div>
                        <div class="bg-info">
                            <h4 class="h4">Get Results Via Text Message</h4>
                            <p>
                                If you would like to get your results this way, please 
                                <a href="http://www.gopherstateevents.com/misc/sms_kiosk.asp?part_id=<%=lPartID%>&amp;my_bib=<%=iMyBib%>&amp;event_id=<%=lEventID%>" 
                                style="font-weight: bold;color: red;" onclick="openThis(this.href,1024,768);return false;">click here</a>.
                                We will NOT use your mobile number for ANY REASON other than sending you results for your performance in our races.  AND WE WILL
                                NEVER, EVER schlep it off to someone else, for fee or free.  PROMISE!  Don't trust us?  Don't do it.
                            </p>
	                    </div>

                        <div class="bg-success">
                            <h4 class="h4">Packet Pickup Information:</h4>
                            <%If sPacketPickup = vbNullString Then%>
                                <p>You can pick up your bib and any other information made available by the event
                                organizers on the day of the race.  Please arrive AT LEAST 30 minutes prior to the race.</p>  
                            <%Else%>
                                <%=sPacketPickup%>
                            <%End If%>
                        </div>

                        <div class="bg-info">
                            <h4 class="h4">Results Email & Pix</h4>
                            <p style="margin: 0 0 5px 0;padding: 0;font-size: 0.85em;text-align: left">
                                You should receive results within a few minutes after finishing.  Later in the day a link to our finish line pix will appear on the 
                                results page.  You will be notified when they are online and you can download your finish line picture free-of-charge.
                            </p>
                        </div>
                        <div class="bg-warning">
                            <h4 class="h4">Online Results</h4>
                            <p>
                                After the race you can find your results 
                                <a href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_id=<%=lEventID%>" 
                                style="font-weight:bold;" onclick="openThis(this.href,1024,768);return false;">here</a>.
                                If we have your email address on file you will receive your individual results shortly after the race.
                                <span style="font-style: italic;">Have a GREAT fitness or racing experience!.</span>
                            </p>
			            </div>
                    </div>
                </div>

                <div class="row bg-warning">
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
            <td style="width: 25%;padding-top: 15px;">
                <%If Not sLogo & "" = "" Then%>
                    <%If sWebsite = vbNullString Then%>
                        <img src="/events/logos/<%=sLogo%>" alt="Logo">
                    <%Else%>
                        <a href="<%=sWebsite%>" onclick="openThis(this.href,1024,768);return false;">
                            <img class="img-responsive" src="http://www.gopherstateevents.com/events/logos/<%=sLogo%>" alt="Logo"></a>
                    <%End If%>
                <%End If%>
                <%If CInt(iEdition) <> 1 Then%>
                    <div class="bg-success">
                        <a href="http://www.gopherstateevents.com/records/records.asp?event_id=<%=lEventID%>"
                        onclick="openThis(this.href,1024,768);return false;">Event Records For <%=sEventName%></a>
                    </div>
                <%End If%>

                <h4 class="h4">Your Next Race?</h4>
                <div>
                    <a href="http://www.gopherstateevents.com/featured_events/featured_clicks.asp?featured_events_id=<%=lFeaturedEventsID%>&amp;click_page=<%=sClickPage%>" 
                        onclick="openThis(this.href,1024,768);return false;">
                        <img src="http://www.gopherstateevents.com/featured_events/images/<%=sBlockImage%>" alt="<%=sBlockImage%>" class="img-responsive">
                    </a>
                    <br>
                </div>

                <%If Not CLng(lSeriesID) = 0 Then%>
                    <div>
		                <h4 class="h4">Run All Races In The Series!</h4>
		                <ul style="margin: 5px 0 0 0;">
				            <%For i = 0 To UBound(SeriesEvents, 2) - 1%>	
                                <li>
                                    <a href="http://www.gopherstateevents.com/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>" 
							        onclick="openThis(this.href,1024,768);return false;" rel="nofollow"><%=SeriesEvents(1, i)%> (<%=SeriesEvents(2, i)%>)</a>
                                </li>
				            <%Next%>
                        </ul>

                        <a href="http://www.gopherstateevents.com/series/age_neutral_results.asp?year=<%=Year(dEventDate)%>&series_id=<%=lSeriesID%>"
                        style="color: green;font-weight: bold;"onclick="openThis(this.href,1024,768);return false;">Here are the current series standings.</a>
                    </div>
                <%End If%>
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
