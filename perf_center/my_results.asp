<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lPartID, lEventID, lRaceID, lSeriesID, lEmailRsltsID, lFeaturedEventsID
Dim sEventName, sRaceName, sPartName, sMyTime, sSeriesName, sLogo, sSuppMsg, sBlockImage, sClickPage
Dim iYear
Dim SeriesEvents()
Dim dEventDate

'Response.Redirect "/misc/taking_break.htm"

lEventID = Request.QueryString("event_id")
If lEventID = "56" Then lEventID = "565"

lPartID = Request.QueryString("part_id")
lRaceID = Request.QueryString("race_id")
lEmailRsltsID = Request.QueryString("email_rslts_id")

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

'get supp message if it exists
'Set rs = Server.CreateObject("ADODB.Recordset")
'sql = "SELECT SuppMsg FROM EmailRslts WHERE EmailRsltsID = " & lEmailRsltsID
'rs.Open sql, conn, 1, 2
'If Not rs(0).Value & "" = "" Then sSuppMsg = Replace(rs(0).Value, "''", "'")
'rs.Close
'Set rs = Nothing

'get event data
sql = "SELECT EventName, EventDate, Logo FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
If Not rs(2).Value & "" = "" Then sLogo = rs(2).Value
Set rs = Nothing

iYear = Year(dEVentDate)

lSeriesID = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID FROM SeriesEvents WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then lSeriesID = rs(0).Value
rs.Close
Set rs = Nothing		

If CLng(lSeriesID) > 0 Then 
    'get series info
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing	
            
    'get series events	
    i = 0
    ReDim SeriesEvents(4, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventName, EventDate, Location, Website FROM Events WHERE EventDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear 
    sql = sql & "' AND EventID <> " & lEventID & " ORDER BY EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If InSeries(rs(0).Value) = True Then
            SeriesEvents(0, i) = rs(0).Value
	        SeriesEvents(1, i) = Replace(rs(1).Value, "''", "'")
            SeriesEvents(2, i) = rs(2).Value
            SeriesEvents(3, i) = rs(3).Value
            SeriesEvents(4, i) = rs(4).Value
	        i = i + 1
	        ReDim Preserve SeriesEvents(4, i)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
		
sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'")
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then sPartName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing
			
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FnlScnds FROM IndResults WHERE RaceID = " & lRaceID & " AND ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then sMyTime = ConvertToMinutes(rs(0).Value)
rs.Close
Set rs = Nothing

Private Function InSeries(lEventID)
    InSeries = False

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT SeriesID FROM SeriesEvents WHERE EventID = " & lEventID & " AND SeriesID = " & lSeriesID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then InSeries = True
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If CSng(sglScnds) <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(CSng(sglScnds) \ 60)
    sSecondPart = CStr(((CSng(sglScnds) / 60) - (CSng(sglScnds) \ 60)) * 60)
    
    'add leading zero to seconds if necessary
    If CSng(sSecondPart) < 10 Then
        sSecondPart = "0" & sSecondPart
    End If
    
    'make sure there are exactly two decimal places
    If Len(sSecondPart) < 5 Then
        If Len(sSecondPart) = 2 Then
            sSecondPart = sSecondPart & ".00"
        ElseIf Len(sSecondPart) = 4 Then
            sSecondPart = sSecondPart & "0"
        End If
    Else
        sSecondPart = Left(sSecondPart, 5)
    End If
    
    'do the conversion
    If CInt(sMinutePart) <= 60 Then
        ConvertToMinutes = sMinutePart & ":" & sSecondPart
    Else
        sHourPart = CStr(CSng(sMinutePart) \ 60)
        sMinutePart = CStr(CSng(sMinutePart) Mod 60)

        If Len(sMinutePart) = 1 Then
            sMinutePart = "0" & sMinutePart
        End If

        ConvertToMinutes = sHourPart & ":" & sMinutePart & ":" & sSecondPart
    End If
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
<title>My GSE&copy; Results</title>
<meta name="description" content="Individual Results for a Gopher State Events (GSE) timed event.">

<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap-theme.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/dist/css/bootstrap-submenu.min.css">

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
    <a href="http://www.gopherstateevents.com/perf_center/my_results.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;part_id=<%=lPartID%>&amp;email_rslts_id=<%=lEmailRsltsID%>"
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
	                    <h3 class="h3">My Results for <%=sEventName%> on <%=dEventDate%></h3>

	                    <p>Dear <%=sPartName%>:</p>
		
	                    <p>
                            Congratulations on your performance in the <%=sEventName%>&nbsp;<%=sRaceName%> on <%=dEventDate%>!  Your time was 
                            <span style="color: red;"><%=sMyTime%></span>.  You may view, print, and download complete results at 
                            <a href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_id=<%=lEventID%>" 
                            style="font-weight: bold;color: red;" onclick="openThis(this.href,1024,768);return false;">here</a>.
                        </p>

                        <%If Not sSuppMsg = vbNullString Then%>
                            <p><%=sSuppMsg%></p>
                        <%End If%>

                        <%If CLng(lSeriesID) > 0 Then%>
                            <p>Series results will be updated within 24 hours.</p>
                        <%End If%>

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
                                <a href="https://twitter.com/gsetiming">
                                    <img src="http://www.gopherstateevents.com/graphics/social_media/fb.png" height="30" alt="Facebook">
                                </a>
                            </li>
                            <li class="list-group-item">
                                <a href="http://www.youtube.com/channel/UCs09DthS7jEZy5srWZEDJQw">
                                    <img src="http://www.gopherstateevents.com/graphics/social_media/youtube.png" alt="YouTube" height="30">
                                </a>
                            </li>
                            <li class="list-group-item">
                                <a href="http://plus.google.com/100097568010679842973?prsrc=3" rel="publisher" style="text-decoration:none;">
                                    <img src="http://www.gopherstateevents.com/graphics/social_media/GooglePlus-512-Red.png" alt="Google+" height="30">
                                </a>
                            </li>
                            <li class="list-group-item">
                                <a href="http://www.linkedin.com/pub/bob-schneider/8/96a/876">
                                    <img src="http://www.gopherstateevents.com/graphics/social_media/LinkedIn-Logo.png" height="30" alt="View Bob Schneider's profile on LinkedIn">
                                </a>     
                            </li>
                            <li class="list-group-item">
                                <a href="https://twitter.com/gsetiming" class="twitter-follow-button" data-show-count="false">
                                    <img src="http://www.gopherstateevents.com/graphics/social_media/Twitter.png" alt="Follow @gsetiming" height="30">
                                </a>
                            </li>
                        </ul>

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
                    </div>
                    <div class="col-md-6">
                        <div class="bg-warning" style="text-align:center;">
                            <a href="http://www.gopherstateevents.com/misc/timers_wanted.asp"><h4 class="h4">Timers Needed!</h4></a>
                        </div>

                        <div class="bg-info" style="padding: 5px;">
                            <h4 class="h4"> Where's My Picture?</h4>
                           <p  style="font-size:0.9em;">
                                We attempt to take a finish line picture of everyone but occassionally finishers may be obscured or missing. Once 
                               ready a link to them will appear on the event's results page. You are free to download your own picture free of charge. Enjoy! 
                            </p>
                        </div>

                        <h4 class="h4 bg-danger">
                            <a href="http://www.gopherstateevents.com/misc/race_survey.asp?event_id=<%=lEventID%>&race_id=<%=lRaceID%>&part_id=<%=lPartID%>"
                               onclick="openThis(this.href,1024,768);return false;">
                                How'd we do?  Click to complete survey.
                            </a>
                        </h4>

                        <h4 class="h4">Your Next Race?</h4>
                        <div>
                            <a href="http://www.gopherstateevents.com/featured_events/featured_clicks.asp?featured_events_id=<%=lFeaturedEventsID%>&amp;click_page=<%=sClickPage%>" 
                                onclick="openThis(this.href,1024,768);return false;">
                                <img src="http://www.gopherstateevents.com/featured_events/images/<%=sBlockImage%>" alt="<%=sBlockImage%>" class="img-responsive">
                            </a>
                        </div>

                        <%If Not CLng(lSeriesID) = 0 Then%>
                            <div>
		                        <h4 class="h4">Run All Races In The Series!</h4>
		                        <ul class="list-group">
				                    <%For i = 0 To UBound(SeriesEvents, 2) - 1%>	
                                        <li class="list-group-item">
                                            <a href="http://www.gopherstateevents.com/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>" 
							                onclick="openThis(this.href,1024,768);return false;" rel="nofollow"><%=SeriesEvents(1, i)%> (<%=SeriesEvents(2, i)%>)</a>
                                        </li>
				                    <%Next%>
                                </ul>

                                <a href="http://www.gopherstateevents.com/series/series_results.asp?year=<%=Year(dEventDate)%>&series_id=<%=lSeriesID%>"
                                style="color: green;font-weight: bold;" onclick="openThis(this.href,1024,768);return false;">Here are the current series standings.</a>
                            </div>
                        <%End If%>

                        <br>
                        <div class="bg-primary">
                            <a class="bg-primary" href=" http://www.gopherstateevents.com/about/privacy.asp"onclick="openThis(this.href,1024,768);return false;">GSE Privacy Policy</a>
                        </div>
                    </div>
                </div>
            </td>
            <td style="width: 25%;padding-top: 15px;">
                <%If Not sLogo & "" = "" Then%>
                    <img class="img-responsive" src="http://www.gopherstateevents.com/events/logos/<%=sLogo%>" alt="Logo"></a>
                <%End If%>

                <h4 class="h4 bg-success">
                    <a href="http://www.gopherstateevents.com/records/records.asp?event_id=<%=lEventID%>"
                    onclick="openThis(this.href,1024,768);return false;">Click to View Event Records</a>
                </h4>

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
