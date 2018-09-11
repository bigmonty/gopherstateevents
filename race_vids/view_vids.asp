<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lPartID, lEventID
Dim sEventName, sPartName
Dim EventVids()
Dim dEventDate
Dim bFound

lEventID = Request.QueryString("event_id")
lPartID = Request.QueryString("part_id")

If CStr(lEventID) & "" = "" Then lEventID = 0
If CStr(lPartID) & "" = "" Then lPartID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
Dim ViraEvents()
ReDim ViraEvents(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '" & Date & "' AND ShowOnline = 'y' AND EventID <> " & lEventID & " ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	ViraEvents(0, i) = rs(0).Value
	ViraEvents(1, i) = Replace(rs(1).Value, "''", "'")
	ViraEvents(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve ViraEvents(2, i)

    If i = 5 Then Exit Do

	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
	
i = 0
Dim PastEvents()
ReDim PastEvents(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate < '" & Date & "' AND EventID <> " & lEventID 
sql = sql & " ORDER BY EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	PastEvents(0, i) = rs(0).Value
	PastEvents(1, i) = Replace(rs(1).Value, "''", "'")
	PastEvents(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve PastEvents(2, i)

    If i = 5 Then Exit Do

	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
	
If Not CLng(lEventID) = 0 Then
	sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
	Set rs = Nothing
		
	If Not CLng(lPartID) = 0 Then
		sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
		Set rs = conn.Execute(sql)
		sPartName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
		Set rs = Nothing
	End If
	
    i = 0
    ReDim EventVids(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT VideoName, VideoLink FROM RaceVids WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While NOt rs.EOF
	    EventVids(0, i) = Replace(rs(0).Value, "''", "'")
        EventVids(1, i) = rs(1).Value
        i = i + 1
        ReDim Preserve EventVids(1, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>View GSE&copy; Event Videos</title>
<meta name="description" content="View Gopher State Events (GSE) video.">
<!--#include file = "../includes/js.asp" -->

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
<div class="container">
    <img class="img=responsive" src="/graphics/html_header.png" alt="Team Header">
	<div id="row">
        <table style="background-color: #fff;border-collapse: collapse;width: 800px;font-size:0.8em;">
            <tr>
                <td colspan="2" style="padding:0;margin:0;background-image:url('http://www.gopherstateevents.com/graphics/gse_results_header.png');width:800px;height:135px;background-repeat: no-repeat;">
 	                <div style="text-align:center;margin-left:450px;width:300px;">
		                <h2 style="padding-top:30px;"><%=sEventName%></h2>
		                <h4 class="h4"><%=dEventDate%></h4>
	                </div>

                    <p style="font-style: italic;margin: 0;padding: 0;border-top: 1px solid #ccc;border-bottom: 1px solid #ccc;text-align: center;">Does this page 
                        look a little disorganized?  
                     <a href="http://www.gopherstateevents.com/race_vids/race_vids.asp?event_id=<%=lEventID%>&amp;part_id=<%=lPartID%>"
                        style="font-weight: bold;">Click here to open it in a web browser</a>.</p>
              </td>
           </tr>
            <tr>
                <td style="width: 450px;padding:25px;" valign="top">
		            <p>Dear <%=sPartName%>:</p>
		
		            <p>We have posted some videos that you might like to look at.  You can view the videos below.  Enjoy!</p>
			
                    <ul>
                        <%For i = 0 To UBound(EventVids, 2) - 1%>
                            <li><a href="<%=EventVids(1, i)%>"><%=EventVids(0, i)%></a></li>
                        <%Next%>
                    </ul>
 
                    <div style="text-align:center;background-color: #ececec;padding: 5px;">
                        <h4 style="text-align:left;margin: 0;padding: 0;">Please Leave Me Alone!</h4>
                        <p style="text-align:left;margin: 0;padding: 0;font-size: 0.85em;">
                            AT GSE we send pre-race, results. and promotional emails to those people who we think might benefit.  We understand that not 
                            everyone appreciates these types of notifications and we want to make it very easy to prevent receiving them if that is your wish.  
                            To get on the "Do Not Send" list simply visit <a href="/misc/do_not_send.asp" 
                            style="font-weight: bold;color: #f00;">this page</a>, enter your email address, and click the button.  Make sure you use the email 
                            address that this was sent to and we will put that email address on our "Do Not Send" list.  NOTE:  This will prevent you from 
                            receiving ANY emails from GSE (pre-race informational, individual results, promotional, etc.)
                        </p>
                    </div>

                    <table style="border-collapse: collapse;text-align: center;margin: 0;padding: 0;">
                        <tr>
                            <td valign="top">
                                <a href="http://www.facebook.com/GopherStateEvents" target="_TOP" title="Gopher State Events, LLC"><img src="http://badge.facebook.com/badge/451629081557890.2085.618718499.png" style="border: 0px;" /></a>
                            </td>
                            <td valign="top"><a href="http://www.my-etraxc.com/"><img src="http://www.gopherstateevents.com/graphics/etraxc_ad.gif"  alt="eTRaXC"style="height: 65px;"></a></td>
                            <td valign="top"><a href="http://www.uptemporacemanagement.com/"><img src="http://www.gopherstateevents.com/graphics/uptempo.jpg" alt="Uptempo Ad" style="height: 65px;"></a></td>
                            <td valign="top">
                                <a href="https://twitter.com/gsetiming" class="twitter-follow-button" data-show-count="false">Follow @gsetiming</a>
                                <script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="width: 280px;text-align:center;padding: 10px;"valign="top">
                    <table>
                       <tr>
                            <td style="padding:0;height:150px;background-color:#ececd8;text-align: left;">
                                <h4 style="color: #039;text-align: left;margin: 5px 0 0 0;padding: 0;">Coming GSE Events</h4>

				                <%For i = 0 To UBound(ViraEvents, 2) - 1%>	
					                <h5 style="font-weight:bold;margin:5px 5px 0 5px;padding: 0;text-align: left;">
						                <a href="/events/raceware_events.asp?event_id=<%=ViraEvents(0, i)%>" 
							                onclick="openThis(this.href,1024,768);return false;" rel="nofollow"><%=ViraEvents(1, i)%></a>
					                </h5>
					                <h5 style="margin:0 0 0 10px;font-weight: normal;padding: 0;text-align: left;"><%=ViraEvents(2, i)%></h5>
				                <%Next%>
                            </td>
		                </tr>
		                <tr>
                            <td style="padding:0;height:150px;background-color:#ececd8;text-align: left;">
		                        <h4 style="margin-top:10px;color:#eb9c12;text-align: left;margin: 5px 0 0 0;padding:0;">Past GSE Events</h4>
		    
				                <%For i = 0 To UBound(PastEvents, 2) - 1%>	
					                <h5 style="font-weight:bold;margin:5px 5px 0 5px;padding: 0;text-align: left;">
						                <a href="/events/raceware_events.asp?event_id=<%=PastEvents(0, i)%>" 
							                onclick="openThis(this.href,1024,768);return false;" rel="nofollow"><%=PastEvents(1, i)%></a>
					                </h5>
					                <h5 style="margin:0 0 0 10px;font-weight: normal;padding: 0;text-align: left;"><%=PastEvents(2, i)%></h5>
				                <%Next%>
                            </td>
		                </tr>
		                <tr>
                            <td style="text-align: center;padding-top: 15px;">
		                        <a href="http://www.gopherstateevents.com/calendar/calendar.asp" style="font-size:0.9em;margin-top:10px;">GSE Event Calendar</a>
                                <br>
		                        <a href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_id=<%=lEventID%>" style="font-size:0.9em;">View All Results</a>
                           </td>
                        </tr>   
                    </table>
                </td>
            </tr>
        </table>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
