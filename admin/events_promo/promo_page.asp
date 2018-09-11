<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lPromoID, lEventID, lSeriesID
Dim iYear
Dim sEventName, sLocation, sLogo, sSeriesName, sSeriesNotes, sWebLink, sMessage, sMapLink, sOnlineReg
Dim SeriesEvents(), Races()
Dim dEventDate

lPromoID = Request.QueryString("promo_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TargetEvent FROM PromoEmail WHERE PromoEmailID = " & lPromoID & " ORDER BY PromoEmailID DESC"
rs.Open sql, conn, 1, 2
lEventID = rs(0).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventName, e.EventDate, e.Location, e.Logo, e.OnlineReg, e.Website, si.MapLink "
sql = sql & "FROM Events e INNER JOIN SiteInfo si ON e.EventID = si.EventID WHERE e.EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sLocation = Replace(rs(2).Value, "''", "'")
sLogo = rs(3).Value
sOnlineReg = rs(4).Value
If Not rs(5).Value & "" = "" Then sWebLink = rs(5).Value
If Not rs(6).Value & "" = "" Then sMapLink = rs(6).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Message FROM PromoEmail WHERE PromoEmailID = " & lPromoID
rs.Open sql, conn, 1, 2
If Not rs(0).Value & "" = "" Then sMessage = Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing

i = 0
ReDim Races(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceName, Dist, StartTime, OnlineRegLink FROM RaceData WHERE EventID = " & lEventID 
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Races(0, i) = Replace(rs(0).Value, "''", "'")
    Races(1, i) = Replace(rs(1).Value, "_", " ")
    Races(2, i) = rs(2).Value
    If Not rs(3).Value & "" = "" Then sOnlineReg = rs(3).Value
    i = i + 1
    ReDim Preserve Races(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

iYear = Year(dEventDate)
If CStr(lEventID) = vbNullString Then lEventID = 0

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
    ReDim SeriesEvents(4, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventName, EventDate, Location, Website, Logo FROM Events WHERE EventDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear 
    sql = sql & "' AND EventID <> " & lEventID & " ORDER BY EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If InSeries(rs(0).Value) = True Then
            SeriesEvents(0, i) = rs(0).Value
	        SeriesEvents(1, i) = Replace(rs(1).Value, "''", "'")
            SeriesEvents(2, i) = rs(2).Value
            SeriesEvents(3, i) = rs(4).Value
            SeriesEvents(4, i) = rs(5).Value
	        i = i + 1
	        ReDim Preserve SeriesEvents(4, i)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
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
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Event Promotion</title>
<meta name="description" content="Promotional race information for a Gopher State Events (GSE) timed event.">
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    html{
	    height:100%;
    }
</style>
</head>

<body>
<div class="container">
    <div>
        <a href="http://www.gopherstateevents.com/admin/events_promo/promo_page.asp?promo_id=<%=lPromoID%>"
        style="font-weight: bold;">View this page in a web browser</a>.
    </div>

    <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">

    <table class="table table-condensed">
        <tr>
            <td rowspan="2" style="width:200px;">
                <h3 class="h3">Mark Your Calendar:<br><%=sEventName%></h3>
                <h4 class="h4"><%=dEventDate%><br><%=sLocation%>.</h4>

                <%If Not sLogo & "" = "" Then%>
                    <%If Not sWebLink & "" = "" Then%>
                        <a href="<%=sWebLink%>">
                            <img class="img-responsive" src="http://www.gopherstateevents.com/events/logos/<%=sLogo%>" alt="Race Logo">
                        </a>
                    <%Else%>
                        <a href="http://www.gopherstateevents.com/events/raceware_events.asp?event_id=<%=lEventID%>">
                            <img class="img-responsive" src="http://www.gopherstateevents.com/events/logos/<%=sLogo%>" alt="Race Logo">
                        </a>
                    <%End If%>
                <%Else%>
                    &nbsp;
                <%End If%>
                <br>
                <div class="bg-success" style="text-align:center;">
                    <a href="http://www.gopherstateevents.com/misc/timers_wanted.asp"><h4 class="h4">Timers Needed!</h4></a>
                </div>
            </td>
            <td>
                <table class="table table-striped table-condensed">
                    <tr>
                        <th>Race</th>
                        <th>Distance</th>
                        <th>Start Time</th>
                    </tr>
                    <%For i = 0 To UBound(Races, 2) - 1%>
                        <tr>
                            <td><%=Races(0, i)%></td>
                            <td><%=Races(1, i)%></td>
                            <td><%=Races(2, i)%></td>
                        </tr>
                    <%Next%>
                </table>

                <%If Not sMessage = vbNullString Then%>
                    <p><%=sMessage%></p>
                <%End If%>

                <table class="table table-condensed">
                    <tr>
                        <td>
                            <%If Not sMapLink = vbNullString Then%>
                                <a href="<%=sMapLink%>" onclick="openThis(this.href,1024,768);return false;">
                                <img src="http://www.gopherstateevents.com/graphics/social_media/map_quest.jpg" alt="<%=sMapLink%>" class="img-responsive" style="width: 65px;">
                                </a>
                            <%End If%>
                        </td>
                        <td>
                        <td>
                            <%If Not sWebLink = vbNullString Then%>
                                <a href="<%=sWebLink%>" onclick="openThis(this.href,1024,768);return false;">
                                <img src="http://www.gopherstateevents.com/graphics/social_media/web_logo.jpg" alt="<%=sWebLink%>" class="img-responsive" style="width: 65px;">
                                </a>
                            <%End If%>
                        </td>
                        <td>
                            <%If Not sOnlineReg = vbNullString Then%>
                                <a href="<%=sOnlineReg%>" onclick="openThis(this.href,1024,768);return false;">
                                <img src="/graphics/social_media/reg_logo.jpg" alt="Online Registration" style="height: 75px;">
                                </a>
                            <%End If%>
                        </td>
                    </tr>
                </table>                
                
                <%If Not CLng(lSeriesID) = 0 Then%>
                    <div>
		                <h4 class="h4">Do All Events In The Series!</h4>
		                <ul style="list-style-type: none;display: block;">
			                <%For i = 0 To UBound(SeriesEvents, 2) - 1%>	
                                <li style="display: inline;">
                                    <%If SeriesEvents(4, i) & "" = "" Then%>
                                        <%If SeriesEvents(3, i) & "" = "" Then  'no logo-no website%>
                                            <a href="/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>" 
							                onclick="openThis(this.href,1024,768);return false;" rel="nofollow">
                                                <%=SeriesEvents(1, i)%> (<%=SeriesEvents(2, i)%>)
                                            </a>
                                        <%Else  'no logo-website%>
                                            <a href="<%=SeriesEvents(3, i)%>" onclick="openThis(this.href,1024,768);return false;" rel="nofollow">
                                                <%=SeriesEvents(1, i)%> (<%=SeriesEvents(2, i)%>)
                                            </a>
                                        <%End If%>
                                    <%Else%>
                                        <%If SeriesEvents(3, i) & "" = "" Then  'logo- no website%>
                                            <a href="/events/raceware_events.asp?event_id=<%=SeriesEvents(0, i)%>" 
							                    onclick="openThis(this.href,1024,768);return false;" rel="nofollow">
                                                <img src="http://www.gopherstateevents.com/events/logos/<%=SeriesEvents(4, i)%>" alt="Logo"
                                                        style="height: 100px;">
                                            </a>
                                        <%Else  'logo-website%>
                                            <a href="<%=SeriesEvents(3, i)%>" onclick="openThis(this.href,1024,768);return false;" rel="nofollow">
                                                <img src="http://www.gopherstateevents.com/events/logos/<%=SeriesEvents(4, i)%>" alt="Logo"
                                                        style="height: 100px;">
                                            </a>
                                        <%End If%>
                                    <%End If%>
                                </li>
			                <%Next%>
                        </ul>
                    </div>
                <%End If%>
            </td>
        </tr>
        <tr>
            <td>
                <h4 class="h4">Please Leave Me Alone!</h4>
                <p>
                    At GSE we send pre-race, results. and promotional emails to those people who we think might benefit.  We understand that not 
                    everyone appreciates these types of notifications and we want to make it very easy to prevent receiving them if that is your wish.  
                    To get on the "Do Not Send" list simply visit <a href="http://www.gopherstateevents.com/misc/do_not_send.asp" 
                    style="font-weight: bold;color: #f00;">this page</a>, enter your email address, and click the button.  Make sure you use the email 
                    address that this was sent to and we will put that email address on our "Do Not Send" list.  NOTE:  This will prevent you from 
                    receiving ANY emails from GSE (pre-race informational, individual results, promotional, etc.)
                </p>
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
