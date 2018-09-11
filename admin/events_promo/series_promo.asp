<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lPromoID, lEventID, lSeriesID
Dim iYear
Dim sEventName, sLocation, sLogo, sSeriesName, sSeriesNotes, sWebLink, sInfoSheet, sRaceDist, sPreFee, sRegFee, sMessage, sMapLink
Dim SeriesEvents()
Dim dEventDate

lPromoID = Request.QueryString("promo_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate, Location, TargetEvent, WebLink FROM PromoEmail WHERE PromoEmailID = " & lPromoID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sLocation = Replace(rs(2).Value, "''", "'")
lEventID = rs(3).Value
sWebLink = rs(4).Value
rs.Close
Set rs = Nothing

iYear = Year(dEventDate)

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Logo FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sLogo = rs(0).Value
rs.Close
Set rs = Nothing

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
    sql = "SELECT SeriesName, SeriesNotes FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    If Not rs(1).Value & "" = "" Then sSeriesNotes = Replace(rs(1).Value, "''", "'")
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

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT InfoSheet, WebLink, MapLink, RaceDist, PreFee, RegFee, Message FROM PromoEmail WHERE PromoEmailID = " & lPromoID
rs.Open sql, conn, 1, 2
sInfoSheet = rs(0).Value
If Not rs(1).Value & "" = "" Then 
    sWebLink = Replace(rs(1).Value, "http://", "")
    sWebLink = "http://" & Trim(sWebLink)
End If

If Not rs(2).Value & "" = "" Then 
    sMapLink = Replace(rs(2).Value, "http://", "")
    sMapLink = "http://" & Trim(sMapLink)
End If

sRaceDist = rs(3).Value
sPreFee = rs(4).Value
sRegFee = rs(5).Value
If Not rs(6).Value & "" = "" Then sMessage = Replace(rs(6).Value, "''", "'")
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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Event Promotion</title>
<meta name="description" content="Promotional information for a Gopher State Events (GSE) series event.">
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    html{
	    height:100%;
    }
</style>
</head>

<body>
<div class="bg-info">
    <a href="http://www.gopherstateevents.com/admin/events_promo/series_promo.asp?promo_id=<%=lPromoID%>"
    style="font-weight: bold;">View this page in a web browser</a>.
</div>

<div class="container">
        <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
        <table class="table table-condensed">
        <tr>
            <td rowspan="2">
                <%If CLng(lSeriesID) = 0 Then%>
                    <h3 class="h3">Mark Your Calendar:<br><%=sEventName%></h3>
                    <h4 class="h4"><%=dEventDate%><br><%=sLocation%>.</h4>
                <%Else%>
                    <h3 class="h3">Mark Your Calendar for the next race in the <%=sSeriesName%></h3>
                    <h4 class="h4"><%=sEventName%></h4>
                    <h5 class="h5"><%=dEventDate%><br><%=sLocation%>.</h5>
                <%End If%>

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
                <%If Not sInfoSheet & "" = "" Then%>
				    <a href="<%=sInfoSheet%>" onclick="openThis(this.href,1024,768);return false;">
                        <img src="http://www.gopherstateevents.com/graphics/social_media/info_logo.jpg" alt="Info Sheet" class="img-responsive" style="width: 65px;">
                    </a>
                <%End If%>
                <%If Not sMapLink = vbNullString Then%>
				    <a href="<%=sMapLink%>" onclick="openThis(this.href,1024,768);return false;">
                    <img src="http://www.gopherstateevents.com/graphics/social_media/map_quest.jpg" alt="<%=sMapLink%>" class="img-responsive" style="width: 65px;">
                    </a>
			    <%End If%>
            </td>
           <td class="bg-success">
                <label style="font-weight: bold;">Distance(s):</label>&nbsp;<%=sRaceDist%><br>
                <label style="font-weight: bold;">Pre-Reg Fee:</label>&nbsp;<%=sPreFee%><br>
                <label style="font-weight: bold;">Race Day Fee:</label>&nbsp;<%=sRegFee%>

                <%If Not sMessage = vbNullString Then%>
                    <p><%=sMessage%></p>
                <%End If%>
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
                <%'If CLng(lEventID) = 528 or Clng(lEventID) = 545 Then%>
                    <div>
                        <h4 class="h4 bg-info">How about this idea:  a "Broken Marathon"</h4>
                        <a href="/misc/broken_marathon.asp"><img class="img-responsive" src="/graphics/broken_marathon.jpg" alt="The Broken Marathon"></a>
                    </div>
                <%'End If%>
            </td>
        </tr>
    </table>


    <div class="bg-warning">
        <h4 class="h4">Please Leave Me Alone!</h4>
        <p>
            AT GSE we send pre-race, results. and promotional emails to those people who we think might benefit.  We understand that not 
            everyone appreciates these types of notifications and we want to make it very easy to prevent receiving them if that is your wish.  
            To get on the "Do Not Send" list simply visit <a href="www.gopherstateevents.com/misc/do_not_send.asp" 
            style="font-weight: bold;color: #f00;">this page</a>, enter your email address, and click the button.  Make sure you use the email 
            address that this was sent to and we will put that email address on our "Do Not Send" list.  NOTE:  This will prevent you from 
            receiving ANY emails from GSE (pre-race informational, individual results, promotional, etc.)
        </p>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
