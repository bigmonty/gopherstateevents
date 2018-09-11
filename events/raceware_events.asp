<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID, lSeriesID
Dim i, j, k
Dim sEventName, sSiteName, sAddress, sMapLink, sDirections, sHomeImage, sClub, sComments, sEmail, sWebSite, sErrMsg, sLogo, sInfoSheet, sOnlineReg
Dim sPacketPickup, sSeriesName, sSeriesLogo, sSeriesInfo
Dim iLastAge, iEventType, iYear
Dim RaceArray(), MAgeGrps(), FAgeGrps()
Dim dEventDate, dWhenShutdown

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")
iYear = Year(Date)

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID FROM SeriesEvents WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then lSeriesID = rs(0).Value
rs.Close
Set rs = Nothing

If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

If Clng(lSeriesID) > 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName, Logo, InfoSheet FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    sSeriesLogo = rs(1).Value
    sSeriesInfo = rs(2).Value
    rs.Close
    Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT InfoSheet FROM InfoSheet WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sInfoSheet = "http://www.gopherstateevents.com/events/info_sheets/" & rs(0).Value
End If
rs.Close
Set rs = Nothing

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventName, e.EventDate, e.WebSite, e.Club, ed.Email, e.Comments, si.SiteName, si.Address, si.MapLink, e.WhenShutdown, "
sql = sql & "e.Logo, e.PacketPickup, e.EventType FROM Events e INNER JOIN EventDir ed ON e.EventDirID = ed.EventDirID INNER JOIN SiteInfo si ON e.EventID = si.EventID "
sql = sql & "WHERE e.EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
    sWebSite = rs(2).Value 
	If Not rs(3).Value & "" = "" Then sClub = Replace(rs(3).Value, "''", "'")
	sEmail = rs(4).Value
	If Not rs(5).Value & "" = "" Then sComments = Replace(rs(5).Value, "''", "'")
	If Not rs(6).Value & "" = "" Then sSiteName = Replace(rs(6).Value, "''", "'")
	If Not rs(7).Value & "" = "" Then sAddress = Replace(rs(7).Value, "''", "'")
	sMapLink = rs(8).Value
	dWhenShutdown = CDate(rs(9).Value)
    sLogo = rs(10).Value
    If Not rs(11).Value & "" = "" Then spacketPickup = Replace(rs(11).Value, "''", "'")
    iEventType = rs(12).Value
End If
rs.Close
Set rs = Nothing

'get information
i = 0
ReDim RaceArray(11, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Dist, Type, EntryFeePre, EntryFee, StartTime, Certified, StartType, MAwds, FAwds, RaceID, RaceName, OnlineRegLink FROM RaceData "
sql = sql & "WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	For j = 0 to 10
		RaceArray(j, i) = rs(j).Value
	Next
		
    RaceArray(11, i) = rs(11).Value
    sOnlineReg = RaceArray(11, i)

	i = i + 1
	ReDim Preserve RaceArray(11, i)
		
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetMAgeGrps(lRaceID)
	ReDim MAgeGrps(0)
	k = 0

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EndAge, NumAwds FROM AgeGroups WHERE (Gender = 'm') AND (RaceID = " & lRaceID
	sql = sql & ") ORDER BY EndAge"
	rs.Open sql, conn, 1, 2

	If rs.RecordCount = 1 Then
		MAgeGrps(0) = "None"
			
		k = k + 1
		ReDim Preserve MAgeGrps(k)
	Else
		Do While Not rs.EOF
			If k = 0 Then
				MAgeGrps(k) = rs(0).Value & " and Under (" & rs(1).Value & ") "
				iLastAge = rs(0).Value
			Else
				If rs(0).Value = 110 Then
					MAgeGrps(k) = CInt(iLastAge) + 1 & " and Over (" & rs(1).Value & ")"
				Else
					MAgeGrps(k) = CInt(iLastAge) + 1 & " - " & rs(0).Value & " (" & rs(1).Value & ") "
					iLastAge = rs(0).Value
				End If
			End If
			
			k = k + 1
			ReDim Preserve MAgeGrps(k)
			
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
End Sub

Private Sub GetFAgeGrps(lRaceID)
	ReDim FAgeGrps(0)
	k = 0
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EndAge, NumAwds FROM AgeGroups WHERE (Gender = 'f') AND (RaceID = " & lRaceID
	sql = sql & ") ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	
	If rs.RecordCount = 1 Then
		FAgeGrps(0) = "None"
			
		k = k + 1
		ReDim Preserve FAgeGrps(k)
	Else
		Do While Not rs.EOF
			If k = 0 Then
				FAgeGrps(k) = rs(0).Value & " and Under (" & rs(1).Value & ") "
				iLastAge = rs(0).Value
			Else
				If rs(0).Value = 110 Then
					FAgeGrps(k) = CInt(iLastAge) + 1 & " and Over (" & rs(1).Value & ")"
				Else
					FAgeGrps(k) = CInt(iLastAge) + 1 & " - " & rs(0).Value & " (" & rs(1).Value & ") "
					iLastAge = rs(0).Value
				End If
			End If
			
			k = k + 1
			ReDim Preserve FAgeGrps(k)
			
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
End Sub

Private Function GetThisType(lEventType)
	sql = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs = conn.Execute(sql)
	GetThisType = rs(0).Value
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Event Page:<%=sEventName%> on <%=dEventDate%></title>
<meta name="description" content="GSE Fitness Event Information for timing of road races, nordic ski, showshoe events, mountain bike, duathlon, and triathlons.">
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
    <h4 class="h4"><%=sEventName%> - <%=dEventDate%></h4>
    <div class="row">
	    <div class="col-sm-8">
            <table class="table">
                <tr>
			        <th>
				        <%If Not sAddress = vbNullString Then%>
					        Address:
				        <%End If%>
			        </th>
			        <td><%=sAddress%></td>
                </tr>
                <tr>
			        <th>Pre-Reg 'til:</th>
			        <td><%=dWhenShutdown%></td>
                </tr>
                <tr>
                    <%If sClub = vbNullString Then%>
                        <td colspan="2">&nbsp;</td>
                    <%Else%>
			            <th>Host:</th>
			            <td><%=sClub%></td>
                    <%End If%>
                </tr>
                <tr>
                    <%If sPacketPickup = vbNullString Then%>
                        <td colspan="2">&nbsp;</td>
                    <%Else%>
			            <th>Packet Pickup:</th>
			            <td><%=sPacketPickup%></td>
                    <%End If%>
                </tr>
                <tr>
			        <%If Date >= CDate(dEventDate) Then%>
				        <th>Results:</th>
				        <td><a href="/results/fitness_events/results.asp?event_type=<%=iEventType%>&event_id=<%=lEventID%>&first_rcd=1">Results</a></td>
			        <%End If%>
                </tr>
		        <tr>
                    <th>Notes:</th>
			        <td>
                        <%=sComments%>
                    </td>
		        </tr>
                <%If CLng(lSeriesID) > 0 Then%>
                    <tr>
                        <td>&nbsp;</td>
                        <td>
                            <h4 style="background: none;">This race is a part of a series.  Series information can be found 
                                <a href="/series/series_info.asp?series_id=<%=lSeriesID%>&amp;year=<%=iYear%>"
                                    style="color: red;">here</a>.</h4>
                            <%If sSeriesInfo & "" = "" Then%>
                                <%If Not sSeriesLogo & "" = "" Then%>
                                    <img src="/admin/series/logos/<%=sSeriesLogo%>" alt="Info Sheet" style="width: 150px;float: left;">
                                <%End If%>
                            <%Else%>
                                <%If sSeriesLogo & "" = "" Then%>
                                    <a href="/admin/series/info_sheets/<%=sSeriesInfo%>" target="_blank">Info Sheet</a>
                                <%Else%>
                                    <a href="/admin/series/info_sheets/<%=sSeriesInfo%>" target="_blank"><img src="/admin/series/logos/<%=sSeriesLogo%>" 
                                        alt="Info Sheet" style="width: 150px;float: left;"></a>
                                <%End If%>
                            <%End If%>
                        </td>
                    </tr>
                <%End If%>
            </table>
        </div>
        <div class="col-sm-4">
            <%If Not sLogo & "" = "" Then%>
                <img src="/events/logos/<%=sLogo%>" alt="Logo" style="width: 150px;margin: 0 0 5px 5px;">
                <br>
            <%End If%>
            <%If Not sSiteName = vbNullString Then%>
				<a href="<%=sMapLink%>" onclick="openThis(this.href,1024,768);return false;">
                <img src="/graphics/social_media/map_quest.jpg" alt="<%=sSiteName%>" style="height: 75px;">
                </a>
                <br>
			<%End If%>
            <%If Not sOnlineReg = vbNullString Then%>
				<a href="<%=sOnlineReg%>" onclick="openThis(this.href,1024,768);return false;">
                <img src="/graphics/social_media/reg_logo.jpg" alt="Online Registration" style="height: 75px;">
                </a>
                <br>
			<%End If%>
            <%If Not sInfoSheet = vbNullString Then%>
				<a href="<%=sInfoSheet%>" onclick="openThis(this.href,1024,768);return false;">
                <img src="/graphics/social_media/info_logo.jpg" alt="Info Sheet" style="height: 65px;">
                </a>
                <br>
			<%End If%>
            <%If Not sWebsite = vbNullString Then%>
				<a href="<%=sWebsite%>" onclick="openThis(this.href,1024,768);return false;">
                <img src="/graphics/social_media/web_logo.jpg" alt="<%=sWebsite%>" style="height: 75px;">
                </a>
			<%End If%>
        </div>
	</div>
    <div class="row">
        <h4 class="h4">Races</h4>
			<table class="table">
		    <%For i = 0 to UBound(RaceArray, 2) - 1%>
			    <tr>
				    <th style="text-align:left;background-color:#ececd8;" colspan="6">
                        Race Name:&nbsp;<span style="font-weight:normal;"><%=RaceArray(10, i)%></span>
				    </th>
			    </tr>
			    <tr>
				    <th>Distance:</th>
				    <td><%=Replace(RaceArray(0, i), "_", " ")%></td>
				    <th>Race Type:</th>
				    <td><%=GetThisType(RaceArray(1, i))%></td>
				    <th>Start Time:</th>
				    <td><%=RaceArray(4, i)%></td>
			    </tr>
			    <tr>
				    <th>Start Type:</th>
				    <td><%=RaceArray(6, i)%></td>
				    <th>Male Awards:</th>
				    <td><%=RaceArray(7, i)%></td>
				    <th>Female Awards:</th>
				    <td><%=RaceArray(8, i)%></td>
			    </tr>
			    <tr>
				    <th>Pre-Reg Fee:</th>
				    <td>$<%=RaceArray(2, i)%></td>
				    <th>Race Day Fee:</th>
				    <td>$<%=RaceArray(3, i)%></td>
				    <th>Certified?</th>
				    <td><%=RaceArray(5, i)%></td>
			    </tr>
			    <tr>
				    <th style="text-align:center;" colspan="3">Mens Age Groups (Awards):</th>
                    <th style="text-align:center;" colspan="3">Womens Age Groups (Awards):</th>
			    </tr>
			    <tr>
				    <td style="text-align:center;" colspan="3">
					    <%Call GetMAgeGrps(RaceArray(9,i))%>
                        <ul style="list-style-type: none;">
					        <%For j = 0 to UBound(MAgeGrps) - 1%>
						        <li><%=MAgeGrps(j)%></li>
					        <%Next	%>
                        </ul>
				    </td>
				    <td style="text-align:center;" colspan="3">
					    <%Call GetFAgeGrps(RaceArray(9,i))%>
                        <ul style="list-style-type: none;">
					        <%For j = 0 to UBound(FAgeGrps) - 1%>
						        <li><%=FAgeGrps(j)%></li>
					        <%Next	%>
                        </ul>
				    </td>
			    </tr>
		    <%Next%>
	    </table>
    </div>
<!--#include file = "../includes/footer.asp" --> 
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>