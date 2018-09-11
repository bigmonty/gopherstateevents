<%@ Language=VBScript%>
<%
Option Explicit
%>
<!--#include file = "../../../includes/set_lcid.asp" -->
<%
Dim conn, rs, sql
Dim i
Dim lRaceID, lFeaturedEventsID, lEventID, lSuppLegID
Dim sEventName, sRaceName, sBannerImage, sClickPage, sLogo, sWeather, sRaceReport, sIndivRelay, sShowAge, sTimed, sEventRaces, sLegName, sOtherName
Dim sEventClass, sLocation, sGender, sHasSplits, sDist, sChipStart, sAllowDuplAwds, sSortRsltsBy, sTypeFilter
Dim iNumLaps, iTtlRcds, iNumRace, iNumMAgeGrps, iRaceType, iNumAgeGrps, iEventType, iNumFAgeGrps
Dim RaceGallery(), CustomFields(), Races(), Events
Dim dEventDate
Dim bShowFeatured, bRsltsOfficial

sClickPage = Request.ServerVariables("URL")

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect "http://www.google.com"
If CLng(lRaceID) < 0 Then Response.Redirect "http://www.google.com"

iEventType = Request.QueryString("event_type")
If CStr(iEventType) = vbNullString Then iEventType = 5
If Not IsNumeric(iEventType) Then Response.Redirect("http://www.google.com")
If CLng(iEventType) < 0 Then Response.Redirect("http://www.google.com")

lSuppLegID = 0

sTimed = "y"

Select Case CInt(iEventType)
    Case 46
        sTypeFilter = "AND EventType IN(4, 6)"
    Case 910
        sTypeFilter = "AND EventType IN(9, 10)"
    Case Else
        sTypeFilter = "AND EventType = " & iEventType
End Select

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")

    If CStr(lEventID) = vbNullString Then lEventID = 0
    If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
    If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")
    If CStr(sGender) = vbNullString Then sGender = "B"

    lRaceID = GetFirstRace()
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
	sGender = Request.Form.Item("gender")
End If

'log this user if they are just entering the site
If Session("access_results") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'fitness_results')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate <= '" & Date & "' " & sTypeFilter & " ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    Events = rs.GetRows()
Else
    ReDim Events(2, 0)
End If
rs.Close
Set rs = Nothing

'determine if we should show ad or featured event
Dim iMyNum
Randomize
iMyNum = Int((rnd*10))+1

bShowFeatured = False
If iMyNum mod 2 = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FeaturedEventsID, BannerImage, Views FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date 
    sql = sql & "' AND '" & Date + 360 & "') AND Active = 'y' ORDER BY NewID()"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        lFeaturedEventsID = rs(0).Value
        sBannerImage = rs(1).Value
        rs(2).Value = CLng(rs(2).Value) + 1
        rs.Update
        bShowFeatured = True
    Else
        bShowFeatured = False
    End If
    rs.Close
    Set rs = Nothing
End If
	
i = 0
ReDim RaceGallery(0)

If CLng(lEventID) > 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sEventRaces = sEventRaces & rs(0).Value & ", "
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SuppLegID, LegName, OtherName FROM SuppLeg WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        lSuppLegID = rs(0).Value
        sLegName = Replace(rs(1).Value, "''", "'")
        If Not rs(2).Value & "" = "" Then sOtherName = Replace(rs(2).Value, "''", "'")
    End If
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ") AND FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iTtlRcds = rs.RecordCount
    rs.Close
    Set rs = Nothing

	'get event information
	sql = "SELECT EventName, EventDate, EventClass, Location, Logo FROM Events WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
    sEventClass = rs(2).Value
    sLocation = rs(3).Value
    sLogo = "/events/logos/" & rs(4).Value
	Set rs = Nothing

    'get the weather, race report
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Weather, RaceReport FROM RaceReport WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        If Not rs(0).Value & "" = "" Then sWeather = Replace(rs(0).Value, "''", "'")
        If Not rs(1).Value & "" = "" Then sRaceReport = Replace(rs(1).Value, "''", "'")
    End If
    rs.Close
  	Set rs = Nothing

    'get races
    ReDim Races(1, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
        Races(0, i) = rs(0).Value
        Races(1, i) = Replace(rs(1).Value, "''", "'")
        i = i + 1
        ReDim Preserve Races(1, i)
        rs.MoveNext
    Loop
	rs.Close
	Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventID FROM OfficialRslts WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then bRsltsOfficial = True
	rs.Close
	Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT SuppLegID FROM SuppLeg WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then lSuppLegID = rs(0).Value
	rs.Close
	Set rs = Nothing
	
    i = 0
    ReDim RaceGallery(0)
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
	
    If sEventClass = vbNullString Then sEventClass = "B"
	
    If CLng(lRaceID) = 0 Then lRaceID = GetFirstRace()

    'num race finishers
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " & lRaceID
    sql = sql & " AND ir.FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    iNumRace = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'check for team results
    Dim sHasTeams
    sHasTeams = "n"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID FROM Teams WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sHasTeams = "y"
    rs.Close
    Set rs = Nothing

    'check for custom fields
    i = 0
    ReDim CustomFields(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT CustomFieldsID, FieldName FROM CustomFields WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        CustomFields(0, i) = rs(0).Value
        CustomFields(1, i) = rs(1).Value
        i = i + 1
        ReDim Preserve CustomFields(1, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If sGender = "B" Then
        iNumMAgeGrps = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EndAge FROM AgeGroups WHERE Gender = 'M' AND RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iNumMAgeGrps = rs.RecordCount
        rs.Close
        Set rs = Nothing

        iNumFAgeGrps = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EndAge FROM AgeGroups WHERE Gender = 'F' AND RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iNumFAgeGrps = rs.RecordCount
        rs.Close
        Set rs = Nothing
    Else
        iNumAgeGrps = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iNumAgeGrps = rs.RecordCount
        rs.Close
        Set rs = Nothing
    End If

    sHasSplits = "n"
    sIndivRelay = "indiv"
	sql = "SELECT Dist, RaceName, Type, AllowDuplAwds, ChipStart, SortRsltsBy, NumSplits, IndivRelay, Timed, ShowAge, NumLaps "
    sql = sql & "FROM RaceData WHERE RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	sDist = rs(0).Value
	sRaceName = rs(1).Value
	iRaceType = rs(2).Value
	sAllowDuplAwds = rs(3).Value
    sChipStart = rs(4).Value
    sSortRsltsBy = rs(5).Value
    If CInt(rs(6).Value) > 0 Then sHasSplits = "y"
    sIndivRelay = rs(7).Value
    sTimed = rs(8).Value
    sShowAge = rs(9).Value
    iNumLaps = rs(10).Value
	Set rs = Nothing
End If

Private Function GetFirstRace()
    GetFirstRace = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetFirstRace = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>Gopher State Events Searchable Results</title>
   
<!--JQuery CDN-->  
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.0/jquery.min.js"></script>
<script src="//code.jquery.com/jquery-1.12.4.js"></script>

<!-- bootstrap JavaScript & CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
  
<!--Data Table references-->   
<link rel="stylesheet" href="https://cdn.datatables.net/1.10.15/css/jquery.dataTables.min.css">
<script src="https://cdn.datatables.net/1.10.15/js/jquery.dataTables.min.js"></script>

<script src="/misc/scripts.js"></script>

<script>
$(document).ready(function() {
    $('#results').DataTable( {
        "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]],
        "order": [[ 0, 'asc' ]],
        "columnDefs": [
	        {
		        "targets": [6,7,8],
		        "orderable": false
	       }],
         "ajax": {"url":"data_source_example.asp?race_id=<%=lRaceID%>"}
    } );
} );
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../../includes/header.asp" -->

     <div class="row no-print">
       <div class="col-xs-10">
            <%If bShowFeatured = True Then%>
                <a href="http://www.gopherstateevents.com/featured_events/featured_clicks.asp?featured_events_id=<%=lFeaturedEventsID%>&amp;click_page=<%=sClickPage%>" 
                    onclick="openThis(this.href,1024,768);return false;">
                    <img src="http://www.gopherstateevents.com/featured_events/images/<%=sBannerImage%>" alt="<%=sBannerImage%>" class="img-responsive">
                </a>
            <%Else%>
                <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                <!-- GSE Banner Ad -->
                <ins class="adsbygoogle"
                     style="display:inline-block;width:728px;height:90px"
                     data-ad-client="ca-pub-1381996757332572"
                     data-ad-slot="1411231449"></ins>
                <script>
                (adsbygoogle = window.adsbygoogle || []).push({});
                </script>
            <%End If%>

            <h2 class="h2 bg-success">Gopher State Events Results: <%=sEventName%> on <%=dEventDate%></h2>
 
            <div class="row">
                <div class="col-md-6">
			        <form role="form" class="form-inline" name="which_event" method="post" action="data_table_sample.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>">
			        <div class="form_group">
                        <label for="events">Event:</label>
			            <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()" style="font-size:0.9em;">
				            <option value="">&nbsp;</option>
				            <%For i = 0 to UBound(Events, 2)%>
					            <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						            <option value="<%=Events(0, i)%>" selected><%=Replace(Events(1, i), "''", "'")%>&nbsp;(<%=Events(2, i)%>)</option>
					            <%Else%>
						            <option value="<%=Events(0, i)%>"><%=Replace(Events(1, i), "''", "'")%>&nbsp;(<%=Events(2, i)%>)</option>
					            <%End If%>
				            <%Next%>
			            </select>
			            <input class="form-control" type="hidden" name="submit_event" id="submit_event" value="submit_event">
			            <input class="form-control" type="submit" name="get_event" id="get_event" value="Get These">
			        </div>
                    </form>
                </div>
                <div class="col-md-6">
                    <%If CLng(lEventID) = 0 Then%>
                        &nbsp;
                    <%Else%>
				        <form role="form" class="form-inline" name="get_races" method="post" action="data_table_sample.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>">
                        <div class="form_group">
				            <label for="races">Race:</label>
				            <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
					            <%For i = 0 to UBound(Races, 2) - 1%>
						            <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
							            <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
						            <%Else%>
							            <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
						            <%End If%>
					            <%Next%>
				            </select>
				            <label for="gender">Gender:</label>
				            <select class="form-control" name="gender" id="gender" onchange="this.form.get_race.click()">
					            <%Select Case sGender%>
						            <%Case "M"%>
                                        <option value="B">Combined</option>
							            <option value="M" selected>Male</option>
							            <option value="F">Female</option>
						            <%Case "F"%>
                                        <option value="B">Combined</option>
							            <option value="M">Male</option>
							            <option value="F" selected>Female</option>
						            <%Case Else%>
                                        <option value="B" selected>Combined</option>
							            <option value="M">Male</option>
							            <option value="F">Female</option>
					            <%End Select%>
				            </select>
				            <input class="form-control" type="hidden" name="submit_race" id="submit_race" value="submit_race">
				            <input class="form-control" type="submit" name="get_race" id="get_race" value="View">
                        </div>
				        </form>
                    <%End If%>
                </div>
            </div>

      		<%If CLng(lEventID) > 0 Then%>
                <%If sTimed = "y" Then%>
		            <%If Not CLng(lRaceID) = 0 Then%>
                        <div class="row">
                            <div class="col-xs-3">Total Finishers:&nbsp;<%=iTtlRcds%></div>
                            <div class="col-xs-3"><%=sRaceName%> Finishers:&nbsp;<%=iNumRace%></div>
                            <div class="col-xs-3">Distance: <%=sDist%></div>
                            <div class="col-xs-3">Location: <%=sLocation%></div>
                        </div>

   			            <%If CDate(dEventDate) > Date Then%>
				            <div class="bg-warning">
                                This event is currently scheduled for <%=dEventDate%>.  The results will be available on that date.
                            </div>
			            <%Else%>
                            <%If CDate(Date) < CDate(dEventDate) + 7 Then%>
			                    <%If bRsltsOfficial = False Then%>
				                    <div class="bg-warning">
                                        <span style="color: red;font-size: 1.0em;font-weight: bold;">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE</span>
                                        Report any issues <a href="mailto:bob.schneider@gopherstateevents.com">here</a>.
                                    </div>
			                    <%End If%>
                            <%End If%>
			            <%End If%>

			            <ul class="list-inline">
                            <%If UBound(CustomFields, 2) > 0 Then%>
                                <%For i = 0 To UBound(CustomFields, 2) - 1%>
                                    <li class="list-group-item-info">
                                        <a href="javascript:pop('../custom_fields_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;custom_fields_id=<%=CustomFields(0, i)%>',1000,700)"><%=CustomFields(1, i)%></a>
                                    </li>
                                <%Next%>
                            <%End If%>

                            <%If sIndivRelay = "relay" Then%>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('../relay_by_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)">Results 
                                    by Split</a>
                                </li>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('../relay_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)" >Results 
                                    w/Splits</a>
                                </li>
                            <%End If%>

                            <%If CInt(iNumLaps) > 1 Then%>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('../rslts_by_lap.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)">Results 
                                    by Lap</a>
                                </li>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('../rslts_w_laps.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)" >Results 
                                    w/Laps</a>
                                </li>
                            <%End If%>

				            <%If sHasSplits = "y" And sGender <> "B" Then%>
                                <li class="list-group-item-warning">
                                    <a href="../splits/results_w-splits.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Results With Splits</a>
                                </li>
                                <li class="list-group-item-warning">
                                    <a href="../splits/rank_by_split.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Rank By Split</a>
                                </li>
                            <%End If%>
                            <li class="list-group-item-info">
                                <a href="javascript:pop('../print_rslts.asp?rslts_event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>&amp;sort_rslts_by=<%=sSortRsltsBy%>',1000,700)">Print</a>
                            </li>
                            <li class="list-group-item-info">
                                <a href="../dwnld_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>" 
					            onclick="openThis(this.href,1024,768);return false;">Download</a>
                            </li>
				            <%If Session("role") = "admin" Then%>
                                <li class="list-group-item-info">
                                    <a href="../usatf_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
					                onclick="openThis(this.href,1024,768);return false;">USATF Rslts</a>
                                </li>
                            <%End If%>
				            <%If sHasTeams = "y" Then%>
                                <li class="list-group-item-info">
                                    <a href="../team_results.asp?race_id=<%=lRaceID%>" onclick="openThis(this.href,1024,768);return false;">Team Results</a>
                                </li>
                            <%End If%>
                                <li class="list-group-item-info">
                                    <a href="/records/records.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Records</a>
                                </li>
                            <%If CInt(iRaceType) = 5 Then%>
                                <%If sShowAge = "y" Then%>
                                    <li class="list-group-item-success">
                                        <a href="../age_graded.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                        onclick="openThis(this.href,1024,768);return false;">Age-Graded</a>
                                    </li>
                                <%End If%>
                            <%End If%>
                            <%If CInt(iRaceType) >= 9 Then%>
                                <li class="list-group-item-warning">
                                    <a href="../trans_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Trans Data</a>
                                </li>
                                <li class="list-group-item-warning">
				                    <a href="../multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Results w/Splits</a>
                                </li>
                            <%End If%>
				            <%If sGender = "B" Then%>
                                <%If CInt(iNumMAgeGrps) > 1 Then%>
				                    <li class="list-group-item-info">
                                        <a href="../awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=M"
                                        onclick="openThis(this.href,1024,768);return false;">Male Awards</a>
                                    </li>
                                    <li class="list-group-item-info">
				                        <a href="../age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=M"
                                        onclick="openThis(this.href,1024,768);return false;">Male Age Groups</a>
                                    </li>
                                <%End If%>
                                <%If CInt(iNumFAgeGrps) > 1 Then%>
				                    <li class="list-group-item-info">
                                        <a href="../awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=F"
                                        onclick="openThis(this.href,1024,768);return false;">Female Awards</a>
                                    </li>
                                    <li class="list-group-item-info">
				                        <a href="../age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=F"
                                        onclick="openThis(this.href,1024,768);return false;">Female Age Groups</a>
                                    </li>
                                <%End If%>
                            <%Else%>
                                <%If CInt(iNumAgeGrps) > 1 Then%>
				                    <li class="list-group-item-info">
                                        <a href="../awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                                        onclick="openThis(this.href,1024,768);return false;">Awards</a>
                                    </li>
                                    <li class="list-group-item-info">
				                        <a href="../age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                                        onclick="openThis(this.href,1024,768);return false;">Age Groups</a>
                                    </li>
                                <%End If%>
                                <%If CLng(lSuppLegID) > 0 Then%>
                                    <li class="list-group-item-success">
                                        <a href="/results/fitness_events/supp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>" 
                                        onclick="openThis(this.href,1024,768);return false;">Rslts w/Splits</a>
                                    </li>
                                <%End If%>
			                <%End If%>
                            <%If UBound(Races, 2) > 1 And sShowAge = "y" Then%>
                                <li class="list-group-item-success">
                                    <a href="../blended_results.asp?event_id=<%=lEventID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Blended Results</a>
                                </li>
                            <%End If%>
                        </ul>
                    <%End If%>

                    <%If CLng(lEventID) = 650 Then%>
                        <div class="bg-danger text-danger" style="text-align: right;margin-bottom: 2px;">
                            <a href="http://www.gopherstateevents.com/results/fitness_events/results_search/data_table_sample.asp?event_type=2&event_id=651">
                                View Vasaloppet Sunday Results
                            </a>
                        </div>
                    <%ElseIf CLng(lEventID) = 651 Then%>
                        <div class="bg-danger text-danger" style="text-align: right;margin-bottom: 2px;">
                            <a href="http://www.gopherstateevents.com/results/fitness_events/results_search/data_table_sample.asp?event_type=46&event_id=650">
                                View Vasaloppet Results
                            </a>
                        </div>
                    <%End If%>

                   <table id="results" class="display" cellspacing ="0" width="100%">
                        <thead>
                            <tr>
                                <th>Pl</th>
                                <th>Bib</th>
                                <th>First Name</th>
                                <th>Last Name</th>
                                <th>MFX</th>
                                <th>Age</th>
                                <th>Chip Time</th>
                                <th>Gun Time</th>
                                <th>Start Time</th>
                                <th>City</th>
                                <th>St</th>
                                <th>Certif</th>
                            </tr>
                        </thead>
                        <tfoot>
                            <tr>
                                <th>Pl</th>
                                <th>Bib</th>
                                <th>First Name</th>
                                <th>Last Name</th>
                                <th>MFX</th>
  				                <th>Age</th>
                                <th>Chip Time</th>
                                <th>Gun Time</th>
                                <th>Start Time</th>
                                <th>City</th>
                                <th>St</th>
                                <th>Certif</th>
                            </tr>
                        </tfoot>                        
                    </table>
                <%End If%>
           <%End If%>
        </div>
        <div class="col-xs-2">
            <%If CLng(lEventID) > 0 Then%>
                <%If Not sLogo & "" = "" Then%>
                    <img class="img-responsive" src="<%=sLogo%>" alt="Event Logo">
                <%End If%>

                <div style="margin:0;padding:0;text-align:center;">
                    <%If UBound(RaceGallery) = 0 Then%>
                        <%If Date < CDate(dEventDate) + 10 Then%>
                            <img src="/graphics/no_pix.png" alt="Pix Not Available Yet" class="img-responsive">
                        <%End If%>
                    <%Else%>
                        <%For i = 0 To UBound(RaceGallery) - 1%>
                           <a href="<%=RaceGallery(i)%>" onclick="openThis(this.href,1024,768);return false;">
                               <img src="/graphics/Camera-icon.png" alt="Race Photos" class="img-responsive">
                           </a>
                        <%Next%>
                    <%End If%>
                </div>

                <%If Not CLng(lEventID) = 0 Then%>
                    <%If Not sWeather & "" = "" Then%>
                        <p style="text-indent:0;font-size:0.85em;"><span style="font-weight:bold;">Weather:</span>&nbsp;<%=sWeather%></p>
                    <%End If%>

                    <%If Not sRaceReport & "" = "" Then%>
                        <p style="text-indent:0;font-size:0.85em;"><span style="font-weight:bold;">Race Report:</span>&nbsp;<%=sRaceReport%></p>
                    <%End If%>
                <%End If%>
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
        </div>
    </div>
</div>
</body>
</html>
