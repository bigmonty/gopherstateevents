<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID, lRaceID
Dim sClickPage, sTypeFilter, sGender
Dim Events

'Response.Redirect "/misc/taking_break.htm"

sClickPage = Request.ServerVariables("URL")

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

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
Dim lFeaturedEventsID
Dim sBannerImage
Dim bShowFeatured
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

ReDim Races(1, 0)
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events Results</title>
 
<link href="//cdn.datatables.net/1.10.2/css/jquery.dataTables.css" rel="stylesheet" type="text/css">
    
<script src="//code.jquery.com/jquery-2.1.4.min.js"></script>
<script src="//cdn.datatables.net/1.10.2/js/jquery.dataTables.min.js"></script>

<!-- bootstrap JavaScript & CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
  
<style type="text/css">
    div#results_info, div#results_paginate{
        display: none;
    }
</style>  

<script>
    $(document).ready(function () {
        $('body').delegate('#events,#races,#gender,#results_length select', 'change', function () {
            getData()
        });

        $('body').delegate('table#results th[data-sort], div#results_wrapper a', 'click', function (e) {
            e.preventDefault();
            e.stopPropagation();
             e.stopImmediatePropagation();
            $("th").removeClass("sorting_asc").removeClass("sorting_desc");
            if ($(this).attr("aria-sort") === undefined) {
                $(this).attr("aria-sort", "ascending");
                $(this).addClass("sorting_asc");
            } else {
                if ($(this).attr("aria-sort") === "descending") {
                    $(this).attr("aria-sort", "ascending");
                    $(this).addClass("sorting_asc");
                } else {
                    $(this).attr("aria-sort", "descending");
                    $(this).addClass("sorting_desc");
                }
            }
            getData();
        });

        $('body').delegate('#results_filter input', 'input', function () {
            getData()
        })

        var dt = null;

        function getUrlVars() {
            var vars = [], hash;
            var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
            for (var i = 0; i < hashes.length; i++) {
                hash = hashes[i].split('=');
                vars.push(hash[0]);
                vars[hash[0]] = hash[1];
            }
            return vars;
        }

        function initByRequest() {
            var params = getUrlVars();
            var _events = params["event_id"];
            if (_events != null && _events != "") {
                $("#events").val(_events);
                getData()
            }
        }

        initByRequest();

        function getData() {
            var races = $('#races').val() ? $('#races').val() : '',
            gender = $('#gender').val() ? $('#gender').val() : '',
            events = $('#events').val() ? $('#events').val() : '',
            results_length = $('#results_length select').val() ? $('#results_length select').val() : 10,
            results_filter = $('#results_filter input').val() ? $('#results_filter input').val() : '',
            results_sort = $('th[data-sort].sorting_asc, th[data-sort].sorting_desc').attr('data-sort') ? $('th[data-sort].sorting_asc, th[data-sort].sorting_desc').attr('data-sort') : 'Pts',
            results_sort_direction = ($('th[data-sort].sorting_asc').length > 0) ? 'ASC' : 'DESC',
            standings_page = $('a.paginate_button.current').attr('data-dt-idx') ? $('a.paginate_button.current').attr('data-dt-idx') : 2;

            var url = '/results/fitness_events/results_array.asp?event_id=' + events + '&gender=' + gender + '&race_id='
                    + races + '&results_filter=' + results_filter + '&results_length=' + results_length
                    + '&results_sort_direction=' + results_sort_direction + '&results_sort=' + results_sort;

            if (dt) {
                dt.fnSettings().sAjaxSource = url;
                dt.dataTable().fnDraw();
            }
            else {
                var settings = {
                    bServerSide: true,
                    sAjaxSource: url,
                    pagingType: "full_numbers",
                    "lengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],
                    fnServerData: function (src, data, cb) {
                        $.post(src, data, cb, "json");
                    }
                };
                dt = $('#results').dataTable(settings);
            }
        };
    });
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="row">
		<div class="col-sm-10">
            <div style="padding: 0 0 0 100px;">
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
            </div>

		    <%If CLng(lRaceID) = 0 Then%>
                <h3 class="h3 bg-primary">Gopher State Events Results</h3>
            <%Else%>
                <h3 class="h3 bg-primary">Gopher State Events Results: <%=sEventName%> (<%=Year(dEventDate)%>)</h3>
            <%End If%>

            <div class="col-xs-1"><label for="events">Event:</label></div>
            <div class="col-xs-3">
                <select class="form-control" name="events" id="events">
				    <option value="">&nbsp;</option>
				    <%For i = 0 to UBound(Events, 2)%>
					    <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						    <option value="<%=Events(0, i)%>" selected><%=Replace(Events(1, i), "''", "'")%>&nbsp;(<%=Events(2, i)%>)</option>
					    <%Else%>
						    <option value="<%=Events(0, i)%>"><%=Replace(Events(1, i), "''", "'")%>&nbsp;(<%=Events(2, i)%>)</option>
					    <%End If%>
				    <%Next%>
                </select>
            </div>
            <div class="col-xs-1"><label for="races">Race:</label></div>
            <div class="col-xs-2">
                <select class="form-control" name="races" id="races">
					<%For i = 0 to UBound(Races, 2) - 1%>
						<%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
							<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
						<%Else%>
							<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
						<%End If%>
					<%Next%>
                </select>
            </div>
            <div class="col-xs-1"><label for="gender">Gender:</label></div>
            <div class="col-xs-2">
                <select class="form-control" name="gender" id="gender">
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
            </div>

            <div class="col-xs-2">
                &nbsp;
            </div>

      		<%If Not CLng(lEventID) = 0 Then%>
                <%If sTimed = "y" Then%>
		            <%If Not CLng(lRaceID) = 0 Then%>
                        <ul class="list-inline" style="margin-left: 60px;">
		                    <%If Not sLocation = vbNullString Then%>
                                <li class="list-group-item list-group-item-warning">Location: <%=sLocation%></li>
                            <%End If%>

                            <li class="list-group-item list-group-item-warning">Distance: <%=sDist%></li>
                            <li class="list-group-item list-group-item-warning">Total Finishers:&nbsp;<%=iTtlRcds%></li>

                            <%If UBound(Races, 2) > 1 Then%>
                                <li class="list-group-item list-group-item-warning"><%=sRaceName%> Finishers:&nbsp;<%=iNumRace%></li>
                            <%End If%>
                        </ul>

   			            <%If CDate(dEventDate) > Date Then%>
				            <div class="bg-info" style="margin-bottom:2px;">
                                This event is currently scheduled for <%=dEventDate%>.  The results will be available on that date.
                            </div>
			            <%Else%>
                            <%If CDate(Date) < CDate(dEventDate) + 7 Then%>
			                    <%If bRsltsOfficial = False Then%>
				                    <div class="bg-danger" style="margin-bottom:2px;">
                                        <span style="color: red;font-weight: bold;">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE</span><br>
                                        Please report any issues to bob.schneider@gopherstateevents.com.
                                    </div>
			                    <%Else%>
				                    <div class="bg-info" style="margin-bottom:2px;">
                                        These results are now official.  Report errors 
				                        via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by telephone (612.720.8427).
                                    </div>
			                    <%End If%>
                            <%End If%>
			            <%End If%>

			            <ul class="list-inline" style="margin-left: 1px;">
                            <li class="list-group-item list-group-item-success">
                                <a href="javascript:pop('finishers_cert.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)" >Finishers Certificate</a>
                            </li>
                            <%If sIndivRelay = "relay" Then%>
                                <li class="list-group-item list-group-item-success">
                                    <a href="javascript:pop('relay_by_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)">Results 
                                    by Split</a>
                                </li>
                                <li class="list-group-item list-group-item-success">
                                    <a href="javascript:pop('relay_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)" >Results 
                                    w/Splits</a>
                                </li>
                            <%End If%>

				            <%If sHasSplits = "y" And sGender <> "B" Then%>
                                <li class="list-group-item list-group-item-warning">
                                    <a href="splits/results_w-splits.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Results w/Splits</a>
                                </li>
                                <li class="list-group-item list-group-item-warning">
                                    <a href="splits/rank_by_split.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Rank By Split</a>
                                </li>
                            <%End If%>
                                <li class="list-group-item list-group-item-danger">
                                    <a href="javascript:pop('print_rslts.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>',1000,700)">Print</a>
                                </li>
				            <%If sGender = "B" Then%>
                                <li class="list-group-item list-group-item-info">
                                    <a href="dwnld_combined.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
					                onclick="openThis(this.href,1024,768);return false;">Download</a>
                                </li>
                            <%Else%>
                                <li class="list-group-item list-group-item-success">
                                    <a href="dwnld_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>" 
					                onclick="openThis(this.href,1024,768);return false;">Download</a>
                                </li>
                            <%End If%>
				            <%If Session("role") = "admin" And CInt(iEventType) = 5 Then%>
                                <li class="list-group-item list-group-item-warning">
                                    <a href="usatf_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
					                onclick="openThis(this.href,1024,768);return false;">USATF Rslts</a>
                                </li>
                            <%End If%>
				            <%If sHasTeams = "y" Then%>
                                <li class="list-group-item list-group-item-danger">
                                    <a href="team_results.asp?race_id=<%=lRaceID%>" onclick="openThis(this.href,1024,768);return false;">Team Results</a>
                                </li>
                            <%End If%>
                                <li class="list-group-item list-group-item-info">
                                    <a href="/records/records.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Records</a>
                                </li>
                            <%If CInt(iRaceType) = 5 Then%>
                                <%If sShowAge = "y" Then%>
                                    <li class="list-group-item list-group-item-success">
                                        <a href="age_graded.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                        onclick="openThis(this.href,1024,768);return false;">Age-Graded</a>
                                    </li>
                                <%End If%>
                            <%End If%>
                            <%If CInt(iRaceType) >= 9 Then%>
                                <li class="list-group-item list-group-item-warning">
                                    <a href="trans_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Transitions</a>
                                </li>
                                <li class="list-group-item list-group-item-warning">
				                    <a href="multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Results w/Splits</a>
                                </li>
                            <%End If%>
				            <%If sGender = "B" Then%>
                                <%If CInt(iNumMAgeGrps) > 1 Then%>
				                    <li class="list-group-item list-group-item-danger">
                                        <a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=M"
                                        onclick="openThis(this.href,1024,768);return false;">Awards-M</a>
                                    </li>
                                    <li class="list-group-item list-group-item-danger">
				                        <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=M"
                                        onclick="openThis(this.href,1024,768);return false;">Age Grps-M</a>
                                    </li>
                                <%End If%>
                                <%If CInt(iNumFAgeGrps) > 1 Then%>
				                    <li class="list-group-item list-group-item-danger">
                                        <a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=F"
                                        onclick="openThis(this.href,1024,768);return false;">Awards-F</a>
                                    </li>
                                    <li class="list-group-item list-group-item-danger">
				                        <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=F"
                                        onclick="openThis(this.href,1024,768);return false;">Age Grps-F</a>
                                    </li>
                                <%End If%>
                            <%Else%>
                                <%If CInt(iNumAgeGrps) > 1 Then%>
				                    <li class="list-group-item list-group-item-danger">
                                        <a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                                        onclick="openThis(this.href,1024,768);return false;">Awards</a>
                                    </li>
                                    <li class="list-group-item list-group-item-danger">
				                        <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                                        onclick="openThis(this.href,1024,768);return false;">Age Grps</a>
                                    </li>
                                <%End If%>
                                <%If CLng(lSuppLegID) > 0 Then%>
                                    <li class="list-group-item list-group-item-info">
                                        <a href="/results/fitness_events/supp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>" 
                                        onclick="openThis(this.href,1024,768);return false;">Rslts w/Splits</a>
                                    </li>
                                <%End If%>
			                <%End If%>
                            <%If UBound(Races, 2) > 1 And sShowAge = "y" Then%>
                                <li class="list-group-item list-group-item-success">
                                    <a href="blended_results.asp?event_id=<%=lEventID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Blended Results</a>
                                </li>
                            <%End If%>
                        </ul>
                    <%End If%>

                    <div class="table-responsive">
		                <table class="table table-striped">
                            <thead>
			                    <tr>
				                    <th style="padding-right: 5px;">Pl</th>
                                    <th style="padding-right: 5px;">Bib</th>
				                    <th style="padding-right: 5px;">Name</th>
				                    <th style="padding-right: 5px;">M/F</th>
  				                    <%If sShowAge = "y" Then%>
                                        <th style="padding-right: 5px;">Age</th>
                                    <%Else%>
                                        <th style="padding-right: 5px;">Age Grp</th>
                                    <%End If%>
				                    <th style="padding-right: 5px;">Chip Time</th>
				                    <th style="padding-right: 5px;">Gun Time</th>
				                    <th style="padding-right: 5px;">Start Time</th>
				                    <th style="text-align:left;">From</th>
			                    </tr>
                            </thead>
                            <tfoot>
			                    <tr>
				                    <th style="padding-right: 5px;">Pl</th>
                                    <th style="padding-right: 5px;">Bib</th>
				                    <th style="padding-right: 5px;">Name</th>
				                    <th style="padding-right: 5px;">M/F</th>
  				                    <%If sShowAge = "y" Then%>
                                        <th style="padding-right: 5px;">Age</th>
                                    <%Else%>
                                        <th style="padding-right: 5px;">Age Grp</th>
                                    <%End If%>
				                    <th style="padding-right: 5px;">Chip Time</th>
				                    <th style="padding-right: 5px;">Gun Time</th>
				                    <th style="padding-right: 5px;">Start Time</th>
				                    <th style="text-align:left;">From</th>
			                    </tr>
                            </tfoot>
		                </table>
                    </div>
                <%Else%>
                    <p>This was a non-timed race.</p>
                <%End If%>
            <%End If%>
        </div>
		<div class="col-sm-2">
            <%If CLng(lEventID) > 0 Then%>
                <%If Not sLogo & "" = "" Then%>
                    <img class="img-responsive" src="/events/logos/<%=sLogo%>" alt="Event Logo" style="margin-top: 2px;">
                <%End If%>

                <div style="margin:2px 0 0 0;padding:0;text-align:center;">
                    <%If sGalleryLink = vbNullString Then%>
                        <%If Date < CDate(dEventDate) + 10 Then%>
                            <img src="/graphics/no_pix.png" alt="Pix Not Available Yet" class="img-responsive"  style="margin: 0;">
                        <%End If%>
                    <%Else%>
                        <a href="<%=sGalleryLink%>" onclick="openThis(this.href,1024,768);return false;">
                            <img src="/graphics/Camera-icon.png" alt="Race Photos" class="img-responsive"  style="margin: 0;">
                        </a>
                    <%End If%>
                </div>

                <%If Not CLng(lEventID) = 0 Then%>
                    <%If Not sWeather & "" = "" Then%>
                        <p style="margin:2px 0 0 0;padding:0;text-indent:0;font-size:0.85em;"><span style="font-weight:bold;">Weather:</span>&nbsp;<%=sWeather%></p>
                    <%End If%>

                    <%If Not sRaceReport & "" = "" Then%>
                        <p style="margin:2px 0 0 0;padding:0;text-indent:0;font-size:0.85em;"><span style="font-weight:bold;">Race Report:</span>&nbsp;<%=sRaceReport%></p>
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
	<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>