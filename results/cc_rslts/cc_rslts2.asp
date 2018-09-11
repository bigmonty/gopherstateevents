<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs, rs2, sql2
Dim i, j, k, m
Dim lMeetID, lRaceID, lSeriesID
Dim sRaceName, sMeetName, sGradeYear, sOrderResultsBy, sScoreMethod, sTeamName, sClickPage, sShowResults, sStageRace
Dim sUnits, sMeetSite, sWeather, sRaceDist, sTeamScores, sResultsNotes, sErrMsg, sLogo, sAdvancement, sIndivRelay
Dim sShowIndiv, sSeriesGender, sSeriesName
Dim iDist, iNumScore, iNumFin, iRaceFin, iNumLaps
Dim RsltsArr, TmRslts, TempArr(9), SortArr(9), MeetArray, Races, RankArr()
Dim dMeetDate
Dim bRsltsOfficial, bTmRsltsReady

'Response.Redirect "/misc/taking_break.htm"

sClickPage = Request.ServerVariables("URL")

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

bTmRsltsReady = True

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate <= '" & Now() & "' AND ShowOnline = 'y' "
sql = sql & "AND Sport = 'Cross-Country' ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
MeetArray = rs.GetRows()
Set rs = Nothing

If Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
	If CStr(lRaceID) = vbNullString Then lRaceID = 0
	If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
	If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")
ElseIf Request.Form.Item("submit_meet") = "submit_meet" Then
	lMeetID = Request.Form.Item("meets")
    If CStr(lMeetID) = vbNullString Then lMeetID = 0
    If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
    If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")
End If

If Not CLng(lMeetID) = 0 Then
    iNumFin = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID FROM IndRslts WHERE MeetsID = " & lMeetID & " AND FnlScnds > 0 and Place > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iNumFin = rs.RecordCount
    rs.Close
    Set rs = Nothing

    bRsltsOfficial = False
	sql = "SELECT MeetsID FROM OfficialRslts WHERE MeetsID = " & lMeetID
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
        bRsltsOfficial = False
    Else
        bRsltsOfficial = True
    End If
	Set rs = Nothing

	sql = "SELECT MeetName, MeetDate, MeetSite, Weather, Logo FROM Meets WHERE MeetsID = " & lMeetID
	Set rs = conn.Execute(sql)
	sMeetName = rs(0).Value
	dMeetDate = rs(1).Value
	If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then sWeather = Replace(rs(3).Value, "''", "'")
    sLogo = rs(4).Value
	Set rs = Nothing
	
	'get year for roster grades
	If Month(dMeetDate) <= 7 Then
	    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
	Else
	    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
	End If

    If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear
	
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID 
	sql = sql & " AND RaceName NOT In ('B-TBD', 'G-TBD') ORDER BY ViewOrder"
	rs.Open sql, conn, 1, 2
	If rs.recordCount > 0 Then 
        Races = rs.GetRows()
    Else
        ReDim Races(0, 1)
    End If
    rs.Close
	Set rs = Nothing
	
	If CLng(lRaceID) = 0 Then lRaceID = Races(0, 0)

	If CLng(lRaceID) > 0 Then
        iRaceFin = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lRaceID & " AND FnlScnds > 0 and Place > 0"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iRaceFin = rs.RecordCount
        rs.Close
        Set rs = Nothing

		sql = "SELECT RaceDesc, RaceDist, RaceUnits, ScoreMethod, NumScore, IndivRelay, TeamScores, ResultsNotes, ShowResults, Advancement, NumLaps, "
        sql = sql & "OrderBy, StageRace FROM Races WHERE RacesID = " & lRaceID
		Set rs = conn.Execute(sql)
		sRaceName = Replace(rs(0).Value, "''", "'")
		iDist = rs(1).Value
		sUnits = rs(2).Value
		sScoreMethod = rs(3).Value
        iNumScore = rs(4).Value
        sIndivRelay = rs(5).Value
        sTeamScores = rs(6).Value
        If Not rs(7).Value & "" = "" Then sResultsNotes = Replace(rs(7).Value, "''", "'")
        sShowResults = rs(8).Value
        sAdvancement = rs(9).Value
        iNumLaps = rs(10).Value
        sOrderResultsBy = rs(11).Value
		sStageRace = rs(12).Value
		Set rs = Nothing

		sRaceDist = iDist & " " & sUnits

		'see if this race is in a series
		sql = "SELECT s.SeriesID, s.SeriesName FROM Series s INNER JOIN SeriesMeets sm ON s.SeriesID = sm.SeriesID "
		sql = sql & "WHERE sm.RacesID = " & lRaceID
		Set rs = conn.Execute(sql)
        If rs.BOF and rs.EOF Then
            lSeriesID = 0
        Else
            lSeriesID = rs(0).Value
            sSeriesName = Replace(rs(1).Value, "''", "'")
        End If
		Set rs = Nothing

		If sOrderResultsBy = "time" Then
			sql = "SELECT r.LastName, r.FirstName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ir.Excludes, ir.TeamPlace, "
			sql = sql & "ir.Bib, ir.RosterID FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir "
			sql = sql & "ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID 
			sql = sql & " AND ir.Place > 0 AND ir.FnlScnds > 0 ORDER BY ir.Excludes, ir.FnlScnds, ir.Place"
		Else
			sql = "SELECT r.LastName, r.FirstName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ir.Excludes, ir.TeamPlace, "
			sql = sql & "ir.Bib, ir.RosterID FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir "
			sql = sql & "ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID 
			sql = sql & " AND ir.Place > 0 AND ir.FnlScnds > 0 ORDER BY ir.Excludes, ir.Place"
		End If
		Set rs = conn.Execute(sql)
		If rs.BOF and rs.EOF Then
			ReDim RsltsArr(9, 0)
		Else
			RsltsArr = rs.GetRows()
		End If
		Set rs = Nothing

		For i = 0 To UBound(RsltsArr, 2)
			RsltsArr(0, i) = Replace(RsltsArr(0, i), "''", "'")
			RsltsArr(1, i) = Replace(RsltsArr(1, i), "''", "'")
			RsltsArr(2, i) = Replace(RsltsArr(2, i), "''", "'")
			RsltsArr(5, i) = Replace(RsltsArr(5, i), "-", "")
			If RsltsArr(6, i) = "y" Then 
				RsltsArr(7,i) = "---"
			Else
				If CInt(RsltsArr(7, i)) = 0 Then RsltsArr(7,i) = "---"
			End If
		Next

		If sTeamScores = "y" Then
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7, t.TeamsID FROM Teams t INNER JOIN TmRslts tr "
			sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> ''"' ORDER BY Score DESC...commented this out to keep team score order in the event of a tie breaker"
			rs.Open sql, conn, 1, 2
			If rs.RecordCount > 0 Then
				TmRslts = rs.GetRows()
			Else
				ReDim TmRslts(9, 0)
				bTmRsltsReady = False
			End If
			rs.Close
			Set rs = Nothing

			If UBound(TmRslts, 2) > 0 Then
				For i = 0 To UBound(TmRslts, 2) - 1
					For j = i + 1 To UBound(TmRslts, 2)
						If CSng(TmRslts(1, i)) > CSng(TmRslts(1, j)) Then
							For k = 0 To 9
								SortArr(k) = TmRslts(k, i)
								TmRslts(k, i) = TmRslts(k, j)
								TmRslts(k, j) = SortArr(k)
							Next
						End If
					Next
				Next
			End If
		End If
	End If
End If
%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/per_mile_cc.asp" -->
<!--#include file = "../../includes/per_km_cc.asp" -->
<!--#include file = "../../includes/clean_input.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="alternate" href="http://gopherstateevents.com" hreflang="en-us" />
<link rel="shortcut icon" href="/assets/images/g-transparent2-351x345.png" type="image/x-icon">
<link rel="stylesheet" href="/assets/web/assets/mobirise-icons/mobirise-icons.css">
<link rel="stylesheet" href="/assets/tether/tether.min.css">
<link rel="stylesheet" href="/assets/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="/assets/bootstrap/css/bootstrap-grid.min.css">
<link rel="stylesheet" href="/assets/bootstrap/css/bootstrap-reboot.min.css">
<link rel="stylesheet" href="/assets/socicon/css/styles.css">
<link rel="stylesheet" href="/assets/dropdown/css/style.css">
<link rel="stylesheet" href="/assets/theme/css/style.css">
<link rel="stylesheet" href="/assets/mobirise/css/mbr-additional.css" type="text/css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.6.0/css/bootstrap-datepicker.css">

<script src="/assets/web/assets/jquery/jquery.min.js"></script>

<script src="/misc/scripts.js"></script>
<style>.async-hide { opacity: 0 !important} </style>
<script>(function(a,s,y,n,c,h,i,d,e){s.className+=' '+y;h.start=1*new Date;
h.end=i=function(){s.className=s.className.replace(RegExp(' ?'+y),'')};
(a[n]=a[n]||[]).hide=h;setTimeout(function(){i();h.end=null},c);h.timeout=c;
})(window,document.documentElement,'async-hide','dataLayer',4000,
{'GTM-TX6CT27':true});</script>

<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-56760028-1', 'auto');

  ga('require', 'GTM-TX6CT27');

  ga('send', 'pageview');
</script>
<title>Gopher State Events Cross-Country Results for <%=sMeetName%> on <%=dMeetDAte%></title>
<meta name="description" content="Cross-Country Results by Gopher State Events for <%=sMeetName%> on <%=dMeetDate%>">
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="row">
		<div class="col-sm-12">
			<div class="row">
				<div class="col-sm-9">
					<div class="row">
						<a href="http://www.gopherstateevents.com/cc_meet/perf_trkr/create_accnt.asp?part_id=0"
							onclick="openThis(this.href,1024,768);return false;">
							<img src="http://www.gopherstateevents.com/graphics/banner_ads/perf_tracker.png" alt="Performance Tracker" class="img-responsive">
						</a>
					</div>

					<h3 class="h3">GSE Cross-country Results</h3>

					<div class="row">
						<div class="col-sm-8">
							<form role="form" class="form-inline" name="get_races" method="post" action="cc_rslts2.asp" style="margin-bottom: 10px;">
							<div  class="form-group">
								<label for="meets">Meet:&nbsp;</label>
								<select class="form-control" name="meets" id="meets" onchange="this.form.get_meet.click();">
									<option value="0">&nbsp;</option>
									<%For i = 0 to UBound(MeetArray, 2)%>
										<%If CLng(lMeetID) = CLng(MeetArray(0, i)) Then%>
											<option value="<%=MeetArray(0, i)%>" selected><%=MeetArray(1, i)%>&nbsp;(<%=MeetArray(2, i)%>)</option>
										<%Else%>
											<option value="<%=MeetArray(0, i)%>"><%=MeetArray(1, i)%>&nbsp;(<%=MeetArray(2, i)%>)</option>
										<%End If%>
									<%Next%>
								</select>
								<input class="form-control" type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
								<input class="form-control" type="submit" name="get_meet" id="get_meet" value="View This Meet">
							</div>
							</form>
						</div>
						<div class="col-sm-4">
							<%If CLng(lMeetID) > 0 Then%>
								<form role="form" class="form-inline" name="get_races" method="post" action="cc_rslts2.asp?meet_id=<%=lMeetID%>&amp;show_indiv=<%=sShowIndiv%>">
								<div class="form-group">
									<label for="races">Race:&nbsp;</label>
									<select class="form-control" name="races" id="races" onchange="this.form.get_race.click();">
										<option value="0">&nbsp;</option>
										<%For i = 0 to UBound(Races, 2)%>
											<%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
												<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
											<%Else%>
												<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
											<%End If%>
										<%Next%>
									</select>
									<input class="form-control" type="hidden" name="submit_race" id="submit_race" value="submit_race">
									<input class="form-control" type="submit" name="get_race" id="get_race" value="Get Results" style="font-size:0.8em;">
								</div>
								</form>
							<%End If%>
						</div>
					</div>
				</div>
				<div class="col-sm-3">
					<%If Not sLogo & "" = "" Then%>
						<img src="/events/logos/<%=sLogo%>" alt="Logo" class="img-responsive">
					<%End If%>
				</div>
			</div>
			
			<%If CLng(lMeetID) > 0 Then%>
				<%If Not sWeather = vbNullString Then%>
					<p>Weather:&nbsp;<%=sWeather%></p>
				<%End If%>
				
                <%If CDate(Date) < CDate(dMeetDate) + 7 Then%>
			        <%If bRsltsOfficial = False Then%>
				        <div class="bg-danger">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE.  Please report any issues to 
                            bob.schneider@gopherstateevents.com.</div>
			        <%Else%>
				        <div class="bg-success">These results are now official.  If you notice any errors please contact us 
				        via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by telephone (612-720-8427).</div>
			        <%End If%>
                <%End If%>

				<ul class="list-inline">
					<li class="list-inline-item">Total Finishers:&nbsp;<%=iNumFin%></li>
					<li class="list-inline-item">Race Finishers:&nbsp;<%=iRaceFin%></li>
					<li class="list-inline-item">Distance:&nbsp;<%=sRaceDist%></li>
					<li class="list-inline-item">Site/Location:&nbsp;<%=sMeetSite%></li>
				</ul>

				<%If Not sResultsNotes & "" = "" Then%>
					<div class="bg-danger">Results Notes:&nbsp;<%=sResultsNotes%></div>
				<%End If%>

				<%If sShowResults = "y" Then%>
					<ul class="list-inline">
						<li class="list-inline-item list-inline-item-success">
							<a href="javascript:pop('awards.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',1024,650)">Awards</a>
						</li>
						<li class="list-inline-item list-inline-item-success">
							<a href="javascript:pop('rslts_by_team.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',625,650)"
							style="color: red;">Results By Team</a>
						</li>
						<li class="list-inline-item list-inline-item-success">
							<a href="javascript:pop('cc_rslts_grade.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',625,650)">Results By Grade</a>
						</li>
						<li class="list-inline-item list-inline-item-success">
							<a href="javascript:pop('print_rslts.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',1024,650)">Print</a>
						</li>
						<li class="list-inline-item list-inline-item-success">
							<a href="javascript:pop('cc_rslts_cumtime.asp?meet_id=<%=lMeetID%>',1024,650)">Cumulative Time</a>
						</li>
						<%If sTeamScores = "y" Then%>
							<li class="list-inline-item list-inline-item-success">
								<a href="comp_rslts.asp?meet_id=<%=lMeetID%>" onclick="openThis(this.href,1024,768);return false;">Comprehensive</a>
							</li>
						<%End If%>
						<li class="list-inline-item list-inline-item-success">
							<a href="dual_rslts.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>" 
							onclick="openThis(this.href,800,600);return false;">Dual Meet</a>
						</li>
						<li class="list-inline-item list-inline-item-success">
							<a href="combined_scores.asp?meet_id=<%=lMeetID%>" 
							onclick="openThis(this.href,800,600);return false;">Combine Team Scores</a>
						</li>
						<li class="list-inline-item list-inline-item-success">
							<a href="digital_results.asp?meet_id=<%=lMeetID%>" 
							onclick="openThis(this.href,800,600);return false;">Bib Look-Up</a>
						</li>
						<!--
						<%If Not sSeriesName = vbNullString Then%>
							&nbsp;|&nbsp;
							<a href="series_rslts.asp?series_id=<%=lSeriesID%>&amp;gender=<%=sSeriesGender%>" 
								onclick="openThis(this.href,800,600);return false;"><%=sSeriesName%></a>
						<%End If%>
						-->
					</ul>
			
					<br>

					<%If sTeamScores = "y" Then%>
						<h4 class="h4">Team Results: Top <%=iNumScore%> Per Team</h4>

						<table class="table table-striped">
							<tr>
								<th>Pl</th>
								<th>Team</th>
								<th>Score</th>
								<th>R1</th>
								<th>R2</th>
								<th>R3</th>
								<th>R4</th>
								<th>R5</th>
								<th>R6</th>
								<th>R7</th>
							</tr>
							<%For i = 0 to UBound(TmRslts, 2)%>
								<tr>
									<td><%=i + 1%></td>
									<td style="text-align:left;white-space:nowrap;"><%=TmRslts(0, i)%></td>
									<td><%=TmRslts(1, i)%></td>
									<td><%=TmRslts(2, i)%></td>
									<td><%=TmRslts(3, i)%></td>
									<td><%=TmRslts(4, i)%></td>
									<td><%=TmRslts(5, i)%></td>
									<td><%=TmRslts(6, i)%></td>
									<td><%=TmRslts(7, i)%></td>
									<td><%=TmRslts(8, i)%></td>
								</tr>

								<%If sShowIndiv = "y" Then%>
									<tr>
										<td style="padding-left: 50px;" colspan="10">
											<%If bTmRsltsReady = True Then%>
												<%Call GetIndiv(TmRslts(9, i))%>

												<%If UBound(OurRslts, 2) > 0 Then%>
													<table  class="table table-condensed">
														<%For j = 0 to UBound(OurRslts, 2)%>
															<tr>
																<td><%=TmRslts(j + 2, i)%></td>
																<td><%=OurRslts(1, j)%>-<%=OurRslts(3, j)%>, <%=OurRslts(2, j)%></td>
																<td><%=OurRslts(4, j)%></td>
																<td><%=OurRslts(5, j)%></td>
															</tr>

															<%If j = 6 Then Exit For%>
														<%Next%>
													</table>
												<%End If%>
											<%End If%>
										</td>
									</tr>
								<%End If%>
							<%Next%>
						</table>
					<%End If%>

					<h4 class="h4">Individual Results</h4>

					<table class="table table-striped">
						<tr>
							<th>Pl</th>
							<th>Tm</th>
							<th>Bib-Name</th>
							<th>Team</th>
							<th>Gr</th>
							<th>M/F</th>
							<th>Time</th>
							<th>Per Mi</th>
							<th>Per Km</th>
						</tr>
						<%k = 1%>
						<%For i = 0 to UBound(RsltsArr, 2)%>
							<tr>
								<td>
									<%If RsltsArr(6, i) = "y" Then%>
										-
									<%Else%>
										<%=k%>
										<%k = k + 1%>
									<%End If%>
								</td>
								<td><%=RsltsArr(7, i)%></td>
								<td><%=RsltsArr(8, i)%> - <%=RsltsArr(0, i)%>, <%=RsltsArr(1, i)%></td>
								<td><%=RsltsArr(2, i)%></td>
								<td><%=RsltsArr(3, i)%></td>
								<td><%=RsltsArr(4, i)%></td>
								<td><%=RsltsArr(5, i)%></td>
								<td><%=PacePerMile(RsltsArr(5, i), iDist, sUnits)%></td>
								<td><%=PacePerKM(RsltsArr(5, i), iDist, sUnits)%></td>
							</tr>
						<%Next%>
					</table>
				</div>
            <%End If%>
        <%End If%>
	</div>	
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
