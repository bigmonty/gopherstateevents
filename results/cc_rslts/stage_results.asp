<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs
Dim i
Dim lMeetID, lRaceID, lSeriesID
Dim sClickPage, sSport, sGradeYear, sMeetSite, sLogo, sWeather, sShowPix, sRaceName, sUnits, sScoreMethod
Dim sIndivRelay, sTeamScores, sResultsNotes, sShowResults, sAdvancement, sTechnique, sOrderResultsBy
Dim sMeetName, sSeriesName
Dim iNumFin, iDist, iNumScore, iNumLaps
Dim MeetArray, Races, MeetTeams
Dim dMeetDate
Dim bRsltsOfficial
Dim cdoConfig

'advancement variables
Dim lMyRosterID
Dim iNumAdvance, iNumTeams, iIndAdvTtl
Dim sExcludeTeam, sAdvanceTo, sAdvance
Dim AdvTeams()

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

sSport = Request.QueryString("sport")
If sSport = vbNullString Then sSport = "cc"

If sSport = "cc" or sSport = "Cross-Country" Then
    sSport = "Cross-Country"
Else
    sSport = "Nordic Ski"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate <= '" & Now() & "' AND ShowOnline = 'y' AND Sport = '" & sSport 
sql = sql & "' ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
MeetArray = rs.GetRows()
Set rs = Nothing

If Request.Form.Item("submit_meet") = "submit_meet" Then
	lMeetID = Request.Form.Item("meets")
    If CStr(lMeetID) = vbNullString Then lMeetID = 0
    If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
    If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
End If

If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect "http://www.google.com"

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

	sql = "SELECT MeetName, MeetDate, MeetSite, Weather, Sport, Logo, ShowPix FROM Meets WHERE MeetsID = " & lMeetID
	Set rs = conn.Execute(sql)
	sMeetName = rs(0).Value
	dMeetDate = rs(1).Value
	If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then sWeather = Replace(rs(3).Value, "''", "'")
    sSport = rs(4).Value
    sLogo = rs(5).Value
    sShowPix = rs(6).Value
	Set rs = Nothing
	
	'get year for roster grades
	If Month(dMeetDate) <= 7 Then
	    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
	Else
	    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
	End If

    If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

	sql = "SELECT t.TeamsID, t.TeamName, t.Gender FROM Teams t INNER JOIN MeetTeams mt ON t.TeamsID = mt.TeamsID "
	sql = sql & "WHERE mt.MeetsID = " & lMeetID & " ORDER BY t.TeamName"
	Set rs = conn.Execute(sql)
	MeetTeams = rs.GetRows()
	Set rs = Nothing
	
	sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID & " ORDER BY ViewOrder"
	Set rs = conn.Execute(sql)
	Races = rs.GetRows()
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
        sql = sql & "Technique, OrderBy FROM Races WHERE RacesID = " & lRaceID
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
        sTechnique = rs(11).Value
        sOrderResultsBy = rs(12).Value
		Set rs = Nothing

        If sAdvancement = "y" Then
            sql = "SELECT NumAdvance, NumTeams, ExcludeTeam, AdvanceTo FROM Advancement WHERE RacesID = " & lRaceID
            Set rs = conn.Execute(sql)
            iNumAdvance = rs(0).Value
            iNumTeams = rs(1).Value
            sExcludeTeam = rs(2).Value
            sAdvanceTo = rs(3).Value
            Set rs = Nothing
        End If

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
	End If
End If

Private Function MeetName()
    sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lMeetID
    Set rs = conn.Execute(sql)
    MeetName = Replace(rs(0).Value, "''", "'")
    Set rs = Nothing
End Function

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/per_mile_cc.asp" -->
<!--#include file = "../../includes/per_km_cc.asp" -->
<!--#include file = "../../includes/clean_input.asp" -->
<%

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events Cross-Country Results</title>
<meta name="description" content="Cross-Country & Nordic Ski Results by Gopher State Events, a conventional timing service offererd by H51 Software, LLC in Minnetonka, MN.">
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="row">
		<div class="col-sm-10">
            <a href="http://www.gopherstateevents.com/cc_meet/perf_trkr/create_accnt.asp?part_id=0"
                onclick="openThis(this.href,1024,768);return false;">
                <img src="http://www.gopherstateevents.com/graphics/banner_ads/perf_tracker.png" alt="Performance Tracker" class="img-responsive">
            </a>

			<hr>

		    <h3 class="h3">GSE <%=sSport%> Results</h3>

			<form role="form" class="form-inline" name="get_races" method="post" action="cc_rslts.asp?sport=<%=sSport%>" style="margin-bottom: 10px;">
			<div  class="form-group">
				<label for="meets">Select Meet:</label>
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
			
			<%If Not CLng(lMeetID) = 0 Then%>
				<h4 class="h4">Results for <%=sMeetName%> on <%=dMeetDate%></h4>

				<%If Not sWeather = vbNullString Then%>
					<p>Weather:&nbsp;<%=sWeather%></p>
				<%End If%>
				
				<%If CDate(Date) < CDate(dMeetDate) + 7 Then%>
					<%If bRsltsOfficial = False Then%>
						<p class="bg-danger">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE.  Please report any issues to 
							bob.schneider@gopherstateevents.com.</p>
					<%Else%>
						<p class="bg-success">These results are now official.  If you notice any errors please contact us 
						via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by telephone (612-720-8427).</p>
					<%End If%>
				<%End If%>

				<form role="form" class="form-inline" name="get_races" method="post" action="cc_rslts.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;rslts_page=<%=sRsltsPage%>&amp;show_indiv=<%=sShowIndiv%>">
				<div class="form-group">
					<label for="races">Select Race:</label>
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
				
				<%If sSport = "Nordic Ski" and sTechnique <> "" Then%>
					<h5 class="h5">Technique: <%=sTechnique%></h5>
				<%Else%>
					<br>
				<%End If%>

				<ul class="list-inline">
					<li class="list-group-item">Total Finishers:&nbsp;<%=iNumFin%></li>
					<li class="list-group-item">Race Finishers:&nbsp;<%=iRaceFin%></li>
					<li class="list-group-item">Distance:&nbsp;<%=sRaceDist%></li>
					<li class="list-group-item">Site/Location:&nbsp;<%=sMeetSite%></li>
				</ul>

				<%If Not sResultsNotes & "" = "" Then%>
					<p class="bg-danger">Results Notes:&nbsp;<%=sResultsNotes%></p>
				<%End If%>

				<%If sShowResults = "y" Then%>
				<%End If%>
			<%End If%>
		</div>
        <div class="col-sm-2">
            <%If Not sLogo & "" = "" Then%>
                <img src="/events/logos/<%=sLogo%>" alt="Logo" class="img-responsive">
            <%End If%>
            <%'If sShowPix = "y" Or Session("role") = "admin" Then%>
                <div style="margin:0;padding:0;text-align:center;">
                    <%If UBound(RaceGallery) = 0 Then%>
                        <%If Date < CDate(dMeetDate) + 10 Then%>
                            <img src="/graphics/no_pix.png" alt="Pix Not Available Yet" class="img-responsive">
                        <%End If%>
                    <%Else%>
                        <%For i = 0 To UBound(RaceGallery) - 1%>
                            <a href="<%=RaceGallery(i)%>" onclick="openThis(this.href,1024,768);return false;">
                                <img src="/graphics/Camera-icon.png" alt="Race Photos" class="img-responsive" style="margin: 0;padding: 0;">
                            </a>
                        <%Next%>
                    <%End If%>
                </div>
            <%'End If%>
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
<!--#include file = "../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing

conn.Close
Set conn = Nothing
%>
</body>
</html>
