<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs, rs2, sql2
Dim i, j, k, m
Dim lMeetID, lRaceID, lSeriesID, lRosterID, lWhichTeam, lPrelimID, lFeaturedEventsID
Dim sRaceName, sSeriesName, sSeriesGender, sMeetName, sGradeYear, sOrderResultsBy, sScoreMethod, sRsltsPage, sTeamName, sUnits, sMeetSite, sWeather
Dim sRaceDist, sSport, sIndivRelay, sTeamScores, sResultsNotes, sErrMsg, sLogo, sShowIndiv, sClickPage, sBannerImage, sShowResults, sShowPix, sTechnique
Dim sIsStage, sRaceStatus
Dim iDist, iNumScore, iNumFin, iRaceFin, iNumLaps, iNumPrelims
Dim RsltsArr, TmRslts, TempArr(9), TmRslts4(), MeetTeams, SortArr(9), MeetArray, Races, RaceGallery(), RankArr(), OurRslts
Dim dMeetDate
Dim bRsltsOfficial, bTmRsltsReady

'advancement variables
Dim lMyRosterID
Dim iNumAdvance, iNumTeams, iIndAdvTtl
Dim sExcludeTeam, sAdvanceTo, sAdvance, sAdvancement
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

sShowIndiv = Request.QueryString("show_indiv")
If sShowIndiv = vbNullString Then sShowIndiv = "n"
bTmRsltsReady = True

lRaceID = Request.QueryString("race_id")
lWhichTeam = Request.QueryString("which_team")
sRsltsPage = Request.QueryString("rslts_page")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate <= '" & Now() & "' AND ShowOnline = 'y' AND Sport = '" & sSport 
sql = sql & "' ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
MeetArray = rs.GetRows()
Set rs = Nothing
If Request.Form.Item("get_team") = "get_team" Then
	lWhichTeam = Request.Form.Item("teams")
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_meet") = "submit_meet" Then
	lMeetID = Request.Form.Item("meets")
    If CStr(lMeetID) = vbNullString Then lMeetID = 0
    If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
    If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	sRsltsPage = Request.Form.Item("which_rslts")
ElseIf Request.Form.Item("get_team") = "get_team" Then
	lWhichTeam = Request.Form.Item("which_team")
End If

If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect "http://www.google.com"

If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect "http://www.google.com"

If CStr(lWhichTeam) = vbNullString Then lWhichTeam = 0
If Not IsNumeric(lWhichTeam) Then Response.Redirect "http://www.google.com"

If sRsltsPage = vbNullString Then sRsltsPage = "overall_rslts.asp"

ReDim TmRslts4(5, 0)
ReDim RaceGallery(0)

If Not CLng(lMeetID) = 0 Then
    iNumFin = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID FROM IndRslts WHERE MeetsID = " & lMeetID & " AND FnlScnds > 0 and Place > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iNumFin = rs.RecordCount
    rs.Close
    Set rs = Nothing

    i = 0
    sql = "SELECT EmbedLink FROM RaceGallery WHERE MeetsID = " & lMeetID
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        RaceGallery(i) = rs(0).Value
        i = i + 1
        ReDim Preserve RaceGallery(i)
        rs.MoveNext
    Loop
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

	If Not CLng(lRaceID) = 0 Then
        iRaceFin = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lRaceID & " AND FnlScnds > 0 and Place > 0"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iRaceFin = rs.RecordCount
        rs.Close
        Set rs = Nothing

		sql = "SELECT RaceDesc, RaceDist, RaceUnits, ScoreMethod, NumScore, IndivRelay, TeamScores, ResultsNotes, ShowResults, Advancement, NumLaps, "
        sql = sql & "Technique, OrderBy, StageRace FROM Races WHERE RacesID = " & lRaceID
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
        sIsStage = rs(13).Value
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

        If sIsStage = "y" Then
            'see if it is a destination race
            Dim lGroupRacesID, lRaceGroupsID

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT GroupRacesID, RaceGroupsID FROM GroupRaces WHERE RacesID = " & lRaceID
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then 
                lGroupRacesID = rs(0).Value
                lRaceGroupsID = rs(1).Value
            ENd If
            rs.Close
            Set rs = Nothing

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT RaceStatus FROM GroupRaces WHERE GroupRacesID = " & lGroupRacesID
            rs.Open sql, conn, 1, 2
            If Not rs(0).Value & "" = "" Then sRaceStatus = rs(0).Value
            rs.Close
            Set rs = Nothing

            'get num prelims
            i = 0
            Dim Prelims()
            ReDim Prelims(2, 0)
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT GroupRacesID, RaceStatus, RacesID FROM GroupRaces WHERE RaceGroupsID = " & lRaceGroupsID & " AND RaceStatus <> 'Destination'"
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then iNumPrelims = rs.RecordCount
            Do While Not rs.EOF
                Prelims(0, i) = rs(0).Value
                Prelims(1, i) = rs(1).Value
                Prelims(2, i) = rs(2).Value
                i = i + 1
                ReDim Preserve Prelims(2, i)
                rs.MoveNext
            Loop
            rs.Close
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
	
		If sRsltsPage = "overall_rslts.asp" Then
            If sScoreMethod = "Pursuit" Then
                Dim sPursuitName, sTechnique1, sTechnique2
                Dim sPursuitDist, sPrelimDist, dPrelimDate
                Dim sPrelimMeet

	            'get pursuit data
	            sql = "SELECT RaceDist, RaceUnits, RaceDesc, RaceName, Technique FROM Races WHERE RacesID = " & lRaceID
	            Set rs = conn.Execute(sql)
                If rs.BOF and rs.EOF Then
	                '--
                Else
	                sPursuitDist = rs(0).Value & " " & rs(1).Value
	                sTechnique2 = rs(4).Value
	                sPursuitName = rs(3).Value
                End If
	            Set rs = Nothing
	      
	            'get prelim race id
	            sql = "SELECT PrelimRace FROM Pursuit WHERE RacesID = " & lRaceID
	            Set rs = conn.Execute(sql)
                If rs.BOF and rs.EOF Then
	                lPrelimID = 0
                Else
	                lPrelimID = rs(0).Value
                End If
	            Set rs = Nothing
	      
	            'get prelim race info
	            sql = "SELECT r.RaceDist, r.RaceUnits, m.MeetDate, m.MeetName, r.Technique FROM Races r INNER JOIN Meets m "
	            sql = sql & "ON r.MeetsID = m.MeetsID WHERE RacesID = " & lPrelimID
	            Set rs = conn.Execute(sql)
                If rs.BOF and rs.EOF Then
	                '--
                Else
	                sPrelimDist = rs(0).Value & " " & rs(1).Value
	                dPrelimDate = rs(2).Value
	                sPrelimMeet = rs(3).Value
                    sTechnique1 = rs(4).Value
                End If
	            Set rs = Nothing

 	            i = 0
	            ReDim TmRslts(9, 0)
	            sql = "SELECT t.TeamsID, t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
	            sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID
	            Set rs = conn.Execute(sql)
	            Do While Not rs.EOF
	                For j = 0 To 9
	                    TmRslts(j, i) = Trim(rs(j).Value)
	                Next
	                i = i + 1
	                ReDim Preserve TmRslts(9, i)
	                rs.MoveNext
	            Loop
	            Set rs = Nothing
	
	            'sort the arrays
	            For i = 0 To UBound(TmRslts, 2) - 2
	                For j = i + 1 To UBound(TmRslts, 2) - 1
	                    If ConvertToSeconds(TmRslts(2, i)) < ConvertToSeconds(TmRslts(2, j)) Then
	                        For k = 0 To 9
	                            TempArr(k) = TmRslts(k, i)
	                            TmRslts(k, i) = TmRslts(k, j)
	                            TmRslts(k, j) = TempArr(k)
	                        Next
	                    End If
	                Next
	            Next
	        
	            'get a finishers array for this race
	            i = 0
	            ReDim PursuitIndRslts(8, 0)
	            sql = "SELECT r.RosterID, r.FirstName, r.LastName, t.TeamName, ir.RaceTime, ir.Bib FROM IndRslts ir INNER JOIN Roster r"
	            sql = sql & " ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID "
	            sql = sql & " INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
	            sql = sql & "WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' AND ir.Excludes = 'n' "
	            sql = sql & "ORDER BY ir.Place, ir.FnlScnds"
	            Set rs = conn.Execute(sql)
	            Do While Not rs.EOF
	                PursuitIndRslts(0, i) = rs(0).Value                 'name
	                PursuitIndRslts(1, i) = rs(5).Value & "-" & Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	                PursuitIndRslts(2, i) = GetGrade(rs(0).Value)       'grade
	                PursuitIndRslts(3, i) = rs(3).Value                 'team
	                PursuitIndRslts(8, i) = ConvertToMinutes(ConvertToSeconds(rs(4).Value)) 'pursuit time...conversions are to eliminate decimal places
	                
	                i = i + 1
	                ReDim Preserve PursuitIndRslts(8, i)
	                rs.MoveNext
	            Loop
	            Set rs = Nothing
	        
	            'create rank array
	            ReDim RankArr(2, UBound(PursuitIndRslts, 2))
	            For i = 0 To UBound(PursuitIndRslts, 2) - 1
	                RankArr(0, i) = PursuitIndRslts(0, i)
	                RankArr(1, i) = PursuitIndRslts(8, i)
	            Next
	        
	            'sort the rank array
	            For i = 0 To UBound(RankArr, 2) - 2
	                For j = i + 1 To UBound(RankArr, 2) - 1
	                    If ConvertToSeconds(RankArr(1, i)) > ConvertToSeconds(RankArr(1, j)) Then
	                        For k = 0 To 1
	                            TempArr(k) = RankArr(k, i)
	                            RankArr(k, i) = RankArr(k, j)
	                            RankArr(k, j) = TempArr(k)
	                        Next
	                    End If
	                Next
	            Next
	        
	            For i = 0 To UBound(RankArr, 2) - 1
	                RankArr(2, i) = i + 1
	            Next
	        
	            'sort the rank array
	            For i = 0 To UBound(PursuitIndRslts, 2) - 1
	                For j = 0 To UBound(RankArr, 2) - 1
	                    If CLng(PursuitIndRslts(0, i)) = CLng(RankArr(0, j)) Then
	                        PursuitIndRslts(7, i) = RankArr(2, j)
	                        Exit For
	                    End If
	                Next
	            Next
	        
	            'now get prelim time and rank
	            i = 1
	            sql = "SELECT RosterID, RaceTime FROM IndRslts WHERE RacesID = " & lPrelimID & " AND RaceTime > '00:00' AND Place > 0 ORDER BY FnlScnds"
	            Set rs = conn.Execute(sql)
	            Do While Not rs.EOF
	                For j = 0 To UBound(PursuitIndRslts, 2) - 1
	                    If CLng(PursuitIndRslts(0, j)) = CLng(rs(0).Value) Then
	                        PursuitIndRslts(4, j) = ConvertToMinutes(Round(ConvertToSeconds(rs(1).Value) + ConvertToSeconds(PursuitIndRslts(8, j)), 2))
	                        PursuitIndRslts(5, j) = i
	                        PursuitIndRslts(6, j) = ConvertToMinutes(Round(ConvertToSeconds(rs(1).Value), 0))
	                        i = i + 1
	                    End If
	                Next
	                rs.MoveNext
	            Loop
	            Set rs = Nothing
           Else
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
'                    If Not RsltsArr(3, i) & "" = "" Then RsltsArr(3, i) = GetGrade(RsltsArr(3,i))
                    If RsltsArr(6, i) = "y" Then 
                        RsltsArr(7,i) = "---"
                    Else
                        If CInt(RsltsArr(7, i)) = 0 Then RsltsArr(7,i) = "---"
                    End If
                Next

                If sTeamScores = "y" Then
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    If sSport = "Cross-Country" Then
			            sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7, t.TeamsID FROM Teams t INNER JOIN TmRslts tr "
			            sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> ''"' ORDER BY Score DESC...commented this out to keep team score order in the event of a tie breaker"
                    Else
			            sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7, t.TeamsID FROM Teams t INNER JOIN TmRslts tr "
			            sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> '' ORDER BY Score"
                    End If
			        rs.Open sql, conn, 1, 2
                    If rs.RecordCount > 0 Then
	                    TmRslts = rs.GetRows()
                    Else
	                    ReDim TmRslts(9, 0)
                        bTmRsltsReady = False
                    End If
                    rs.Close
			        Set rs = Nothing

                    If sSport = "Cross-Country" Then
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
                    Else
                        For i = 0 To UBound(TmRslts, 2) - 1
                            For j = i + 1 To UBound(TmRslts, 2)
                                If CSng(TmRslts(1, i)) < CSng(TmRslts(1, j)) Then
                                    For k = 0 To 9
                                        SortArr(k) = TmRslts(k, i)
                                        TmRslts(k, i) = TmRslts(k, j)
                                        TmRslts(k, j) = SortArr(k)
                                    Next
                                End If
                            Next
                        Next
	
			            i = 0
			            ReDim TmRslts4(5, 0)
                        If CInt(iNumScore) > 4 Then
			                sql = "SELECT t.TeamName, tr.R1, tr.R2, tr.R3, tr.R4 FROM Teams t INNER JOIN TmRslts tr "
			                sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> '' ORDER BY Score"
			                Set rs = conn.execute(sql)
                            If rs.BOF and rs.EOF Then
	                            '--
                            Else
			                    Do While Not rs.EOF
			                        TmRslts4(0, i) = rs(0).Value
                                    TmRslts4(1, i) = "0"
			                        For j = 1 To 4
                                        If IsNumeric(Trim(rs(j).Value)) Then TmRslts4(1, i) = CSng(TmRslts4(1, i)) + 101-CSng(Trim(rs(j).Value))
                                    Next
                                    TmRslts4(2, i) = Trim(rs(1).Value)
			                        TmRslts4(3, i) = Trim(rs(2).Value)
			                        TmRslts4(4, i) = Trim(rs(3).Value)
			                        TmRslts4(5, i) = Trim(rs(4).Value)
				                    i = i + 1
				                    ReDim Preserve TmRslts4(5, i)
				                    rs.MoveNext
			                    Loop
                            End If
			                Set rs = Nothing

                             For i = 0 To UBound(TmRslts4, 2) - 2
                                For j = i + 1 To UBound(TmRslts4, 2) - 1
                                    If CSng(TmRslts4(1, i)) < CSng(TmRslts4(1, j)) Then
                                        For k = 0 To 5
                                            SortArr(k) = TmRslts4(k, i)
                                            TmRslts4(k, i) = TmRslts4(k, j)
                                            TmRslts4(k, j) = SortArr(k)
                                        Next
                                    End If
                                Next
                            Next
                       End If
                    End If
                End If
            End If
		Else
			If Not CLng(lWhichTeam) = 0 Then
				sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lWhichTeam
				Set rs = conn.Execute(sql)
				sTeamName = rs(0).Value & " (" & rs(1).Value & ")"
				Set rs = Nothing
				
                If sOrderResultsBy = "time" Then
				    sql = "SELECT ir.Bib, r.LastName, r.FirstName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime FROM IndRslts ir "
				    sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
				    sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lRaceID & " AND t.TeamsID = " 
				    sql = sql & lWhichTeam & " AND ir.Place > 0 AND ir.FnlScnds > 0 ORDER BY ir.Excludes, ir.FnlScnds, ir.Place"
                Else
				    sql = "SELECT ir.Bib, r.LastName, r.FirstName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime FROM IndRslts ir "
				    sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
				    sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lRaceID & " AND t.TeamsID = " 
				    sql = sql & lWhichTeam & " AND ir.Place > 0 AND ir.FnlScnds > 0 ORDER BY ir.Excludes, ir.Place"
                End If
				Set rs = conn.Execute(sql)
                If rs.BOF and rs.EOF Then
                    ReDim RsltsArr(5, 0)
                Else
                    RsltsArr = rs.GetRows()
                End If
			    Set rs = Nothing

                For i = 0 To UBound(RsltsArr, 2)
                    RsltsArr(1, i) = Replace(RsltsArr(1, i), "''", "'")
                    RsltsArr(2, i) = Replace(RsltsArr(2, i), "''", "'")
                    RsltsArr(5, i) = Replace(RsltsArr(5, i), "-", "")
                Next
			End If
		End If
	End If
End If

Function GetPlace(lRosterID)
	GetPlace = 0
	If sOrderResultsBy = "time" Then
        sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lRaceID & " AND Place > 0 ORDER BY FnlScnds"
    Else
        sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lRaceID & " AND Place > 0 ORDER BY Place"
    End If
	Set rs2 = conn.Execute(sql2)
	Do While Not rs2.EOF
		GetPlace = GetPlace + 1
		If CLng(rs2(0).Value) = CLng(lRosterID) Then Exit Do
		rs2.MoveNext
	Loop
	Set rs2 = Nothing
End Function

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/per_mile_cc.asp" -->
<!--#include file = "../../includes/per_km_cc.asp" -->
<!--#include file = "../../includes/clean_input.asp" -->
<%
	
Private Function GetGrade(lMyID)
	sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
	Set rs2 = conn.Execute(sql2)
    If rs2.BOF and rs2.EOF Then
        GetGrade = 0
    Else
        GetGrade = rs2(0).Value
    End If
	Set rs2 = Nothing
End Function

Private Sub GetIndiv(lThisTeam)
	sql = "SELECT r.RosterID, ir.Bib, r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.RaceTime FROM IndRslts ir INNER JOIN Roster r"
	sql = sql & " ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID "
	sql = sql & "WHERE ir.RacesID = " & lRaceID & " AND r.TeamsID = " & lThisTeam & " AND ir.Place > 0 AND ir.FnlScnds > 0 AND ir.Excludes = 'n' "
	sql = sql & "ORDER BY ir.FnlScnds, ir.Place"
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
	    ReDim OurRslts(5, 0)
    Else
	    OurRslts = rs.GetRows()
    End If
    Set rs = Nothing
End Sub

Private Function GetPrelimTime(lThisPart, lThisRace)
    GetPrelimTime = vbNullString

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql ="SELECT RaceTime FROM IndRslts WHERE RosterID = " & lThisPart & " AND RacesID = " & lThisRace
    rs.Open sql, conn, 1,2
    If rs.RecordCount > 0 Then GetPrelimTime = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function
%>
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
<title>Gopher State Events Results for <%=sMeetName%> on <%=dMeetDAte%></title>
<meta name="description" content="<%=sSport%> Results by Gopher State Events for <%=sMeetName%> on <%=dMeetDate%>">

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

					<h3 class="h3">GSE <%=sSport%> Results</h3>

					<div class="row">
						<div class="col-sm-8">
							<form role="form" class="form-inline" name="get_races" method="post" action="cc_rslts.asp?sport=<%=sSport%>" style="margin-bottom: 10px;">
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
								<form role="form" class="form-inline" name="get_races" method="post" action="cc_rslts.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;rslts_page=<%=sRsltsPage%>&amp;show_indiv=<%=sShowIndiv%>">
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
					<%'If sShowPix = "y" Or Session("role") = "admin" Then%>
						<div style="margin:0;padding:0;text-align:center;">
							<%If UBound(RaceGallery) > 0 Then%>
								<%For i = 0 To UBound(RaceGallery) - 1%>
									<a href="<%=RaceGallery(i)%>" onclick="openThis(this.href,1024,768);return false;">
										<img src="/graphics/Camera-icon.png" alt="Race Photos" class="img-responsive" style="margin: 0;padding: 0;">
									</a>
								<%Next%>
							<%End If%>
						</div>
					<%'End If%>
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
					<%If sSport = "Nordic Ski" and sTechnique <> "" Then%>
						<li class="list-inline-item">Technique:&nbsp;<%=sTechnique%></li>
					<%End If%>
					<li class="list-inline-item">Total Finishers:&nbsp;<%=iNumFin%></li>
					<li class="list-inline-item">Race Finishers:&nbsp;<%=iRaceFin%></li>
					<li class="list-inline-item">Distance:&nbsp;<%=sRaceDist%></li>
					<li class="list-inline-item">Site/Location:&nbsp;<%=sMeetSite%></li>
				</ul>

				<%If Not sResultsNotes & "" = "" Then%>
					<div class="bg-danger">Results Notes:&nbsp;<%=sResultsNotes%></div>
				<%End If%>

				<%If sShowResults = "y" Then%>
					<%If sScoreMethod="Pursuit" Then%>
						<ul class="list-inline">
							<li class="list-inline-item">
								<a href="javascript:pop('rslts_by_team.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',625,650)"
								style="color: red;">Results By Team</a>
							</li>
							<li class="list-inline-item">
								<a href="pursuit_results.asp?meet_id=<%=lMeetID%>" onclick="openThis(this.href,1024,768);return false;">Download</a>
							</li>
							<li class="list-inline-item">
								<a href="digital_results.asp?meet_id=<%=lMeetID%>" 
								onclick="openThis(this.href,800,600);return false;">Bib Look-Up</a>
							</li>
						</ul>
							
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
							<%For i = 0 to UBound(TmRslts, 2) - 1%>
								<tr>
									<td><%=i + 1%></td>
									<td><%=TmRslts(1, i)%></td>
									<td><%=TmRslts(2, i)%></td>
									<td><%=TmRslts(3, i)%></td>
									<td><%=TmRslts(4, i)%></td>
									<td><%=TmRslts(5, i)%></td>
									<td><%=TmRslts(6, i)%></td>
									<td><%=TmRslts(7, i)%></td>
									<td><%=TmRslts(8, i)%></td>
									<td><%=TmRslts(9, i)%></td>
								</tr>
							<%Next%>
						</table>

						<h4 class="h4">Individual Results</h4>

						<table class="table table-striped">
							<tr>
								<th rowspan="2" valign="bottom">Place</th>
								<th rowspan="2" valign="bottom">Bib-Name</th>
								<th rowspan="2" valign="bottom">Gr</th>
								<th rowspan="2" valign="bottom">Team</th>
								<th rowspan="2" valign="bottom">Comb Time</th>
								<th colspan="2"><%=sTechnique1%></th>
								<th colspan="2"><%=sTechnique2%></th>
							</tr>
							<tr>
								<th>Rank</th>
								<th>Time</th>
								<th>Rank</th>
								<th>Time</th>
							</tr>
							<%For i = 0 To UBound(PursuitIndRslts, 2) - 1%>
								<tr>
									<td><%=i + 1%>)</td>
									<td><%=PursuitIndRslts(1, i)%></td>
									<td><%=PursuitIndRslts(2, i)%></td>
									<td><%=PursuitIndRslts(3, i)%></td>
									<td><%=PursuitIndRslts(4, i)%></td>
									<td><%=PursuitIndRslts(5, i)%></td>
									<td><%=PursuitIndRslts(6, i)%></td>
									<td><%=PursuitIndRslts(7, i)%></td>
									<td><%=PursuitIndRslts(8, i)%></td>
								</tr>
							<%Next%>
						</table>
					<%Else%>
						<ul class="list-inline">
							<li class="list-inline-item list-inline-item-success">
								<a href="javascript:pop('awards.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>&amp;which_team=<%=lWhichTeam%>',1024,650)">Awards</a>
							</li>
							<li class="list-inline-item list-inline-item-success">
								<a href="javascript:pop('rslts_by_team.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',625,650)"
								style="color: red;">Results By Team</a>
							</li>
							<li class="list-inline-item list-inline-item-success">
								<a href="javascript:pop('cc_rslts_grade.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',625,650)">Results By Grade</a>
							</li>
							<li class="list-inline-item list-inline-item-success">
								<a href="javascript:pop('print_rslts.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>&amp;which_team=<%=lWhichTeam%>',1024,650)">Print</a>
							</li>
							<li class="list-inline-item list-inline-item-success">
								<a href="javascript:pop('cc_rslts_cumtime.asp?meet_id=<%=lMeetID%>',1024,650)">Cumulative Time</a>
							</li>
							<%If CInt(iNumLaps) > 1 Then%>
								<li class="list-inline-item list-inline-item-info">
									<a href="rslts_by_lap.asp?meets_id=<%=lMeetID%>&amp;races_id=<%=lRaceID%>" onclick="openThis(this.href,1024,768);return false;">Results By Lap</a>
								</li>
								<li class="list-inline-item list-inline-item-info">
									<a href="rslts_w_laps.asp?meets_id=<%=lMeetID%>&amp;races_id=<%=lRaceID%>" onclick="openThis(this.href,1024,768);return false;">Results w/Lap Times</a>
								</li>
							<%End If%>
							<%If sTeamScores = "y" Then%>
								<li class="list-inline-item list-inline-item-success">
									<a href="comp_rslts.asp?meet_id=<%=lMeetID%>" onclick="openThis(this.href,1024,768);return false;">Comprehensive</a>
								</li>
							<%End If%>
							<%If Not sRsltsPage = "overall_rslts.asp" Then%>
								<li class="list-inline-item list-inline-item-success">
									<a href="dwnld_overall.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>" 
										onclick="openThis(this.href,800,600);return false;">Download</a>
								</li>
							<%End If%>
							<%If sSport = "Cross-Country" Then%>
								<li class="list-inline-item list-inline-item-success">
									<a href="dual_rslts.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>" 
									onclick="openThis(this.href,800,600);return false;">Dual Meet</a>
								</li>
							<%End If%>
							<%If CLng(lMeetID) = 452 Then%>
								<li class="list-inline-item list-inline-item-success">
									<a href="girls_combined_18.pdf" 
									onclick="openThis(this.href,800,600);return false;">Girls Combined</a>
								</li>
								<li class="list-inline-item list-inline-item-success">
									<a href="boys_combined_18.pdf" 
									onclick="openThis(this.href,800,600);return false;">Boys Combined</a>
								</li>
							<%ElseIf CLng(lMeetID) = 403 Then%>
								<li class="list-inline-item list-inline-item-success">
									<a href="girls_combined.pdf" 
									onclick="openThis(this.href,800,600);return false;">Girls Combined</a>
								</li>
								<li class="list-inline-item list-inline-item-success">
									<a href="boys_combined.pdf" 
									onclick="openThis(this.href,800,600);return false;">Boys Combined</a>
								</li>
							<%Else%>
								<li class="list-inline-item list-inline-item-success">
									<a href="combined_scores.asp?meet_id=<%=lMeetID%>" 
									onclick="openThis(this.href,800,600);return false;">Combine Team Scores</a>
								</li>
							<%End If%>
							<li class="list-inline-item list-inline-item-success">
								<a href="digital_results.asp?meet_id=<%=lMeetID%>" 
								onclick="openThis(this.href,800,600);return false;">Bib Look-Up</a>
							</li>
							<!--
							&nbsp;|&nbsp;
							<a href="javascript:pop('/misc/under_const.htm',550,200)">Records</a>
							&nbsp;|&nbsp;
							<a href="javascript:pop('/misc/under_const.htm',550,200)">Performance List</a>
							<%If Not sSeriesName = vbNullString Then%>
								&nbsp;|&nbsp;
								<a href="series_rslts.asp?series_id=<%=lSeriesID%>&amp;gender=<%=sSeriesGender%>" 
									onclick="openThis(this.href,800,600);return false;"><%=sSeriesName%></a>
							<%End If%>
							-->
							<%If sIndivRelay = "Relay" Then%>
								<li class="list-inline-item list-inline-item-success">
									<a href="relay_rslts.asp?race_id=<%=lRaceID%>" onclick="openThis(this.href,800,600);return false;">Relay</a>
								</li>
							<%End If%>
						</ul>
				
						<form role="form" class="form-inline" name ="get_results" method="post" 
							action="cc_rslts.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>&amp;show_indiv=<%=sShowIndiv%>">
						<div class="form-group">
							<label style="font-weight: bold;">Select Results To View:</label>
							<select class="form-control" name="which_rslts" id="which_rslts" onchange="this.form.get_this.click();">
								<%Select Case sRsltsPage%>
									<%Case "overall_rslts.asp"%>
										<option value="overall_rslts.asp" selected>Overall</option>
										<option value="rslts_by_tm.asp">By Team</option>
									<%Case "rslts_by_tm.asp"%>
										<option value="overall_rslts.asp">Overall</option>
										<option value="rslts_by_tm.asp" selected>By Team</option>
									<%Case Else%>
										<option value="overall_rslts.asp">Overall</option>
										<option value="rslts_by_tm.asp">By Team</option>
								<%End Select%>
							</select>
							<input class="form-control" type="hidden" name="submit_this" id="submit_this" value="submit_this">
							<input class="form-control" type="submit" name="get_this" id="get_this" value="Go">
						</div>
						</form>
			
						<br>

						<%If sRsltsPage = "overall_rslts.asp" Then%>
							<%If sAdvancement = "y" Then%>
								<div class="bg-danger text-danger">
									<%If sExcludeTeam = "y" Then%>
										The top <%=iNumTeams%> teams and the top <%=iNumAdvance%> individuals not on advancing teams will advance to the <%=sAdvanceTo%>.
									<%Else%>
										The top <%=iNumTeams%> teams and the top <%=iNumAdvance%> individuals will advance to the <%=sAdvanceTo%>.
									<%End If%>
								</div>
							<%End If%>

							<%If sTeamScores = "y" Then%>
								<h4 class="h4">Team Results: Top <%=iNumScore%> Per Team</h4>
					
								<div class="bg-success">
									<%If sShowIndiv = "y" Then%>
										<a href="cc_rslts.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>&amp;show_indiv=n">Hide Individual Team Members</a>
									<%Else%>
										<a href="cc_rslts.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>&amp;show_indiv=y">Show Individual Team Members</a>
									<%End If%>
								</div>
								
								<%If sAdvancement = "y" Then%>
									<table class="table table-striped">
										<tr>
											<th>Pl</th>
											<th>Adv</th>
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

										<%
										m = 0
										ReDim AdvTeams(0)
										%>

										<%For i = 0 to UBound(TmRslts, 2)%>
											<%
												If CInt(i) + 1 <= CInt(iNumTeams) Then
													AdvTeams(m) = TmRslts(9, i)
													m = m + 1
													ReDim Preserve AdvTeams(m)
												End If
											%>
											<tr>
												<td><%=i + 1%></td>
												<%If i + 1 <= iNumTeams Then%>
													<td>X</td>
												<%Else%>
													<td>&nbsp;</td>
												<%End If%>
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
				
									<%If UBound(TmRslts4, 2) > 0 Then%>
										<h4 class="h4">Team Scoring: Top 4 Per Team</h4>

										<p>Since this race is not scored using the top 4 performers on each team, for 
										ranking purposes, you may view the team scores using the top 4 performers here.</p>

										<table class="table table-striped">
											<tr>
												<th>Pl</th>
												<th>Team</th>
												<th>Score</th>
												<th>S1</th>
												<th>S2</th>
												<th>S3</th>
												<th>S4</th>
											</tr>
											<%For i = 0 to UBound(TmRslts4, 2) - 1%>
												<tr>
													<td><%=i + 1%></td>
													<td><%=TmRslts4(0, i)%></td>
													<td><%=TmRslts4(1, i)%></td>
													<td><%=TmRslts4(2, i)%></td>
													<td><%=TmRslts4(3, i)%></td>
													<td><%=TmRslts4(4, i)%></td>
													<td><%=TmRslts4(5, i)%></td>
												</tr>
											<%Next%>
										</table>
									<%End If%>
								<%Else%>
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
				
									<%If UBound(TmRslts4, 2) > 0 Then%>
										<h4 class="h4">Team Scoring: Top 4 Per Team</h4>

										<p>Since this race is not scored using the top 4 performers on each team, for 
										ranking purposes, you may view the team scores using the top 4 performers here.</p>

										<table class="table table-striped">
											<tr>
												<th>Pl</th>
												<th>Team</th>
												<th>Score</th>
												<th>S1</th>
												<th>S2</th>
												<th>S3</th>
												<th>S4</th>
											</tr>
											<%For i = 0 to UBound(TmRslts4, 2) - 1%>
												<tr>
													<td><%=i + 1%></td>
													<td><%=TmRslts4(0, i)%></td>
													<td><%=TmRslts4(1, i)%></td>
													<td><%=TmRslts4(2, i)%></td>
													<td><%=TmRslts4(3, i)%></td>
													<td><%=TmRslts4(4, i)%></td>
													<td><%=TmRslts4(5, i)%></td>
												</tr>
											<%Next%>
										</table>
									<%End If%>
								<%End If%>
							<%End If%>

							<h4 class="h4">Individual Results</h4>

							<%If sIsStage = "y" AND sRaceStatus = "Destination" Then%>
								<table class="table table-striped">
									<tr>
										<th rowspan="2">Pl</th>
										<th rowspan="2">Bib-Name</th>
										<th rowspan="2">Team</th>
										<th rowspan="2">Gr</th>
										<th rowspan="2">M/F</th>
										<th rowspan="2">Time</th>
										<th style="text-align: center;" colspan="<%=iNumPrelims%>">Preliminary Times</th>
									</tr>
									<tr>
										<%For i = 0 To UBound(Prelims, 2) - 1%>
											<th><%=Prelims(1, i)%></th>
										<%Next%>
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
											<td><%=RsltsArr(8, i)%> - <%=RsltsArr(0, i)%>, <%=RsltsArr(1, i)%></td>
											<td><%=RsltsArr(2, i)%></td>
											<td><%=RsltsArr(3, i)%></td>
											<td><%=RsltsArr(4, i)%></td>
											<td><%=RsltsArr(5, i)%></td>
											<%For j = 0 To UBound(Prelims, 2) - 1%>
												<td><%=GetPrelimTime(RsltsArr(9, i), Prelims(2, j))%></td>
											<%Next%>
										</tr>
									<%Next%>
								</table>
							<%Else%>
								<%If Clng(lMeetID) = 308 Then%>
									<div class="bg-success">
										<a href="/results/cc_rslts/308/Lake Conf Sprint Team Results.pdf" onclick="openThis(this.href,1024,768);return false;">Team Results</a>
									</div>
								<%End If%>

								<%If sAdvancement = "y" Then%>
									<table class="table table-striped">
										<tr>
											<th>Pl</th>
											<th>Tm</th>
											<th>Adv</th>
											<th>Bib-Name</th>
											<th>Team</th>
											<th>Gr</th>
											<th>M/F</th>
											<th>Time</th>
											<th>Per Mi</th>
											<th>Per Km</th>
										</tr>
										<%iIndAdvTtl = 1%>
										<%k = 1%>
										<%For i = 0 to UBound(RsltsArr, 2)%>
											<%
											'determine advancement
											lMyRosterID = RsltsArr(9, i)
		
											'reset advancement code
											sAdvance = vbNullString
		
											'check for advancement via advancing team
											If UBound(AdvTeams) > 1 Then
												For m = 0 To UBound(AdvTeams) - 1
													Set rs = Server.CreateObject("ADODB.Recordset")
													sql = "SELECT RosterID FROM Roster WHERE TeamsID = " & AdvTeams(m) & " AND RosterID = " & lMyRosterID
													rs.Open sql, conn, 1, 2
													If rs.RecordCount > 0 Then sAdvance = "TM"
													rs.Close
													Set rs = Nothing
												Next
											End If

											'determine individual advancement
											If sExcludeTeam = "y" Then  'if team members are excluded from the individual total
												If sAdvance = vbNullString Then
													If iIndAdvTtl <= iNumAdvance Then
														sAdvance = "IND"
														iIndAdvTtl = iIndAdvTtl + 1
													End If
												End If
											Else
											End If
											%>
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
												<td><%=sAdvance%></td>
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
								<%Else%>
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
								<%End If%>
							<%End If%>
						<%Else%>
							<form role="form" class="form-inline" name="select_team" method="post" 
								action="cc_rslts.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>">
							<div class="form-group">
								<label>Select Team:</label>
								<select class="form-control" name="teams" id="teams" onchange="this.form.submit_team.click();">
									<option value="0">&nbsp;</option>
									<%For i = 0 to UBound(MeetTeams, 2)%>
										<%If CLng(lWhichTeam) = CLng(MeetTeams(0, i)) Then%>
											<option value="<%=MeetTeams(0, i)%>" selected><%=MeetTeams(1, i)%>&nbsp;(<%=MeetTeams(2, i)%>)</option>
										<%Else%>
											<option value="<%=MeetTeams(0, i)%>"><%=MeetTeams(1, i)%>&nbsp;(<%=MeetTeams(2, i)%>)</option>
										<%End If%>
									<%Next%>
								</select>&nbsp;&nbsp;
								<input class="form-control" type="hidden" name="get_team" id="get_team" value="get_team">
								<input class="form-control" type="submit" name="submit_team" id="submit_team" value= "View This Team">
							</div>
							</form>

							<%If Not CLng(lWhichTeam) = 0 Then%>
								<table class="table table-striped">
									<tr>
										<th>Rnr</th>
										<th>Name</th>
										<th>Gr</th>
										<th>M/F</th>
										<th>Time</th>
										<th>Per Mi</th>
										<th>Per Km</th>
									</tr>
									<%For i = 0 to UBound(RsltsArr, 2)%>
										<tr>
											<td><%=i + 1%></td>
											<td>
												<%=RsltsArr(0, i)%> - <%=RsltsArr(1, i)%>, <%=RsltsArr(2, i)%>
											</td>
											<td><%=RsltsArr(3, i)%></td>
											<td><%=RsltsArr(4, i)%></td>
											<td><%=RsltsArr(5, i)%></td>
											<td>
												<%If RsltsArr(5, i) <> "00:00" Then%>
													<%=PacePerMile(RsltsArr(5, i), iDist, sUnits)%>
												<%End If%>
											</td>
											<td>
												<%If RsltsArr(5, i) <> "00:00" Then%>
													<%=PacePerKm(RsltsArr(5, i), iDist, sUnits)%>
												<%End If%>
											</td>
										</tr>
									<%Next%>
								</table>
							<%End If%>
						<%End If%>
					<%End If%>
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
