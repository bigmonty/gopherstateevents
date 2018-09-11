<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs, rs2, sql2
Dim i, j, k
Dim lMeetID, lRaceID, lSeriesID, lRosterID, lWhichTeam, lPrelimID
Dim sRaceName, sSeriesName, sSeriesGender, sMeetName, sGradeYear, sOrderResultsBy, sScoreMethod, sRsltsPage, sTeamName, sUnits, sMeetSite, sWeather
Dim sRaceDist, sSport, sIndivRelay, sTeamScores, sResultsNotes, sErrMsg, sLogo, sShowIndiv
Dim iDist, iNumScore, iBibToFind
Dim BibRslts(5), RsltsArr, TmRslts, TmRslts4(), MeetTeams, SortArr(9), MeetArray, Races, PursuitTmRslts(), TempArr(9), RankArr(), RaceGallery(), OurRslts
Dim dMeetDate
Dim bRsltsOfficial, bTmRsltsReady

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

lWhichTeam = Request.QueryString("which_team")
If CStr(lWhichTeam) = vbNullString Then lWhichTeam = 0
If Not IsNumeric(lWhichTeam) Then Response.Redirect("http://www.google.com")
If CLng(lWhichTeam) < 0 Then Response.Redirect("http://www.google.com")

sRsltsPage = Request.QueryString("rslts_page")

sSport = Request.QueryString("sport")
If sSport = vbNullString Then sSport = "cc"

If sSport = "cc" or sSport = "Cross-Country" Then
    sSport = "Cross-Country"
Else
    sSport = "Nordic Ski"
End If

sShowIndiv = Request.QueryString("show_indiv")
If sShowIndiv = vbNullString Then sShowIndiv = "y"
bTmRsltsReady = True

iBibToFind = Request.QueryString("bib_to_find")
If CStr(iBibToFind) = vbNullString Then iBibToFind = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate <= '" & Now() & "' AND ShowOnline = 'y' AND Sport = '" & sSport 
sql = sql & "' ORDER BY MeetDate DESC"
Set rs = conn2.Execute(sql)
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
ElseIf Request.form.Item("submit_bib") = "submit_bib" Then
    iBibToFind = Request.Form.Item("bib_to_find")
End If

If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect "http://www.google.com"

If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect "http://www.google.com"

If CStr(lWhichTeam) = vbNullString Then lWhichTeam = 0
If Not IsNumeric(lWhichTeam) Then Response.Redirect "http://www.google.com"

If sRsltsPage = vbNullString Then sRsltsPage = "overall_rslts.asp"

'log this user if they are just entering the site
If Session("cc_results") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'cc_results')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

ReDim TmRslts4(5, 0)

If Not CLng(lMeetID) = 0 Then
    i = 0
    ReDim RaceGallery(0)
    sql = "SELECT EmbedLink FROM RaceGallery WHERE MeetsID = " & lMeetID
    Set rs = conn2.Execute(sql)
    Do While Not rs.EOF
        RaceGallery(i) = rs(0).Value
        i = i + 1
        ReDim Preserve RaceGallery(i)
        rs.MoveNext
    Loop
    Set rs = Nothing

    bRsltsOfficial = False
	sql = "SELECT MeetsID FROM OfficialRslts WHERE MeetsID = " & lMeetID
	Set rs = conn2.Execute(sql)
    If rs.BOF and rs.EOF Then
        bRsltsOfficial = False
    Else
        bRsltsOfficial = True
    End If
	Set rs = Nothing
	
	sql = "SELECT MeetName, MeetDate, MeetSite, Weather, Sport, OrderBy, Logo FROM Meets WHERE MeetsID = " & lMeetID
	Set rs = conn2.Execute(sql)
	sMeetName = rs(0).Value
	dMeetDate = rs(1).Value
	If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then sWeather = Replace(rs(3).Value, "''", "'")
    sSport = rs(4).Value
    sOrderResultsBy = rs(5).Value
    sLogo = rs(6).Value
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
	Set rs = conn2.Execute(sql)
	MeetTeams = rs.GetRows()
	Set rs = Nothing
	
	sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID
	Set rs = conn2.Execute(sql)
	Races = rs.GetRows()
	Set rs = Nothing
	
	If CLng(lRaceID) = 0 Then lRaceID = Races(0, 0)

	If Not CLng(lRaceID) = 0 Then
		sql = "SELECT RaceDesc, RaceDist, RaceUnits, ScoreMethod, NumScore, IndivRelay, TeamScores, ResultsNotes FROM Races WHERE RacesID = " & lRaceID
		Set rs = conn2.Execute(sql)
		sRaceName = Replace(rs(0).Value, "''", "'")
		iDist = rs(1).Value
		sUnits = rs(2).Value
		sScoreMethod = rs(3).Value
        iNumScore = rs(4).Value
        sIndivRelay = rs(5).Value
        sTeamScores = rs(6).Value
        If Not rs(7).Value & "" = "" Then sResultsNotes = Replace(rs(7).Value, "''", "'")
		Set rs = Nothing

		sRaceDist = iDist & " " & sUnits

		'see if this race is in a series
		sql = "SELECT s.SeriesID, s.SeriesName FROM Series s INNER JOIN SeriesMeets sm ON s.SeriesID = sm.SeriesID "
		sql = sql & "WHERE sm.RacesID = " & lRaceID
		Set rs = conn2.Execute(sql)
        If rs.BOF and rs.EOF Then
            lSeriesID = 0
        Else
			lSeriesID = rs(0).Value
			sSeriesName = Replace(rs(1).Value, "''", "'")
        End If
		Set rs = Nothing
	
		If sRsltsPage = "overall_rslts.asp" Then
            If sScoreMethod = "Pursuit" Then
                Dim sPursuitName, sPursuitDesc, sPrelimDesc
                Dim sPursuitDist, sPrelimDist, dPrelimDate
                Dim sPrelimMeet

	            'get pursuit data
	            sql = "SELECT RaceDist, RaceUnits, RaceDesc, RaceName FROM Races WHERE RacesID = " & lRaceID
		        Set rs = conn2.Execute(sql)
                If rs.BOF and rs.EOF Then
                    '--
                Else
	                sPursuitDist = rs(0).Value & " " & rs(1).Value
	                sPursuitDesc = rs(2).Value
	                sPursuitName = rs(3).Value
                End If
	            Set rs = Nothing
	      
	            'get prelim race id
	            sql = "SELECT PrelimRace FROM Pursuit WHERE RacesID = " & lRaceID
		        Set rs = conn2.Execute(sql)
                If rs.BOF and rs.EOF Then
                    lPrelimID = 0
                Else
	                lPrelimID = rs(0).Value
                End If
	            Set rs = Nothing
	      
	            'get prelim race info
	            sql = "SELECT r.RaceDist, r.RaceUnits, m.MeetDate, r.RaceDesc, m.MeetName FROM Races r INNER JOIN Meets m "
	            sql = sql & "ON r.MeetsID = m.MeetsID WHERE RacesID = " & lPrelimID
		        Set rs = conn2.Execute(sql)
                If rs.BOF and rs.EOF Then
                    '--
                Else
	                sPrelimDist = rs(0).Value & " " & rs(1).Value
	                dPrelimDate = rs(2).Value
	                sPrelimDesc = rs(3).Value
	                sPrelimMeet = rs(4).Value
                End If
	            Set rs = Nothing

 	            i = 0
	            ReDim PursuitTmRslts(9, 0)
	            sql = "SELECT t.TeamsID, t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
	            sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID
	            Set rs = conn2.Execute(sql)
	            Do While Not rs.EOF
	                For j = 0 To 9
	                    PursuitTmRslts(j, i) = Trim(rs(j).Value)
	                Next
	                i = i + 1
	                ReDim Preserve PursuitTmRslts(9, i)
	                rs.MoveNext
	            Loop
	            Set rs = Nothing
	
	            'sort the arrays
	            For i = 0 To UBound(PursuitTmRslts, 2) - 2
	                For j = i + 1 To UBound(PursuitTmRslts, 2) - 1
	                    If ConvertToSeconds(PursuitTmRslts(2, i)) < ConvertToSeconds(PursuitTmRslts(2, j)) Then
	                        For k = 0 To 9
	                            TempArr(k) = PursuitTmRslts(k, i)
	                            PursuitTmRslts(k, i) = PursuitTmRslts(k, j)
	                            PursuitTmRslts(k, j) = TempArr(k)
	                        Next
	                    End If
	                Next
	            Next
	        
	            'get a finishers array for this race
	            i = 0
	            ReDim PursuitIndRslts(8, 0)
	            sql = "SELECT r.RosterID, r.FirstName, r.LastName, t.TeamName, ir.RaceTime, ir.Bib, g.Grade" & sGradeYear & " FROM IndRslts ir "
                sql = sql & "INNER JOIN Roster r ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID "
	            sql = sql & " INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
	            sql = sql & "WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' "
	            sql = sql & "ORDER BY ir.Place"
		        Set rs = conn2.Execute(sql)
	            Do While Not rs.EOF
	                PursuitIndRslts(0, i) = rs(0).Value                 'name
	                PursuitIndRslts(1, i) = rs(5).Value & "-" & Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	                PursuitIndRslts(2, i) = rs(6).Value                 'grade
	                PursuitIndRslts(3, i) = rs(3).Value                 'team
	                PursuitIndRslts(8, i) = rs(4).Value                 'pursuit time
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
	            sql = "SELECT RosterID, RaceTime FROM IndRslts WHERE RacesID = " & lPrelimID & " AND RaceTime > '00:00' AND Place > 0 ORDER BY RaceTime"
	            Set rs = conn2.Execute(sql)
	            Do While Not rs.EOF
	                For j = 0 To UBound(PursuitIndRslts, 2) - 1
	                    If CLng(PursuitIndRslts(0, j)) = CLng(rs(0).Value) Then
	                        PursuitIndRslts(4, j) = ConvertToMinutes(ConvertToSeconds(rs(1).Value) + ConvertToSeconds(PursuitIndRslts(8, j)))
	                        PursuitIndRslts(5, j) = i
	                        PursuitIndRslts(6, j) = rs(1).Value
	                        i = i + 1
	                    End If
	                Next
	                rs.MoveNext
	            Loop
	            Set rs = Nothing
            Else
                If sOrderResultsBy = "Time" Then
			        sql = "SELECT r.LastName, r.FirstName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ir.Excludes, ir.TeamPlace, ir.Bib "
                    sql = sql & "FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
                    sql = sql & "INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' "
			        sql = sql & "ORDER BY ir.FnlScnds"
                Else
			        sql = "SELECT r.LastName, r.FirstName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ir.Excludes, ir.TeamPlace, ir.Bib "
                    sql = sql & "FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
                    sql = sql & "INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' "
			        sql = sql & "ORDER BY ir.Place"
                End If
				Set rs = conn2.Execute(sql)
                If rs.BOF and rs.EOF Then
                    ReDim RsltsArr(8, 0)
                Else
                    RsltsArr = rs.GetRows()
                End If
			    Set rs = Nothing

                For i = 0 To UBound(RsltsArr, 2)
                    RsltsArr(0, i) = Replace(RsltsArr(0, i), "''", "'")
                    RsltsArr(1, i) = Replace(RsltsArr(1, i), "''", "'")
                    RsltsArr(2, i) = Replace(RsltsArr(2, i), "''", "'")
'                    If Not RsltsArr(3, i) & "" = "" Then RsltsArr(3, i) = GetGrade(RsltsArr(3,i))
                    If RsltsArr(6, i) = "y" Then 
                        RsltsArr(7,i) = "---"
                    Else
                        If CInt(RsltsArr(7, i)) = 0 Then RsltsArr(7,i) = "---"
                    End If
                Next
	
                If sTeamScores = "y" Then
                    If sSport = "Cross-Country" Then
			            sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7, t.TeamsID FROM Teams t INNER JOIN TmRslts tr "
			            sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> '' ORDER BY Score DESC"
                    Else
			            sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7, t.TeamsID FROM Teams t INNER JOIN TmRslts tr "
			            sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> '' ORDER BY Score"
                    End If
				    Set rs = conn2.Execute(sql)
                    If rs.BOF and rs.EOF Then
                        ReDim TmRslts(9, 0)
                        bTmRsltsReady = False
                    Else
                        TmRslts = rs.GetRows()
                    End If
			        Set rs = Nothing

                    If sSport = "Cross-Country" Then
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
			                Set rs = conn2.Execute(sql)
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
				Set rs = conn2.Execute(sql)
				sTeamName = rs(0).Value & " (" & rs(1).Value & ")"
				Set rs = Nothing
				
				Set rs = Server.CreateObject("ADODB.Recordset")
                If sOrderResultsBy = "Time" Then
				    sql = "SELECT ir.Bib, r.LastName, r.FirstName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime FROM IndRslts ir "
				    sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
				    sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lRaceID & " AND t.TeamsID = " 
				    sql = sql & lWhichTeam & " AND ir.Place > 0 ORDER BY ir.FnlScnds"
                Else
				    sql = "SELECT ir.Bib, r.LastName, r.FirstName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime FROM IndRslts ir "
				    sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
				    sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lRaceID & " AND t.TeamsID = " 
				    sql = sql & lWhichTeam & " AND ir.Place > 0 ORDER BY ir.Place"
                End If
				Set rs = conn2.Execute(sql)
                If rs.BOF and rs.EOF Then
                    ReDim RsltsArr(5, 0)
                Else
                    RsltsArr = rs.GetRows()
                End If
			    Set rs = Nothing

                For i = 0 To UBound(RsltsArr, 2)
                    RsltsArr(1, i) = Replace(RsltsArr(1, i), "''", "'")
                    RsltsArr(2, i) = Replace(RsltsArr(2, i), "''", "'")
'                    If Not RsltsArr(3, i) & "" = "" Then RsltsArr(3, i) = GetGrade(RsltsArr(3, i))
                Next
			End If
		End If
	End If
End If

If Not CInt(iBibToFind) = 0 Then
    sql = "SELECT r.RosterID, r.FirstName, r.LastName, t.TeamName, r.Gender, ir.RaceTime, g.Grade" & sGradeYear & " FROM IndRslts ir "
	sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
	sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lRaceID & " AND ir.Bib = " & iBibToFind
    Set rs = conn2.Execute(sql)
    If rs.BOF and rs.EOF Then
        sErrMsg = "I'm sorry.  That bib number was not found in thre results for this race.  Please check another race in this meet."
    Else
        BibRslts(0) = GetPlace(rs(0).Value)
        BibRslts(1) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        BibRslts(2) = rs(3).Value
        BibRslts(3) = rs(4).Value
        BibRslts(4) = rs(6).Value
        BibRslts(5) = rs(5).Value
    End If
    Set rs = Nothing
End If

Function GetPlace(lRosterID)
	GetPlace = 0
	sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lRaceID & " AND Place > 0 ORDER BY FnlScnds"
	Set rs2 = conn2.Execute(sql2)
	Do While Not rs2.EOF
		GetPlace = GetPlace + 1
		If CLng(rs2(0).Value) = CLng(lRosterID) Then Exit Do
		rs2.MoveNext
	Loop
	Set rs2 = Nothing
End Function

Private Function ConvertToSeconds(sTime)
    Dim sSubStr(3), Count, j
    Dim sglSeconds(3), k

    sTime = Replace(sTime, "-", "")

    'find out how many substrings are needed
    If sTime & "" = "" Then
		ConvertToSeconds = 0
    Else
		For k = 1 To Len(sTime)
		    If Mid(sTime, k, 1) = ":" Then Count = Count + 1
		Next
    
		'break the time into substrings
		For k = 1 To Len(sTime)
		    If Mid(sTime, k, 1) = ":" Then
		        j = j + 1
		    Else
		        sSubStr(j) = sSubStr(j) & Mid(sTime, k, 1)
		    End If
		Next
    
		'do the conversion
		For k = 0 To Count
		    j = Count - k
		    If sSubStr(k) = vbNullString Then
		        sglSeconds(k) = 0
		    Else
		        sglSeconds(k) = CSng(sSubStr(k)) * (60 ^ j)
		    End If
		    ConvertToSeconds = ConvertToSeconds + sglSeconds(k)
		Next
	End If
End Function

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If sglScnds <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(sglScnds \ 60)
    sSecondPart = CStr(((sglScnds / 60) - (sglScnds \ 60)) * 60)
    
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

Private Function MeetName()
    sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lMeetID
    Set rs = conn2.Execute(sql)
    MeetName = Replace(rs(0).Value, "''", "'")
    Set rs = Nothing
End Function

%>
<!--#include file = "../../includes/clean_input.asp" -->
<%

Private Sub GetIndiv(lThisTeam)
	sql = "SELECT r.RosterID, ir.Bib, r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.RaceTime FROM IndRslts ir INNER JOIN Roster r"
	sql = sql & " ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID "
	sql = sql & "WHERE ir.RacesID = " & lRaceID & " AND r.TeamsID = " & lThisTeam & " AND ir.Place > 0 AND ir.FnlScnds > 0 AND ir.Excludes = 'n' "
	sql = sql & "ORDER BY ir.FnlScnds"
	Set rs = conn2.Execute(sql)
    If rs.BOF and rs.EOF Then
	    ReDim OurRslts(5, 0)
    Else
	    OurRslts = rs.GetRows()
    End If
    Set rs = Nothing
End Sub

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events Cross-Country/Nordic Ski Results</title>
<meta name="description" content="Cross-Country & Nordic Ski Results by Gopher State Events, a conventional timing service offererd by H51 Software, LLC in Minnetonka, MN.">
<!--#include file = "../../includes/js.asp" --> 
<script>
function chkFlds2() {
 	if (document.find_bib.bib_to_find.value == '')
		{
  		alert('You must submit a bib number to look for.');
  		return false
  		}
 	else
		if (isNaN(document.find_bib.bib_to_find.value))
    		{
			alert('The bib number can not contain non-numeric values');
			return false
			}
	else
   		return true
}
</script>

<style type="text/css">
    td,th{
        padding-left: 5px;
    }
</style>
</head>
<body>
<div class="container">
    <div class="row">
        <img src="/graphics/html_header.png" class="img-responsive" alt="Individual Results">
	    <h4 class="h4">GSE <%=sSport%> Results</h4>
        <a href="http://www.gopherstateevents.com" style="font-weight: bold;">Return To Main Site</a>
    </div>

    <div class="row">
        <div class="col-xs-9">
	        <form role="form" name="get_races" method="post" action="cc_rslts_mobile.asp?sport=<%=sSport%>&amp;show_indiv=<%=sShowIndiv%>">
	        <div class="form-group">
                <h5 class="h5">Select Meet:</h5>
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
	            <input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
	            <input class="form-control" type="submit" name="get_meet" id="get_meet" value="View This Meet">
            </div>
	        </form>

	        <%If Not CLng(lMeetID) = 0 Then%>
		        <h4 class="h4">Results for <%=sMeetName%> on <%=dMeetDate%></h4>

		        <%If Not sWeather = vbNullString Then%>
			        <p>Weather:</span>&nbsp;<%=sWeather%></p>
		        <%End If%>
				
                <%If CDate(Date) < CDate(dMeetDate) + 7 Then%>
			        <%If bRsltsOfficial = False Then%>
				        <p>
                            <span style="color: red;font-weight: bold;">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE</span><br>
                            Please report any issues to bob.schneider@gopherstateevents.com.
                        </p>
			        <%Else%>
				        <p>These results are now official.  If you notice any errors please contact us 
				        via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by telephone (612-720-8427).</p>
			        <%End If%>
                <%End If%>

                <%If CLng(lMeetID) = 249 Then%>
                    <a href="/results/cc_rslts/249/B-V_Rslts.txt">Boys Varsity</a>
                    &nbsp;|&nbsp;
                    <a href="/results/cc_rslts/249/G-V_Rslts.txt">Girls Varsity</a>
                    &nbsp;|&nbsp;
                    <a href="/results/cc_rslts/249/B-JV1_Rslts.txt">Boys JV1</a>
                    &nbsp;|&nbsp;
                    <a href="/results/cc_rslts/249/G-JV1_Rslts.txt">Girls JV1</a>
                    &nbsp;|&nbsp;
                    <a href="/results/cc_rslts/249/B-V_Rslts-Relay.txt">Boys Varsity Splits</a>
                    &nbsp;|&nbsp;
                    <a href="/results/cc_rslts/249/G-V_Rslts-Relay.txt">Girls Varsity Splits</a>
                    &nbsp;|&nbsp;
                    <a href="/results/cc_rslts/249/B-JV1_Rslts-Relay.txt">Boys JV1 Splits</a>
                    &nbsp;|&nbsp;
                    <a href="/results/cc_rslts/249/G-JV1_Rslts-Relay.txt">Girls JV1 Splits</a>
                <%Else%>
			        <div style="margin-bottom:10px;">	
				        <form name="get_races" method="post" action="cc_rslts_mobile.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;rslts_page=<%=sRsltsPage%>&amp;show_indiv=<%=sShowIndiv%>"
                                style="margin: 10px 0 10px 0;">
				        <span style="font-weight:bold;">Select Race:</span>
				        <select name="races" id="races" onchange="this.form.get_race.click();">
					        <option value="0">&nbsp;</option>
					        <%For i = 0 to UBound(Races, 2)%>
						        <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
							        <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
						        <%Else%>
							        <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
						        <%End If%>
					        <%Next%>
				        </select>
				        <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
				        <input type="submit" name="get_race" id="get_race" value="Get Results">
				        </form>
			        </div>
				    
                    <p>
                        Distance:&nbsp;<%=sRaceDist%>
                        <br>
                        Site/Location:&nbsp;<%=sMeetSite%>
                        <%If Not sResultsNotes & "" = "" Then%>
                            <br>
                            <span style="color: red;">Results Notes:&nbsp;<%=sResultsNotes%></span>
                        <%End If%>
                    </p>

			        <%If sScoreMethod="Pursuit" Then%>
                        <h4 class="h4">Prelim Meet:&nbsp;<%=sPrelimMeet%></h4>
					            
                        <h4 class="h4">Team Results: Top <%=iNumScore%> Per Team</h4>						
					
                        <table  class="table-striped">
					        <tr>
						        <th style="text-align:right">Pl</th>
						        <th>Team</th>
						        <th style="text-align:center">Score</th>
						        <th style="text-align:center;">R1</th>
						        <th style="text-align:center;">R2</th>
						        <th style="text-align:center;">R3</th>
						        <th style="text-align:center;">R4</th>
						        <th style="text-align:center;">R5</th>
						        <th style="text-align:center;">R6</th>
						        <th style="text-align:center;">R7</th>
					        </tr>
					        <%For i = 0 to UBound(PursuitTmRslts, 2)%>
							    <tr>
								    <td style="text-align:right"><%=i + 1%></td>
								    <td style="white-space:nowrap;"><%=PursuitTmRslts(1, i)%></td>
								    <td style="text-align:center"><%=PursuitTmRslts(2, i)%></td>
								    <td style="text-align:center;"><%=PursuitTmRslts(3, i)%></td>
								    <td style="text-align:center;"><%=PursuitTmRslts(4, i)%></td>
								    <td style="text-align:center;"><%=PursuitTmRslts(5, i)%></td>
								    <td style="text-align:center;"><%=PursuitTmRslts(6, i)%></td>
								    <td style="text-align:center;"><%=PursuitTmRslts(7, i)%></td>
								    <td style="text-align:center;"><%=PursuitTmRslts(8, i)%></td>
								    <td style="text-align:center;"><%=PursuitTmRslts(9, i)%></td>
							    </tr>
					        <%Next%>
				        </table>

				        <h4 class="h4">Individual Results</h4>

                        <table class="table-striped">
                            <tr>
                                <th rowspan="2" valign="bottom">Place</th>
                                <th rowspan="2" valign="bottom">Bib-Name</th>
                                <th rowspan="2" valign="bottom">Gr</th>
                                <th rowspan="2" valign="bottom">Team</th>
                                <th style="text-align:center;" rowspan="2" valign="bottom">Comb Time</th>
                                <th style="text-align: center;" colspan="2">Round 1</th>
                                <th style="text-align: center;" colspan="2">Round 2</th>
                            </tr>
                            <tr>
                                <th style="text-align:center;">Rank</th>
                                <th style="text-align:center;">Time</th>
                                <th style="text-align:center;">Rank</th>
                                <th style="text-align:center;">Time</th>
                            </tr>
                            <%For i = 0 To UBound(PursuitIndRslts, 2) - 1%>
                                <tr>
                                    <td><%=i + 1%>)</td>
                                    <td style="white-space: nowrap;"><%=PursuitIndRslts(1, i)%></td>
                                    <td><%=PursuitIndRslts(2, i)%></td>
                                    <td><%=PursuitIndRslts(3, i)%></td>
                                    <td style="text-align:center;"><%=PursuitIndRslts(4, i)%></td>
                                    <td style="text-align:center;"><%=PursuitIndRslts(5, i)%></td>
                                    <td style="text-align:center;"><%=PursuitIndRslts(6, i)%></td>
                                    <td style="text-align:center;"><%=PursuitIndRslts(7, i)%></td>
                                    <td style="text-align:center;"><%=PursuitIndRslts(8, i)%></td>
                                </tr>
                            <%Next%>
                        </table>
			        <%Else%>
			            <div>
					        <a href="javascript:pop('rslts_by_team.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',625,650)"
                                style="color: red;">Results By Team</a>
					        &nbsp;|&nbsp;
				            <a href="javascript:pop('cc_rslts_grade.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>',625,650)">Finish Order By Grade</a>
				            &nbsp;|&nbsp;
				            <a href="comp_rslts.asp?meet_id=<%=lMeetID%>" onclick="openThis(this.href,1024,768);return false;">Comprehensive</a>
				            &nbsp;|&nbsp;
				            <a href="dwnld_overall.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>" 
						            onclick="openThis(this.href,800,600);return false;">Dwnld Overall</a>
				            <%If sSport = "Cross-Country" Then%>
                                &nbsp;|&nbsp;
					            <a href="dual_rslts.asp?meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>" 
						            onclick="openThis(this.href,800,600);return false;">Dual Meet</a>
                            <%End If%>
				            <%If sIndivRelay = "Relay" Then%>
                                &nbsp;|&nbsp;
					            <a href="relay_rslts.asp?race_id=<%=lRaceID%>" onclick="openThis(this.href,800,600);return false;">Relay Results</a>
				            <%End If%>
			            </div>

                        <%If sTeamScores = "y" Then%>
		                    <h4 class="h4">Team Results: Top <%=iNumScore%> Per Team</h4>
						
                            <div style="background-color: #ececec;padding-left: 10px;">
                                <%If sShowIndiv = "y" Then%>
                                    <a href="cc_rslts_mobile.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>&amp;show_indiv=n&amp;bib_to_find=<%=iBibToFind%>">Hide Individual Team Members</a>
                                <%Else%>
                                    <a href="cc_rslts_mobile.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;rslts_page=<%=sRsltsPage%>&amp;bib_to_find=<%=iBibToFind%>">Show Individual Team Members</a>
                                <%End If%>
                            </div>
						
		                    <table class="table-striped">
			                    <tr>
				                    <th style="text-align:right">Pl</th>
				                    <th>Team</th>
				                    <th style="text-align:center">Score</th>
				                    <th style="text-align:center;">R1</th>
				                    <th style="text-align:center;">R2</th>
				                    <th style="text-align:center;">R3</th>
				                    <th style="text-align:center;">R4</th>
				                    <th style="text-align:center;">R5</th>
				                    <th style="text-align:center;">R6</th>
				                    <th style="text-align:center;">R7</th>
			                    </tr>
			                    <%For i = 0 to UBound(TmRslts, 2)%>
					                <tr>
						                <td style="text-align:right"><%=i + 1%></td>
						                <td style="white-space: nowrap;"><%=TmRslts(0, i)%></td>
						                <td style="text-align:center"><%=TmRslts(1, i)%></td>
						                <td style="text-align:center;"><%=TmRslts(2, i)%></td>
						                <td style="text-align:center;"><%=TmRslts(3, i)%></td>
						                <td style="text-align:center;"><%=TmRslts(4, i)%></td>
						                <td style="text-align:center;"><%=TmRslts(5, i)%></td>
						                <td style="text-align:center;"><%=TmRslts(6, i)%></td>
						                <td style="text-align:center;"><%=TmRslts(7, i)%></td>
						                <td style="text-align:center;"><%=TmRslts(8, i)%></td>
					                </tr>

                                    <%If sShowIndiv = "y" Then%>
                                        <tr>
                                            <td colspan="10" valign="top">
                                                <%If bTmRsltsReady = True Then%>
                                                    <%Call GetIndiv(TmRslts(9, i))%>
            
                                                    <%If UBound(OurRslts, 2) > 0 Then%>
	                                                    <table class="table-striped">
			                                                <%For j = 0 to UBound(OurRslts, 2)%>
					                                            <tr>
						                                            <td><%=TmRslts(j + 2, i)%></td>
						                                            <td style="white-space: nowrap;"><%=OurRslts(1, j)%>-<%=OurRslts(3, j)%>, <%=OurRslts(2, j)%></td>
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

			                    <table class="table-striped">
				                    <tr>
					                    <th style="text-align:right">Pl</th>
					                    <th>Team</th>
					                    <th style="text-align:center">Score</th>
					                    <th style="text-align:center;">S1</th>
					                    <th style="text-align:center;">S2</th>
					                    <th style="text-align:center;">S3</th>
					                    <th style="text-align:center;">S4</th>
				                    </tr>
				                    <%For i = 0 to UBound(TmRslts4, 2) - 1%>
						                <tr>
							                <td style="text-align:right"><%=i + 1%></td>
							                <td style="white-space: nowrap;"><%=TmRslts4(0, i)%></td>
							                <td style="text-align:center"><%=TmRslts4(1, i)%></td>
							                <td style="text-align:center;"><%=TmRslts4(2, i)%></td>
							                <td style="text-align:center;"><%=TmRslts4(3, i)%></td>
							                <td style="text-align:center;"><%=TmRslts4(4, i)%></td>
							                <td style="text-align:center;"><%=TmRslts4(5, i)%></td>
						                </tr>
				                    <%Next%>
			                    </table>
                            <%End If%>
                        <%End If%>

	                    <h4 class="h4">Individual Results</h4>

                        <table class="table-condensed">
                            <tr>
                                <td style="text-align:center;">
                                    <form name="find_bib" method="post" action="cc_rslts_mobile.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>&amp;race_id=<%=lRaceID%>&amp;show_indiv=<%=sShowIndiv%>" 
                                        onsubmit="return chkFlds2();">
                                    <span style="font-weight:bold;">Bib To Find:</span>
                                    <br>
                                    <input type="text" name="bib_to_find" id="bib_to_find" size="3" value="<%=iBibToFind%>">
                                    <input type="hidden" name="submit_bib" id="submit_bib" value="submit_bib">
                                    <input type="submit" name="submit_lookup" id="submit_lookup" value="Find Bib">
                                    </form>
                                </td>
                                <td>
                                    <%If Not CInt(iBibToFind) = 0 Then%>
                                        <%If sErrMsg = vbNullString Then%>
                                            <table>
                                                <tr><th>Pl</th><th>Name</th><th>School</th><th>MF</th><th>Gr</th><th>Time</th></tr>
                                                <tr>
                                                    <%For i = 0 To 5%>
                                                        <td><%=BibRslts(i)%></td>
                                                    <%Next%>
                                                </tr>
                                            </table>
                                        <%Else%>
                                            <p><%=sErrMsg%></p>
                                        <%End If%>
                                    <%End If%>
                                </td>
                            </tr>
                        </table>

	                    <table class="table-striped">
		                    <tr>
			                    <th style="text-align:right">Pl</th>
			                    <th style="text-align:right">Tm</th>
			                    <th>Bib-Name</th>
			                    <th>Team</th>
			                    <th style="text-align:center">Gr</th>
			                    <th style="text-align:center">M/F</th>
			                    <th style="text-align:center">Time</th>
		                    </tr>
		                    <%k = 1%>
		                    <%For i = 0 to UBound(RsltsArr, 2)%>
				                <tr>
					                <td>
						                <%If RsltsArr(8, i) = "y" Then%>
							                -
						                <%Else%>
							                <%=k%>
							                <%k = k + 1%>
						                <%End If%>
					                </td>
					                <td><%=RsltsArr(7, i)%></td>
					                <td style="white-space: nowrap;"><%=RsltsArr(8, i)%> - <%=RsltsArr(0, i)%>, <%=RsltsArr(1, i)%></td>
					                <td style="white-space: nowrap;"><%=RsltsArr(2, i)%></td>
					                <td><%=RsltsArr(3, i)%></td>
					                <td><%=RsltsArr(4, i)%></td>
					                <td><%=RsltsArr(5, i)%></td>
				                </tr>
		                    <%Next%>
	                    </table>
                    <%End If%>
                <%End If%>
        </div>
        <div class="col-xs-3">
            <%If CDate(dMeetDate) > CDate("9/1/2013") Then%>
                <%If Not sLogo & "" = "" Then%>
                    <img src="/events/logos/<%=sLogo%>" class="img-responsive" alt="Logo">
                    <br><br>
                <%End If%>

                <%If UBound(RaceGallery) = 0 Then%>
                    <%If Date < CDate(dMeetDate) + 10 Then%>
                        <img src="/graphics/no_pix.png" alt="Pix Not Available Yet" class="img-responsive">
                    <%End If%>
                <%Else%>
                    <%For i = 0 To UBound(RaceGallery) - 1%>
                        <%=RaceGallery(i)%>
                        <br>
                    <%Next%>
                <%End If%>

    <!--
                    <h4  style="background-color: #80bfff;padding-left: 5px;margin:4px;color: #000;">Media Order Form</h4>
                    <p style="margin: 4px 4px 0 4px;padding: 2px 2px 0 2px;">
                        <a href="javascript:pop('/graphics/demo_pic.png',477,640)"><img src="/graphics/demo_pic.png" 
                            style="width: 75px;float: right;margin-bottom:10px;" alt="Demo"></a>
                            Order 5-sec video clip and a finish line pic for $10.
                        <br><br>
                        <a href="javascript:pop('http://youtu.be/s7hNfF26vBw',1024,768)" style="font-weight: bold;">View Sample Video</a>
                    </p>
                    <div style="clear:both;"></div>

                    <form name="order_video" method="post" action="cc_rslts.asp?meet_id=<%=lMeetID%>" onsubmit="return chkFlds();">
                    <table style="background-color:#80bfff;margin:0 4px 4px 4px;">
                        <tr><th>Bib No:</th><td><input type="text" name="bib_num" id="bib_num" size="3"></td></tr>
                        <tr><th>Email:</th><td><input type="text" name="email" id="email" size="25"></td></tr>
                        <tr>
                            <td colspan="2" style="text-align: center;">
                                <input type="hidden" name="submit_order" id="submit_order" value="submit_order">
                                <input type="submit" name="submit4x" id="submit4x" value="Order Video">
                            </td>
                        </tr>
                    </table>
                    </form>
                    <p style="margin: 0 4px 0 4px;padding: 2px;font-size: 0.8em;">Your order will incude a finish line picture and a video.  We will 
                        verify the order and send an online payment linkr.  Once payment is received we will email your media.</p>
    -->
                <%Else%>
                    <!--#include file = "../../includes/vira_sponsors.asp" -->
                <%End If%>
            <%End If%>
        </div>
    </div>
</div>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
