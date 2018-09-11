<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k, m, n
Dim lThisMeet, lThisRace
Dim iOurFinishers, iRacePlace, iNumScore, iScoringPlaces, iMultiplier, iNumInd
Dim sMeetName, sRaceName, sThisTime, sTempTime, sTeamScores, sUpdateThese, sOrderResultsBy
Dim dMeetDate
Dim RacesArr(), TeamsArr(), TempArr(8), RaceTimes(), TimesArr(), TempTimes(), PlaceArr(), ComplTeams(), DuplTimes()
Dim bIsDupl

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")

lThisRace = Request.QueryString("this_race")
If CStr(lThisRace) = vbNullString Then lThisRace = 0

sUpdateThese = Request.QueryString("update_these")
If sUpdateThese = vbNullString Then sUpdateThese = "n"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

If Request.Form.Item("submit_race") = "submit_race" Then
    lThisRace = Request.Form.Item("races")
    If CStr(lThisRace) = vbNullString Then lThisRace = 0
End If

i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisMeet
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

If CLng(lThisRace) = 0 Then lThisRace = Races(0, 0)

sql = "SELECT RaceName FROM Races WHERE RacesID = " & lThisRace
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
Set rs = Nothing

If sUpdateThese = "y" Then Call UpdateTeamScores()

i = 0
ReDim TmRslts(8, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lThisRace & " AND tr.Score <> ''"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	TmRslts(0, i) = rs(0).Value
	TmRslts(1, i) = rs(1).Value
	TmRslts(2, i) = Trim(rs(2).Value)
	TmRslts(3, i) = Trim(rs(3).Value)
	TmRslts(4, i) = Trim(rs(4).Value)
	TmRslts(5, i) = Trim(rs(5).Value)
	TmRslts(6, i) = Trim(rs(6).Value)
	TmRslts(7, i) = Trim(rs(7).Value)
	TmRslts(8, i) = Trim(rs(8).Value)
	i = i + 1
	ReDim Preserve TmRslts(8, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub UpdateTeamScores()
    'get order by
    sql = "SELECT OrderBy FROM Meets WHERE MeetsID = " & lThisMeet
    Set rs = conn.Execute(sql)
    If rs(0).Value = "Place" Then
        sOrderResultsBy = "Place"
    Else
        sOrderResultsBy = "Time"
    End If
    Set rs = Nothing
    
    'now get num scorers for this race
    sql = "SELECT TeamScores FROM Races WHERE RacesID = " & lThisRace
    Set rs = conn.Execute(sql)
    sTeamScores = rs(0).Value
    Set rs = Nothing

    ReDim DuplTimes(3, 0)
               
    If sTeamScores = "y" Then
        'delete results for this race from the table
        sql = "DELETE FROM TmRslts WHERE RacesID = " & lThisRace
        Set rs = conn.Execute(sql)
        Set rs = Nothing
        
        'now get num scorers for this race
        sql = "SELECT NumScore FROM Races WHERE RacesID = " & lThisRace
        Set rs = conn.Execute(sql)
        iNumScore = rs(0).Value
        Set rs = Nothing
                    
        'get the order of finish if by time
        If sOrderResultsBy = "Time" Then
            Dim IndRsltsArr

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT r.TeamsID, ir.Place, ir.RaceTime FROM IndRslts ir INNER JOIN Roster r  ON r.RosterID = ir.RosterID WHERE  ir.RacesID = " 
            sql = sql & lThisRace & " AND ir.Place > 0 AND ir.ElpsdTime > '00:00' AND FnlScnds > 0 AND ir.Excludes = 'n' ORDER BY ir.FnlScnds, ir.Place"
            rs.Open sql, conn, 1, 2
            IndRsltsArr = rs.GetRows()
            rs.Close
            Set rs = Nothing
            
            'now reset the place
            For i = 0 To UBound(IndRsltsArr, 2)
                IndRsltsArr(1, i) = i + 1
            Next
        End If
    
        'get teams array for this meet
        j = 0
        ReDim TeamsArr(8, 0)
        sql = "SELECT t.TeamsID FROM Teams t INNER JOIN MeetTeams mt ON t.TeamsID = mt.TeamsID WHERE mt.MeetsID = " & lThisMeet
        Set rs = conn.Execute(sql)
        Do While Not rs.EOF
            TeamsArr(0, j) = rs(0).Value
            j = j + 1
            ReDim Preserve TeamsArr(8, j)
            rs.MoveNext
        Loop
        Set rs = Nothing
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT PlaceOnTeam, TeamPlace FROM IndRslts WHERE RacesID = " & lThisRace
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            rs(0).Value = 0    'reset this field for all teams in this race
            rs(1).Value = 0    'reset this field for all teams in this race
            rs.Update
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
                    
        For j = 0 To UBound(TeamsArr, 2) - 1
            Select Case ScoreMeth()
                Case "Place"
                    'assign a place on the team for each finisher
                    m = 0
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    sql = "SELECT ir.PlaceOnTeam, ir.TeamPlace FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
                    sql = sql & "WHERE ir.RacesID = " & lThisRace & " AND r.TeamsID = " & TeamsArr(0, j)
                    sql = sql & " AND ir.RaceTime > '00:00' AND ir.Excludes = 'n' AND ir.Place > 0 AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
                    rs.Open sql, conn, 1, 2
                    Do While Not rs.EOF
                        m = m + 1
                        rs(0).Value = m     'assign this runner a position on their team based on their place
                        rs(1).Value = 0     'reset their team place
                        rs.Update
                        rs.MoveNext
                    Loop
                    rs.Close
                    Set rs = Nothing
                Case "Time"
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    sql = "SELECT ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON r.RosterID = ir.RosterID WHERE  ir.RacesID = "
                    sql = sql & lThisRace & " AND r.TeamsID = " & TeamsArr(0, j) & " AND ir.Place > 0 AND ir.Excludes = 'n' "
                    sql = sql & "AND ir.FnlScnds > 0 AND ir.RaceTime > '00:00' ORDER BY ir.FnlScnds"
                    rs.Open sql, conn, 1, 2
                    If rs.RecordCount > 0 Then
                        k = 2
                        Do While Not rs.EOF
                            If k <= 8 Then TeamsArr(k, j) = rs(0).Value
                            
                            If k <= iNumScore + 1 Then
                                If TeamsArr(1, j) = vbNullString Then
                                    TeamsArr(1, j) = rs(0).Value
                                Else
                                    TeamsArr(1, j) = CSng(ConvertToSeconds(rs(0).Value)) + CSng(ConvertToSeconds(TeamsArr(1, j)))
                                End If
                            End If
                            
                            k = k + 1
                            rs.MoveNext
                        Loop
                        
                        If k < iNumScore + 2 Then
                            TeamsArr(1, j) = "999999"
                        End If
                   End If
                    rs.Close
                    Set rs = Nothing
                Case "Points"
                    'get the points per place
                    sql = "SELECT NumPlaces, Multiplier FROM ScoreByPts WHERE RacesID = " & lThisRace
                    Set rs = conn.Execute(sql)
                    iScoringPlaces = rs(0).Value
                    iMultiplier = rs(1).Value
                    Set rs = Nothing
                    
                    'set the team score to 0
                    For k = 0 To UBound(TeamsArr, 2)
                        TeamsArr(1, k) = "0"
                    Next

                    'get results
                    If sOrderResultsBy = "Time" Then
                        k = 2
                        For m = 0 To UBound(IndRsltsArr, 2)
                            If CLng(IndRsltsArr(0, m)) = CLng(TeamsArr(0, j)) Then
                                If k <= 8 Then
                                    'check to see if this is a duplicate time
                                    For n = 0 To UBound(DuplTimes, 2) - 1
                                        'if this is a duplicate then score accordingly
                                        If IndRsltsArr(2, m) = DuplTimes(0, n) Then
                                            TeamsArr(k, j) = DuplTimes(3, n)
                                            bIsDupl = True
                                        End If
                                    Next
                                    
                                    If bIsDupl = False Then TeamsArr(k, j) = IndRsltsArr(1, m)
                                    bIsDupl = False
                                End If
                                k = k + 1
                            End If
                        Next
                    Else
                        Set rs = Server.CreateObject("ADODB.Recordset")
                        sql = "SELECT r.TeamsID, ir.Place, ir.ElpsdTime FROM IndRslts ir INNER JOIN Roster r ON r.RosterID = ir.RosterID WHERE  ir.RacesID = "
                        sql = sql & lThisRace & " AND ir.Place > 0 AND ir.RaceTime > '00:00' AND FnlScnds > 0 AND ir.Excludes = 'n' ORDER BY ir.Place"
                        rs.Open sql, conn, 1, 2
                        iRacePlace = 1
                        k = 2
                        Do While Not rs.EOF
                            If rs(0).Value = TeamsArr(0, j) Then
                                If k <= 8 Then TeamsArr(k, j) = iRacePlace
                                k = k + 1
                            End If
                             
                            rs.MoveNext
                            iRacePlace = iRacePlace + 1
                        Loop
                         
                        rs.Close
                        Set rs = Nothing
                    End If
                Case "Pursuit"
                    'get the points per place
                    sql = "SELECT NumPlaces, Multiplier FROM ScoreByPts WHERE RacesID = " & lThisRace
                    Set rs = conn.Execute(sql)
                    iScoringPlaces = rs(0).Value
                    iMultiplier = rs(1).Value
                    Set rs = Nothing
                    
                    'set the team score to 0
                    For k = 0 To UBound(TeamsArr, 2)
                        TeamsArr(1, k) = "0"
                    Next

                    'get results
                    If sOrderResultsBy = "Time" Then
                        k = 2
                        For m = 0 To UBound(IndRsltsArr, 2) 
                            If IndRsltsArr(0, m) = TeamsArr(0, j) Then
                                If k <= 8 Then
                                    'check to see if this is a duplicate time
                                    For n = 0 To UBound(DuplTimes, 2) - 1
                                        'if this is a duplicate then score accordingly
                                        If IndRsltsArr(2, m) = DuplTimes(0, n) Then
                                            TeamsArr(k, j) = DuplTimes(3, n)
                                            bIsDupl = True
                                        End If
                                    Next
                                    
                                    If bIsDupl = False Then TeamsArr(k, j) = IndRsltsArr(1, m)
                                    bIsDupl = False
                                End If
                                k = k + 1
                            End If
                        Next
                    Else
                        Set rs = Server.CreateObject("ADODB.Recordset")
                        sql = "SELECT r.TeamsID, ir.Place, ir.ElpsdTime FROM IndRslts ir INNER JOIN Roster r ON r.RosterID = ir.RosterID WHERE  ir.RacesID = "
                        sql = sql & lThisRace & " AND ir.Place > 0 AND ir.ElpsdTime > '00:00' AND FnlScnds > 0 AND ir.Excludes = 'n' ORDER BY ir.Place"
                        rs.Open sql, conn, 1, 2
                        iRacePlace = 1
                        k = 2
                        Do While Not rs.EOF
                            If rs(0).Value = TeamsArr(0, j) Then
                                If k <= 8 Then TeamsArr(k, j) = iRacePlace
                                k = k + 1
                            End If
                             
                            rs.MoveNext
                            iRacePlace = iRacePlace + 1
                        Loop
                         
                        rs.Close
                        Set rs = Nothing
                    End If
            End Select
        Next

        'sort the time and place arrays
        If ScoreMeth() = "Points" Or ScoreMeth() = "Pursuit" Then
            'get team scores
             For n = 0 To UBound(TeamsArr, 2)
                For k = 2 To 8
                    If k <= iNumScore + 1 Then
                        If TeamsArr(k, n) & "" <> "" Then
                             If CSng(TeamsArr(k, n)) <= iScoringPlaces Then
                                TeamsArr(1, n) = CSng(TeamsArr(1, n)) + (iScoringPlaces - CSng(TeamsArr(k, n)) + 1) * iMultiplier
                            End If
                        End If
                     End If
                Next
            Next
            
            'sort the scores and break any ties
            For n = 0 To UBound(TeamsArr, 2) - 1
                 For k = n + 1 To UBound(TeamsArr, 2)
                     If TeamsArr(1, n) & "" <> "" And TeamsArr(1, k) & "" <> "" Then
                         If CSng(TeamsArr(1, n)) <= CSng(TeamsArr(1, k)) Then
                            If CSng(TeamsArr(1, n)) < CSng(TeamsArr(1, k)) Then
                                For m = 0 To 8
                                    TempArr(m) = TeamsArr(m, n)
                                    TeamsArr(m, n) = TeamsArr(m, k)
                                    TeamsArr(m, k) = TempArr(m)
                                Next
                            Else    'tie breaker here
                                If BreakTie(CLng(TeamsArr(0, n)), CLng(TeamsArr(0, k)), CLng(lThisRace)) = True Then
                                    'tie breaker only executes if a switch of teams is in order
                                    For m = 0 To 8
                                        TempArr(m) = TeamsArr(m, j)
                                        TeamsArr(m, j) = TeamsArr(m, k)
                                        TeamsArr(m, k) = TempArr(m)
                                    Next
                                End If
                            End If
                          End If
                      End If
                  Next
              Next
    
              'if they have a finisher, sort the data and insert their scores
              For n = 0 To UBound(TeamsArr, 2)
                 If TeamsArr(2, n) & "" <> "" Then
                     sql = "INSERT INTO TmRslts(RacesID, TeamsID, Score, R1, R2, R3, R4, R5, R6, R7) VALUES (" & lThisRace & ", "
                     sql = sql & TeamsArr(0, n) & ", '" & TeamsArr(1, n) & "', '" & TeamsArr(2, n) & "', '" & TeamsArr(3, n) & "', '"
                     sql = sql & TeamsArr(4, n) & "', '" & TeamsArr(5, n) & "', '" & TeamsArr(6, n) & "', '" & TeamsArr(7, n) & "', '"
                     sql = sql & TeamsArr(8, n) & "')"
                     Set rs = conn.Execute(sql)
                     Set rs = Nothing
                 End If
             Next
        ElseIf ScoreMeth() = "Time" Then
             For j = 0 To UBound(TeamsArr, 2) - 2
                 For k = j + 1 To UBound(TeamsArr, 2) - 1
                     If TeamsArr(1, j) & "" <> "" And TeamsArr(1, k) & "" <> "" Then
                         If CSng(TeamsArr(1, j)) >= CSng(TeamsArr(1, k)) Then
                             If CSng(TeamsArr(1, j)) > CSng(TeamsArr(1, k)) Then
                                 For m = 0 To 8
                                    TempArr(m) = TeamsArr(m, j)
                                    TeamsArr(m, j) = TeamsArr(m, k)
                                    TeamsArr(m, k) = TempArr(m)
                                 Next
                             Else
                                 If Not TeamsArr(1, j) = "9999" Then
                                     If BreakTie(CLng(TeamsArr(0, j)), CLng(TeamsArr(0, k)), CLng(lThisRace)) = True Then
                                         'tie breaker only executes if a switch of teams is in order
                                           For m = 0 To 8
                                             TempArr(m) = TeamsArr(m, j)
                                             TeamsArr(m, j) = TeamsArr(m, k)
                                             TeamsArr(m, k) = TempArr(m)
                                          Next
                                     End If
                                 End If
                            End If
                         End If
                     End If
                 Next
             Next
            
             For j = 0 To UBound(TeamsArr, 2) - 1
                 If TeamsArr(1, j) = "9999" Or TeamsArr(1, j) = "999999" Then TeamsArr(1, j) = "inc"
             Next
             
            'if they have a finisher, sort the data and insert their scores
            For j = 0 To UBound(TeamsArr, 2) - 1
                If TeamsArr(2, j) & "" <> "" Then
                   If TeamsArr(1, j) = "inc" Then
                       sql = "INSERT INTO TmRslts(RacesID, TeamsID, Score, R1, R2, R3, R4, R5, R6, R7) VALUES (" & lThisRace & ", "
                       sql = sql & TeamsArr(0, j) & ", '" & TeamsArr(1, j) & "', '" & TeamsArr(2, j) & "', '" & TeamsArr(3, j) & "', '"
                       sql = sql & TeamsArr(4, j) & "', '" & TeamsArr(5, j) & "', '" & TeamsArr(6, j) & "', '" & TeamsArr(7, j) & "', '" & TeamsArr(8, j) & "')"
                       Set rs = conn.Execute(sql)
                       Set rs = Nothing
                  Else
                       sql = "INSERT INTO TmRslts(RacesID, TeamsID, Score, R1, R2, R3, R4, R5, R6, R7) VALUES (" & lThisRace & ", "
                       sql = sql & TeamsArr(0, j) & ", '" & ConvertToMinutes(TeamsArr(1, j)) & "', '" & TeamsArr(2, j) & "', '" & TeamsArr(3, j) & "', '"
                       sql = sql & TeamsArr(4, j) & "', '" & TeamsArr(5, j) & "', '" & TeamsArr(6, j) & "', '" & TeamsArr(7, j) & "', '" & TeamsArr(8, j) & "')"
                       Set rs = conn.Execute(sql)
                       Set rs = Nothing
                   End If
                End If
            Next
        Else    'score method = "Place"
            'get the teams that have at least the minimum finishers
            ReDim ComplTeams(8, 0)
            k = 0
            For j = 0 To UBound(TeamsArr, 2) - 1
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT ir.PlaceOnTeam FROM IndRslts ir INNER JOIN Roster r ON r.RosterID = ir.RosterID WHERE ir.RacesID = "
                sql = sql & lThisRace & " AND r.TeamsID = " & TeamsArr(0, j) & "  AND ir.RaceTime > '00:00' AND ir.Excludes = 'n' "
                sql = sql & "AND ir.Place > 0 AND ir.FnlScnds > 0 ORDER BY ir.PlaceOnTeam DESC"
                rs.Open sql, conn, 1, 2
                If rs.RecordCount > 0 Then
                    If rs(0).Value >= iNumScore Then
                        ComplTeams(0, k) = TeamsArr(0, j)
                        ComplTeams(1, k) = 0
                        For m = 2 To 8
                            ComplTeams(m, k) = "--"
                        Next
                        k = k + 1
                        ReDim Preserve ComplTeams(8, k)
                    End If
                End If
                rs.Close
                Set rs = Nothing
            Next

            'first assign a team place for each competitor
            k = 1
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT r.TeamsID, ir.TeamPlace FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE RacesID = "
            sql = sql & lThisRace & " AND ir.Place > 0 AND ir.Excludes = 'n' AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.PlaceOnTeam <= 7 "
            sql = sql & "ORDER BY ir.FnlScnds"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                For j = 0 To UBound(ComplTeams, 2) - 1
                    If CLng(rs(0).Value) = CLng(ComplTeams(0, j)) Then
                        rs(1).Value = k
                        rs.Update
                        k = k + 1
                        Exit For
                    End If
                Next
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
            
            For j = 0 To UBound(ComplTeams, 2) - 1
                k = 2
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT ir.TeamPlace FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = "
                sql = sql & lThisRace & " AND r.TeamsID = " & ComplTeams(0, j) & " AND ir.TeamPlace IS NOT NULL AND ir.TeamPlace > 0 "
                sql = sql & "ORDER BY ir.TeamPlace"
                rs.Open sql, conn, 1, 2
                Do While Not rs.EOF
                    ComplTeams(k, j) = rs(0).Value
                    If k <= iNumScore + 1 Then ComplTeams(1, j) = CInt(ComplTeams(1, j)) + rs(0).Value
                    
                    If k > 7 Then
                        Exit Do
                    Else
                        k = k + 1
                        rs.MoveNext
                    End If
                Loop
                rs.Close
                Set rs = Nothing
            Next
            
            'sort the scores and break any ties
            For n = 0 To UBound(ComplTeams, 2) - 2
                 For k = n + 1 To UBound(ComplTeams, 2) - 1
                    If CSng(ComplTeams(1, n)) >= CSng(ComplTeams(1, k)) Then
                       If CSng(ComplTeams(1, n)) > CSng(ComplTeams(1, k)) Then
                           For m = 0 To 8
                               TempArr(m) = ComplTeams(m, n)
                               ComplTeams(m, n) = ComplTeams(m, k)
                               ComplTeams(m, k) = TempArr(m)
                           Next
                       Else    'tie breaker here
                           If BreakTie(CLng(ComplTeams(0, n)), CLng(ComplTeams(0, k)), CLng(lThisRace)) = True Then
                               'tie breaker only executes if a switch of teams is in order
                               For m = 0 To 8
                                   TempArr(m) = ComplTeams(m, n)
                                   ComplTeams(m, n) = ComplTeams(m, k)
                                   ComplTeams(m, k) = TempArr(m)
                               Next
                           End If
                       End If
                     End If
                  Next
              Next
            
            'insert the scores
            For j = 0 To UBound(ComplTeams, 2) - 1
                sql = "INSERT INTO TmRslts(RacesID, TeamsID, Score, R1, R2, R3, R4, R5, R6, R7) VALUES (" & lThisRace & ", "
                sql = sql & ComplTeams(0, j) & ", '" & ComplTeams(1, j) & "', '" & ComplTeams(2, j) & "', '" & ComplTeams(3, j) & "', '"
                sql = sql & ComplTeams(4, j) & "', '" & ComplTeams(5, j) & "', '" & ComplTeams(6, j) & "', '" & ComplTeams(7, j) & "', '"
                sql = sql & ComplTeams(8, j) & "')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            Next
        End If
        
        For j = 0 To UBound(TeamsArr, 2) - 1
            For k = 1 To 8
                TeamsArr(k, j) = vbNullString
            Next
        Next
    End If
End Sub

Public Function BreakTie(lTeam1, lTeam2, lThisRace)
    Dim x, y
    Dim lTieBreakerID
    Dim Team1Arr(7), Team2Arr(7)
    Dim sngTeam1Cum, sngTeam2Cum
            
    BreakTie = False
    
    sql = "SELECT TieBreakersID FROM Races WHERE RacesID = " & lThisRace
    Set rs = conn.Execute(sql)
    lTieBreakerID = rs(0).Value
    Set rs = Nothing
    
    Select Case CLng(lTieBreakerID)
        Case 1      'mshsl nordic method
            'combined times of the teams top 4 skiers
            'then fastest 5th skier breaks the tie
            'then fastest 6th skier breaks the tie
            'then fastest 7th skier breaks the tie
            
            'get both teams top 6 skiers
            x = 0
            sql = "SELECT ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
            sql = sql & "WHERE ir.RacesID = " & lThisRace & " AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Place > 0 ANd ir.Excludes = 'n' "
            sql = sql & "AND r.TeamsID = " & lTeam1 & " ORDER BY ir.FnlScnds"
            Set rs = conn.Execute(sql)
            Do While Not rs.EOF
                Team1Arr(x) = rs(0).Value
                x = x + 1
                If x = 7 Then Exit Do
                rs.MoveNext
            Loop
            Set rs = Nothing
            
            x = 0
            sql = "SELECT ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
            sql = sql & "WHERE ir.RacesID = " & lThisRace & " AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Place > 0 ANd ir.Excludes = 'n' "
            sql = sql & "AND r.TeamsID = " & lTeam2 & " ORDER BY ir.FnlScnds"
            Set rs = conn.Execute(sql)
            Do While Not rs.EOF
                Team2Arr(x) = rs(0).Value
                x = x + 1
                If x = 7 Then Exit Do
                rs.MoveNext
            Loop
            Set rs = Nothing
            
            'compare the teams' top 4 cumulative times
            sngTeam1Cum = 0
            sngTeam2Cum = 0
            For x = 0 To 3
                sngTeam1Cum = sngTeam1Cum + ConvertToSeconds(Team1Arr(x))
                sngTeam2Cum = sngTeam2Cum + ConvertToSeconds(Team2Arr(x))
            Next
            
            If sngTeam1Cum < sngTeam2Cum Then
                BreakTie = False
                Exit Function
            ElseIf sngTeam2Cum < sngTeam1Cum Then
                BreakTie = True
                Exit Function
            End If
            
            'if neither team has an advantage
            For x = 4 To 6
                If Team1Arr(x) = vbNullString Then   'if team 1 does not have the appropriate skier
                    If Not Team2Arr(x) = vbNullString Then   'and team 2 does, we break the tie
                        BreakTie = True
                        Exit For
                    Else    'and team 2 does not either then the top skier breaks the tie
                        'both teams are empty at this position so we go to top skier
                        If ConvertToSeconds(Team1Arr(0)) <= ConvertToSeconds(Team2Arr(0)) Then
                            BreakTie = False
                            Exit For
                        Else
                            BreakTie = True
                            Exit For
                        End If
                    End If
                Else    'if team 1 does have this skier
                    If Team2Arr(x) = vbNullString Then   'and team 2 doesn't the results stay put
                        BreakTie = False
                        Exit For
                    Else    'and team 2 does too then compare them
                        If ConvertToSeconds(Team2Arr(x)) < ConvertToSeconds(Team1Arr(x)) Then
                            BreakTie = True
                            Exit For
                        Else
                            If ConvertToSeconds(Team1Arr(x)) < ConvertToSeconds(Team2Arr(x)) Then
                                BreakTie = False
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
        Case 2      'mshsl cc running system
            'first 6th runner breaks tie
            'if neither team has a 6th runner first 5th runner breaks the tie
            x = 0
            y = 0
            sql = "SELECT r.TeamsID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " & lThisRace 
            sql = sql & " AND ir.Place > 0 AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Excludes = 'n' ORDER BY ir.FnlScnds"
            Set rs = conn.Execute(sql)
            Do While Not rs.EOF
                If CLng(rs(0).Value) = CLng(lTeam1) Then
                    x = x + 1
                    If x = 6 Then Exit Do
                ElseIf CLng(rs(0).Value) = CLng(lTeam2) Then
                    y = y + 1
                    If y = 6 Then Exit Do
                End If
                rs.MoveNext
            Loop
            Set rs = Nothing
            
            If x = 6 Then
                BreakTie = False
            ElseIf y = 6 Then
                BreakTie = True
            Else
                'if neither team has 6th runner then 5th runner breaks tie
                x = 0
                y = 0
                sql = "SELECT r.TeamsID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " & lThisRace 
                sql = sql & " AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Place > 0 ANd ir.Excludes = 'n'  ORDER BY ir.FnlScnds"
                Set rs = conn.Execute(sql)
                Do While Not rs.EOF
                    If CLng(rs(0).Value) = CLng(lTeam1) Then
                        x = x + 1
                        If x = 5 Then Exit Do
                    ElseIf CLng(rs(0).Value) = CLng(lTeam2) Then
                        y = y + 1
                        If y = 5 Then Exit Do
                    End If
                    rs.MoveNext
                Loop
                Set rs = Nothing
                
                If x = 5 Then
                    BreakTie = False
                Else
                    BreakTie = True
                End If
            End If
        Case 3      'clc nordic reg season method
            'compare the teams' 7th skier
            'compare the teams' 8th skier
            'compare the top skier
            
            'get both teams top 8 skiers
            x = 0
            sql = "SELECT ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " & lThisRace 
            sql = sql & " AND AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Place > 0 ANd ir.Excludes = 'n' AND r.TeamsID = " & lTeam1 
            sql = sql & " ORDER BY ir.FnlScnds"
            Set rs = conn.Execute(sql)
            Do While Not rs.EOF
                Team1Arr(x) = rs(0).Value
                x = x + 1
                If x = 9 Then Exit Do
                rs.MoveNext
            Loop
            Set rs = Nothing
            
            x = 0
            sql = "SELECT ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " & lThisRace 
            sql = sql & " AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Place > 0 ANd ir.Excludes = 'n'  AND r.TeamsID = " & lTeam2 
            sql = sql & " ORDER BY ir.FnlScnds"
            Set rs = conn.Execute(sql)
            Do While Not rs.EOF
                Team2Arr(x) = rs(0).Value
                x = x + 1
                If x = 9 Then Exit Do
                rs.MoveNext
            Loop
            Set rs = Nothing
            
            For x = 6 To 7
                If Team1Arr(x) = vbNullString Then   'if team 1 does not have the appropriate skier
                    If Not Team2Arr(x) = vbNullString Then   'and team 2 does, we break the tie
                        BreakTie = True
                        Exit For
                    Else    'and team 2 does not either then the top skier breaks the tie
                        'both teams are empty at this position so we go to top skier
                        If ConvertToSeconds(Team1Arr(0)) <= ConvertToSeconds(Team2Arr(0)) Then
                            BreakTie = False
                            Exit For
                        Else
                            BreakTie = True
                            Exit For
                        End If
                    End If
                Else    'if team 1 does have this skier
                    If Team2Arr(x) = vbNullString Then   'and team 2 doesn't the results stay put
                        BreakTie = False
                        Exit For
                    Else    'and team 2 does too then compare them
                        If ConvertToSeconds(Team2Arr(x)) < ConvertToSeconds(Team1Arr(x)) Then
                            BreakTie = True
                            Exit For
                        Else
                            If ConvertToSeconds(Team1Arr(x)) < ConvertToSeconds(Team2Arr(x)) Then
                                BreakTie = False
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
        Case 4      'clc conf champ method
            'compare the teams' 5th skier
            'compare the teams' 6th skier
            'compare the top skier
            
            'get both teams top 6 skiers
            x = 0
            sql = "SELECT ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " & lThisRace 
            sql = sql & " AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Place > 0 ANd ir.Excludes = 'n' AND r.TeamsID = " & lTeam1 
            sql = sql & " ORDER BY ir.FnlScnds"
            Set rs = conn.Execute(sql)
            Do While Not rs.EOF
                Team1Arr(x) = rs(0).Value
                x = x + 1
                If x = 6 Then Exit Do
                rs.MoveNext
            Loop
            Set rs = Nothing
            
            x = 0
            sql = "SELECT ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " & lThisRace 
            sql = sql & " AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 AND ir.Place > 0 ANd ir.Excludes = 'n' AND r.TeamsID = " & lTeam2 
            sql = sql & " ORDER BY ir.RaceTime"
            Set rs = conn.Execute(sql)
            Do While Not rs.EOF
                Team2Arr(x) = rs(0).Value
                x = x + 1
                If x = 6 Then Exit Do
                rs.MoveNext
            Loop
            Set rs = Nothing
            
            For x = 4 To 5
                If Team1Arr(x) = vbNullString Then   'if team 1 does not have the appropriate skier
                    If Not Team2Arr(x) = vbNullString Then   'and team 2 does, we break the tie
                        BreakTie = True
                        Exit For
                    Else    'and team 2 does not either then the top skier breaks the tie
                        'both teams are empty at this position so we go to top skier
                        If ConvertToSeconds(Team1Arr(0)) <= ConvertToSeconds(Team2Arr(0)) Then
                            BreakTie = False
                            Exit For
                        Else
                            BreakTie = True
                            Exit For
                        End If
                    End If
                Else    'if team 1 does have this skier
                    If Team2Arr(x) = vbNullString Then   'and team 2 doesn't the results stay put
                        BreakTie = False
                        Exit For
                    Else    'and team 2 does too then compare them
                        If ConvertToSeconds(Team2Arr(x)) < ConvertToSeconds(Team1Arr(x)) Then
                            BreakTie = True
                            Exit For
                        Else
                            If ConvertToSeconds(Team1Arr(x)) < ConvertToSeconds(Team2Arr(x)) Then
                                BreakTie = False
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
    End Select
End Function

%>
<!--#include file = "../../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../../includes/convert_to_minutes.asp" -->
<%

Public Function ScoreMeth()
    sql = "SELECT ScoreMethod FROM Races WHERE RacesID = " & lThisRace
    Set rs = conn.Execute(sql)
    ScoreMeth = rs(0).Value
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE CC/Nordic Results Manager: Update Team Scores</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "../manage_meet_nav.asp" -->
			<%End If%>

			<h4 class="h4">Update Team Scores for <%=sMeetName%> on <%=dMeetDate%>:&nbsp;<%=sRaceName%></h4>
					
			<form class="form-inline bg-success" name="get_races" method="post" action="update_team_scores.asp?meet_id=<%=lThisMeet%>">
			<label for="races">Select Race:</label>
			<select class="form-control" name="races" id="races" onchange="this.form.get_race.click();">
				<%For i = 0 to UBound(Races, 2) - 1%>
					<%If CLng(lThisRace) = CLng(Races(0, i)) Then%>
						<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
					<%Else%>
						<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
			<input type="submit" class="form-control" name="get_race" id="get_race" value="Get Results">
			</form>

			<!--#include file = "results_nav.asp" -->

            <h4 class="h4">Team Results: <%=sRaceName%></h4>
			
            <div class="bg-danger">
                <a href="update_team_scores.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>&amp;update_these=y">Update These</a>
            </div>		

			<table class="table tablbe-striped">
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
						<td><%=TmRslts(0, i)%></td>
						<td><%=TmRslts(1, i)%></td>
						<td><%=TmRslts(2, i)%></td>
						<td><%=TmRslts(3, i)%></td>
						<td><%=TmRslts(4, i)%></td>
						<td><%=TmRslts(5, i)%></td>
						<td><%=TmRslts(6, i)%></td>
						<td><%=TmRslts(7, i)%></td>
						<td><%=TmRslts(8, i)%></td>
					</tr>
				<%Next%>
			</table>
		</div>
    </div>
    <!--#include file = "../../../includes/footer.asp" -->	
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
