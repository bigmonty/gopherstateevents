<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j, k
Dim lPacksID
Dim sPackName, sSport, sGender, sMeetName, sSite, sSortBy, sBlendBy, sSelectQS, sCompOnly, sThisSite, sExclSites, sExclMeets
Dim PackMmbrs(), ViewPerf(), RsltsArr(), TempArr(6), SelUsers(), Meets(), MeetRslts(), SiteRslts(), MeetSites(), AllMeets()
Dim dWhenCreated, dMeetDate, dBegDate, dEndDate
Dim bGetRslts

lPacksID = Request.QueryString("packs_id")
sSortBy = Request.QueryString("sort_by")
sBlendBy = Request.QueryString("blend_by")
sSelectQS = Request.QueryString("select_qs")
sExclSites = Request.QueryString("excl_sites")
sExclMeets = Request.QueryString("excl_meets")
sCompOnly = Request.QueryString("comp_only")
dBegDate = Request.QueryString("beg_date")
dEndDate = Request.QueryString("end_date")

Server.ScriptTimeout = 600

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

ReDim ViewPerf(3, 0)

If Not sSelectQS = vbNullString Then
    j = 0
    ReDim SelUsers(0)
    If Not CStr(sSelectQS) = vbNullString Then
	    For i = 1 To Len(sSelectQS)
		    If Mid(sSelectQS, i, 1) = ";" Then
			    SelUsers(j) = Trim(CStr(SelUsers(j)))
			    j = j + 1
			    ReDim Preserve SelUsers(j)
		    Else
			    SelUsers(j) = SelUsers(j) & Mid(sSelectQS, i, 1)
		    End If
	    Next
    End If

    i = 0
    ReDim ViewPerf(3, 0)
    For j = 0 To UBound(SelUsers) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT pt.RosterID, r.FirstName, r.LastName, r.TeamsID FROM PTPackMmbrs pt INNER JOIN Roster r ON pt.RosterID = r.RosterID "
        sql = sql & "WHERE pt.PerfTrkrPacksID = " & lPacksID & " AND pt.PTPackMmbrsID = " & SelUsers(j) & " ORDER BY r.LastName, r.FirstName"
        rs.Open sql, conn2, 1, 2
        ViewPerf(0, i) = SelUsers(j)
        ViewPerf(1, i) = rs(0).Value
        ViewPerf(2, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        ViewPerf(3, i) = GetTeamName(rs(3).Value)
        rs.Close
        Set rs = Nothing

        i = i + 1
        ReDim Preserve ViewPerf(3, i)
    Next
End If

Dim ExclMeets()
ReDim ExclMeets(0)
If Not sExclMeets = vbNullString Then
    j = 0
    If Not CStr(sExclMeets) = vbNullString Then
	    For i = 1 To Len(sExclMeets)
		    If Mid(sExclMeets, i, 1) = "," Then
			    ExclMeets(j) = Trim(CStr(ExclMeets(j)))
			    j = j + 1
			    ReDim Preserve ExclMeets(j)
		    Else
			    If i = Len(sExclMeets) Then
			        ExclMeets(j) = Trim(CStr(ExclMeets(j))) & Mid(sExclMeets, i, 1)
			        j = j + 1
			        ReDim Preserve ExclMeets(j)
                Else
                    ExclMeets(j) = ExclMeets(j) & Mid(sExclMeets, i, 1)
                End If
		    End If
	    Next
    End If
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PackName, Sport, Gender, WhenCreated FROM PerfTrkrPacks WHERE PerfTrkrPacksID = " & lPacksID
rs.Open sql, conn2, 1, 2
sPackName = Replace(rs(0).Value, "''", "'")
sSport = rs(1).Value
sGender = rs(2).Value
dWhenCreated = rs(3).Value
rs.Close
Set rs = Nothing

i = 0
ReDim PackMmbrs(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT pt.PTPackMmbrsID, pt.RosterID, r.FirstName, r.LastName, r.TeamsID FROM PTPackMmbrs pt INNER JOIN Roster r ON pt.RosterID = r.RosterID "
sql = sql & "WHERE pt.PerfTrkrPacksID = " & lPacksID & " ORDER BY r.LastName, r.FirstName"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    PackMmbrs(0, i) = rs(0).Value
    PackMmbrs(1, i) = rs(1).Value
    PackMmbrs(2, i) = Replace(rs(3).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
    PackMmbrs(3, i) = GetTeamName(rs(4).Value)
    i = i + 1
    ReDim Preserve PackMmbrs(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Month(Date) < 9 Then
    If CStr(dBegDate) = vbNullString Then dBegDate = "8/1/" & Year(Date) - 1
    If CStr(dEndDate) = vbNullString Then dEndDate = "3/1/" & Year(Date)
Else
    If CStr(dBegDate) = vbNullString Then dBegDate = "8/1/" & Year(Date)
    If CStr(dEndDate) = vbNullString Then dEndDate = "3/1/" & Year(Date) + 1
End If

Call GetMeets
Call GetMeetSites

If Not sBlendBy = "part" Then 
    'sort according to display
    If sSortBy = "date" Then
        For j = 0 To UBound(Meets, 2) - 2
            For k = j + 1 To UBound(Meets, 2) - 1
                If CDate(Meets(2, j)) < CDate(Meets(2, k)) Then
                    For i = 0 To 3
                        TempArr(i) = Meets(i, j)
                        Meets(i, j) = Meets(i, k)
                        Meets(i, k) = TempArr(i)
                    Next
                End If
            Next
        Next
    ElseIf sSortBy = "meet" Then
        For j = 0 To UBound(Meets, 2) - 2
            For k = j + 1 To UBound(Meets, 2) - 1
                If CStr(Meets(1, j)) > CStr(Meets(1, k)) Then
                    For i = 0 To 3
                        TempArr(i) = Meets(i, j)
                        Meets(i, j) = Meets(i, k)
                        Meets(i, k) = TempArr(i)
                    Next
                End If
            Next
        Next
    ElseIf sSortBy = "site" Then
        For j = 0 To UBound(Meets, 2) - 2
            For k = j + 1 To UBound(Meets, 2) - 1
                If CStr(Meets(3, j)) > CStr(Meets(3, k)) Then
                    For i = 0 To 3
                        TempArr(i) = Meets(i, j)
                        Meets(i, j) = Meets(i, k)
                        Meets(i, k) = TempArr(i)
                    Next
                End If
            Next
        Next
    End If
End If

Private Sub GetMeets()
    Dim x, y, z

    'get all meets that these participants are in
    x = 0
    ReDim AllMeets(3, 0)
    For y = 0 To UBound(ViewPerf, 2) - 1
        'first get all meets to populate the list box
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate, m.MeetSite FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID INNER JOIN Roster r ON "
        sql = sql & "mt.TeamsID = r.TeamsID INNER JOIN IndRslts ir ON ir.MeetsID = mt.MeetsID WHERE ir.RosterID = " & ViewPerf(1, y) & " AND ir.Place > 0 "
        sql = sql & "AND (m.MeetDate >= '" & dBegDate & "' AND m.MeetDate <= '" & dEndDate & "') AND Sport = '" & sSport & "' ORDER BY m.MeetSite DESC"
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If x = 0 Then
                AllMeets(0, x) = rs(0).Value
                AllMeets(1, x) = Replace(rs(1).Value, "''", "'")
                AllMeets(2, x) = rs(2).Value
                AllMeets(3, x) = Replace(rs(3).Value, "''", "'")
                x = x + 1
                ReDim Preserve AllMeets(3, x)
            Else
                For z = 0 To UBound(AllMeets, 2) - 1
                    If CLng(rs(0).Value) = CLng(AllMeets(0, z)) Then
                        Exit For
                    Else
                        If z = UBound(AllMeets, 2) - 1 Then
                            AllMeets(0, x) = rs(0).Value
                            AllMeets(1, x) = Replace(rs(1).Value, "''", "'")
                            AllMeets(2, x) = rs(2).Value
                            AllMeets(3, x) = Replace(rs(3).Value, "''", "'")
                            x = x + 1
                            ReDim Preserve AllMeets(3, x)
                        End If
                    End If
                Next
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next

    'get meets that these participants are in
    x = 0
    ReDim Meets(3, 0)
    For y = 0 To UBound(ViewPerf, 2) - 1
        'first get all meets to populate the list box
        Set rs = Server.CreateObject("ADODB.Recordset")
        If sExclSites = vbNullString Then
            If sExclMeets = vbNullString Then
                sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate, m.MeetSite FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID INNER JOIN Roster r ON "
                sql = sql & "mt.TeamsID = r.TeamsID INNER JOIN IndRslts ir ON ir.MeetsID = mt.MeetsID WHERE ir.RosterID = " & ViewPerf(1, y) & " AND ir.Place > 0 "
                sql = sql & "AND (m.MeetDate >= '" & dBegDate & "' AND m.MeetDate <= '" & dEndDate & "') AND Sport = '" & sSport & "' ORDER BY m.MeetSite DESC"
            Else
                sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate, m.MeetSite FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID INNER JOIN Roster r ON "
                sql = sql & "mt.TeamsID = r.TeamsID INNER JOIN IndRslts ir ON ir.MeetsID = mt.MeetsID WHERE ir.RosterID = " & ViewPerf(1, y) & " AND ir.Place > 0 "
                sql = sql & "AND (m.MeetDate >= '" & dBegDate & "' AND m.MeetDate <= '" & dEndDate & "') AND Sport = '" & sSport & "' AND m.MeetsID NOT IN ("
                sql = sql & sExclMeets & ") ORDER BY m.MeetSite DESC"
            End If
        Else
            If sExclMeets = vbNullString Then
                sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate, m.MeetSite FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID INNER JOIN Roster r ON "
                sql = sql & "mt.TeamsID = r.TeamsID INNER JOIN IndRslts ir ON ir.MeetsID = mt.MeetsID WHERE ir.RosterID = " & ViewPerf(1, y) & " AND ir.Place > 0 "
                sql = sql & "AND (m.MeetDate >= '" & dBegDate & "' AND m.MeetDate <= '" & dEndDate & "') AND Sport = '" & sSport & "' AND m.MeetSite NOT IN ("
                sql = sql & sExclSites & ") ORDER BY m.MeetSite DESC"
            Else
                sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate, m.MeetSite FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID INNER JOIN Roster r ON "
                sql = sql & "mt.TeamsID = r.TeamsID INNER JOIN IndRslts ir ON ir.MeetsID = mt.MeetsID WHERE ir.RosterID = " & ViewPerf(1, y) & " AND ir.Place > 0 "
                sql = sql & "AND (m.MeetDate >= '" & dBegDate & "' AND m.MeetDate <= '" & dEndDate & "') AND Sport = '" & sSport & "' AND m.MeetSite NOT IN ("
                sql = sql & sExclSites & ") AND m.MeetsID NOT IN (" & sExclMeets & ") ORDER BY m.MeetSite DESC"
            End If
        End If
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If x = 0 Then
                Meets(0, x) = rs(0).Value
                Meets(1, x) = Replace(rs(1).Value, "''", "'")
                Meets(2, x) = rs(2).Value
                Meets(3, x) = Replace(rs(3).Value, "''", "'")
                x = x + 1
                ReDim Preserve Meets(3, x)
            Else
                For z = 0 To UBound(Meets, 2) - 1
                    If CLng(rs(0).Value) = CLng(Meets(0, z)) Then
                        Exit For
                    Else
                        If z = UBound(Meets, 2) - 1 Then
                            Meets(0, x) = rs(0).Value
                            Meets(1, x) = Replace(rs(1).Value, "''", "'")
                            Meets(2, x) = rs(2).Value
                            Meets(3, x) = Replace(rs(3).Value, "''", "'")
                            x = x + 1
                            ReDim Preserve Meets(3, x)
                        End If
                    End If
                Next
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next
End Sub

Private Sub GetMeetSites()
    Dim x, y, z

    x = 0
    ReDim MeetSites(0)
    For y = 0 To UBound(AllMeets, 2) - 1
        If y = 0 Then
            sThisSite = AllMeets(3, y) 
            MeetSites(0) = sThisSite
            x = x + 1
            ReDim Preserve MeetSites(x)
        Else
            If Not CStr(sThisSite) = CStr(AllMeets(3, y)) Then
                sThisSite = AllMeets(3, y) 
                MeetSites(x) = sThisSite
                x = x + 1
                ReDim Preserve MeetSites(x)
            End If
        End If
    Next
End Sub

Private Function GetTeamName(lTeamID)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
    rs2.Open sql2, conn2, 1, 2
    GetTeamName = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

Private Sub MyResults(lThisMmbr)
    Dim x, y, z

    x = 0
    ReDim RsltsArr(6, 0)
	For y = 0 To UBound(Meets, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT ir.RacesID, r.RaceDesc, ir.RaceTime, r.RaceDist, r.RaceUnits FROM IndRslts ir INNER JOIN Races r ON r.RacesID = ir.RacesID "
        sql = sql & "WHERE ir.RosterID = " & lThisMmbr & " AND ir.Place > 0 AND ir.MeetsID = " & Meets(0, y) & " ORDER BY ir.RacesID"
	    rs.Open sql, conn2, 1, 2
	    Do While Not rs.EOF
            Call MeetData(rs(0).Value)

		    RsltsArr(0, x) = sMeetName
		    RsltsArr(1, x) = Month(CDate(dMeetDate)) & "/" & Day(CDate(dMeetDate)) & "/" & Year(CDate(dMeetDate))
            RsltsArr(2, x) = sSite
		    RsltsArr(3, x) = rs(1).Value
            RsltsArr(4, x) = rs(3).Value & " " & rs(4).Value
		    RsltsArr(5, x) = GetPlace(rs(0).Value, lThisMmbr)
		    RsltsArr(6, x) = rs(2).Value
		    x = x + 1
		    ReDim Preserve RsltsArr(6, x)
		    rs.MoveNext
	    Loop
	    rs.Close
	    Set rs = Nothing
    Next

    'sort by date as default first...then run other sorts later if needed
    For x = 0 To UBound(RsltsArr, 2) - 2
        For y = x + 1 To UBound(RsltsArr, 2) - 1
            If CDate(RsltsArr(1, x)) < CDate(RsltsArr(1, y)) Then
                For z = 0 To 6
                    TempArr(z) = RsltsArr(z, x)
                    RsltsArr(z, x) = RsltsArr(z, y)
                    RsltsArr(z, y) = TempArr(z)
                Next
            End If
        Next
    Next

    Select Case sSortBy
        Case "site"
            For x = 0 To UBound(RsltsArr, 2) - 2
                For y = x + 1 To UBound(RsltsArr, 2) - 1
                    If CStr(RsltsArr(2, x)) > CStr(RsltsArr(2, y)) Then
                        For z = 0 To 6
                            TempArr(z) = RsltsArr(z, x)
                            RsltsArr(z, x) = RsltsArr(z, y)
                            RsltsArr(z, y) = TempArr(z)
                        Next
                    End If
                Next
            Next
        Case "date"
            For x = 0 To UBound(RsltsArr, 2) - 2
                For y = x + 1 To UBound(RsltsArr, 2) - 1
                    If CDate(RsltsArr(1, x)) < CDate(RsltsArr(1, y)) Then
                        For z = 0 To 6
                            TempArr(z) = RsltsArr(z, x)
                            RsltsArr(z, x) = RsltsArr(z, y)
                            RsltsArr(z, y) = TempArr(z)
                        Next
                    End If
                Next
            Next
        Case "meet"
            For x = 0 To UBound(RsltsArr, 2) - 2
                For y = x + 1 To UBound(RsltsArr, 2) - 1
                    If CStr(RsltsArr(0, x)) > CStr(RsltsArr(0, y)) Then
                        For z = 0 To 6
                            TempArr(z) = RsltsArr(z, x)
                            RsltsArr(z, x) = RsltsArr(z, y)
                            RsltsArr(z, y) = TempArr(z)
                        Next
                    End If
                Next
            Next
        Case "perf"
            'convert to seconds
            For x = 0 To UBound(RsltsArr, 2) - 1
                RsltsArr(6, x) = ConvertToSeconds(RsltsArr(6, x))
            Next

            'sort
            For x = 0 To UBound(RsltsArr, 2) - 2
                For y = x + 1 To UBound(RsltsArr, 2) - 1
                    If CSng(RsltsArr(6, x)) > CSng(RsltsArr(6, y)) Then
                        For z = 0 To 6
                            TempArr(z) = RsltsArr(z, x)
                            RsltsArr(z, x) = RsltsArr(z, y)
                            RsltsArr(z, y) = TempArr(z)
                        Next
                    End If
                Next
            Next

            'convert to minutes
            For x = 0 To UBound(RsltsArr, 2) - 1
                RsltsArr(6, x) = ConvertToMinutes(RsltsArr(6, x))
            Next
    End Select
End Sub

Private Sub GetMeetRslts(lThisMeet)
    Dim x, y, z

    x = 0
    ReDim MeetRslts(5, 0)
    For y = 0 To UBound(ViewPerf, 2) - 1
	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT r.RaceDesc, ir.RaceTime, ir.RacesID, r.RaceDist, r.RaceUnits FROM IndRslts ir INNER JOIN Races r ON r.RacesID = ir.RacesID WHERE ir.MeetsID = " 
        sql = sql & lThisMeet & " AND ir.RosterID = " & ViewPerf(1, y) & " AND ir.Place > 0 ORDER BY ir.RacesID, ir.RaceTime"
	    rs.Open sql, conn2, 1, 2
	    Do While Not rs.EOF
		    MeetRslts(0, x) = ViewPerf(2, y)
		    MeetRslts(1, x) = ViewPerf(3, y)
            MeetRslts(2, x) = rs(0).Value
            MeetRslts(3, x) = rs(3).Value & " " & rs(4).Value
		    MeetRslts(4, x) = GetPlace(rs(2).Value, ViewPerf(1, y))
            MeetRslts(5, x) = rs(1).Value
		    x = x + 1
		    ReDim Preserve MeetRslts(5, x)
		    rs.MoveNext
	    Loop
	    rs.Close
	    Set rs = Nothing
    Next

    For x = 0 To UBound(MeetRslts, 2) - 2
        For y = x + 1 To UBound(MeetRslts, 2) - 1
            If CStr(MeetRslts(5, x)) > CStr(MeetRslts(5, y)) Then
                For z = 0 To 5
                    TempArr(z) = MeetRslts(z, x)
                    MeetRslts(z, x) = MeetRslts(z, y)
                    MeetRslts(z, y) = TempArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Sub GetSiteRslts(sThisSite)
    Dim x, y, z
    Dim SiteMeets()

    'first sort meet by site
    For x = 0 To UBound(Meets, 2) - 2
        For y = x + 1 To UBound(Meets, 2) - 1
            If CStr(Meets(3, x)) > CStr(Meets(3, y)) Then
                For z = 0 To 3
                    TempArr(z) = Meets(z, x)
                    Meets(z, x) = Meets(z, y)
                    Meets(z, y) = TempArr(z)
                Next
            End If
        Next
    Next

    y = 0
    ReDim SiteMeets(1, 0)
    For x = 0 To UBound(Meets, 2) - 1
        If CStr(Meets(3, x)) = CStr(sThisSite) Then
            SiteMeets(0, y) = Meets(0, x)
            SiteMeets(1, y) = Meets(1, x)
            y = y + 1
            ReDim Preserve SiteMeets(1, y)
        End If
    Next

    x = 0
    ReDim SiteRslts(6, 0)
    For z = 0 To UBound(SiteMeets, 2) - 1
        For y = 0 To UBound(ViewPerf, 2) - 1
	        Set rs = Server.CreateObject("ADODB.Recordset")
	        sql = "SELECT r.RaceDesc, ir.RaceTime, ir.RacesID, r.RaceDist, r.RaceUnits FROM IndRslts ir INNER JOIN Races r ON r.RacesID = ir.RacesID WHERE ir.MeetsID = " 
            sql = sql & SiteMeets(0, z) & " AND ir.RosterID = " & ViewPerf(1, y) & " AND ir.Place > 0 ORDER BY ir.RacesID, ir.RaceTime"
	        rs.Open sql, conn2, 1, 2
	        Do While Not rs.EOF
		        SiteRslts(0, x) = ViewPerf(2, y)
		        SiteRslts(1, x) = ViewPerf(3, y)
                SiteRslts(2, x) = SiteMeets(1, z)
                SiteRslts(3, x) = rs(0).Value
                SiteRslts(4, x) = rs(3).Value & " " & rs(4).Value
		        SiteRslts(5, x) = GetPlace(rs(2).Value, ViewPerf(1, y))
                SiteRslts(6, x) = rs(1).Value
		        x = x + 1
		        ReDim Preserve SiteRslts(6, x)
		        rs.MoveNext
	        Loop
	        rs.Close
	        Set rs = Nothing
        Next
    Next

    For x = 0 To UBound(SiteRslts, 2) - 2
        For y = x + 1 To UBound(SiteRslts, 2) - 1
            If CStr(SiteRslts(6, x)) > CStr(SiteRslts(6, y)) Then
                For z = 0 To 6
                    TempArr(z) = SiteRslts(z, x)
                    SiteRslts(z, x) = SiteRslts(z, y)
                    SiteRslts(z, y) = TempArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Sub MeetData(lThisRace)
    sMeetName = "unknown" 
    dMeetDate = "1/1/1900"
    sSite = vbNullString

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT m.MeetName, m.MeetDate, m.MeetSite FROM Meets m INNER JOIN IndRslts ir ON m.MeetsID = ir.MeetsID WHERE ir.RacesID = " & lThisRace 
   	rs2.Open sql2, conn2, 1, 2
    If rs2.RecordCount > 0 Then
        sMeetName = Replace(rs2(0).Value, "''", "'")
        dMeetDate = rs2(1).Value
        sSite = Left(rs2(2).Value, 15)
    End If
	rs2.Close
	Set rs2 = Nothing
End Sub

Private Sub MeetName(lThisMeet)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT m.MeetName, m.MeetDate FROM Meets m INNER JOIN IndRslts ir ON m.MeetsID = ir.MeetsID WHERE ir.RacesID = " & lThisMeet 
   	rs2.Open sql2, conn2, 1, 2
    MeetName = Replace(rs2(0).Value, "''", "'") & " (" & rs2(1).Value & ")"
	rs2.Close
	Set rs2 = Nothing
End Sub

Function GetPlace(lThisRaceID, lThisRstrID)
	GetPlace = 0
	sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRaceID & " AND Place > 0 ORDER BY RaceTime"
	Set rs2 = conn2.Execute(sql2)
	Do While Not rs2.EOF
		GetPlace = GetPlace + 1
		If CLng(rs2(0).Value) = CLng(lThisRstrID) Then Exit Do
		rs2.MoveNext
	Loop
	Set rs2 = Nothing
End Function

Private Function ConvertToSeconds(sTime)
    Dim sSubStr(3), Count, j
    Dim sglSeconds(3), k

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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Print GSE Performance Tracker Lists</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td,th{
        padding-right: 5px;
    }
</style>
</head>

<body>
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->
    <!--#include file = "perf_trkr_nav.asp" -->

    <div class="row">
        <h4 class="h4">Print My Performance Tracker Lists</h4>

        <div style="margin: 0;padding: 0;text-align: right;font-size: 0.8em;">
            <a href="javascript:window.print();">Send to Printer</a>
        </div>

        <h4 class="h4"><%=sPackName%></h4>
        <h5><%=sSport%>&nbsp;-&nbsp;<%=sGender%> <span style="font-weight: normal;">(created <%=dWhenCreated%>)</span></h5>

        <ul style="font-size: 0.8em;margin-top: 10px;">
            <li><span style="font-weight:bold;">Sort:</span>&nbsp;<%=sSortBy%></li>
            <li><span style="font-weight:bold;">Blend:</span>&nbsp;<%=sBlendBy%></li>
            <li><span style="font-weight:bold;">Dates:</span> From <%=dBegDate%> To <%=dEndDate%></li>
            <li>
                <span style="font-weight:bold;">Excluded Sites:</span>
                <%If sExclSites = vbNullString Then%>
                    None
                <%Else%>
                    <%=sExclSites%>
                <%End If%>
            </li>
            <li>
                <span style="font-weight:bold;">Excluded Meets:</span>
                <%If UBound(ExclMeets) = 0 Then%>
                    None
                <%Else%>
                   <ul>
                    <%For i = 0 To UBound(ExclMeets) - 1%>
                       <li><%=MeetName(ExclMeets(i))%></li>
                    <%Next%>
                   </ul>
                <%End If%>
            </li>
        </ul>

        <%Select Case sBlendBy%>
            <%Case "part"%>
                <ol style="font-size: 0.75em;margin-top: 0;padding-top: 0;">
                    <%For i = 0 To UBound(ViewPerf, 2) - 1%>
                        <%Call MyResults(ViewPerf(1, i))%>
                        <li style="font-weight:bold;margin-top: 10px;">
                            <%=ViewPerf(2, i)%> &nbsp; (<%=ViewPerf(3, i)%>)
                        
                            <%If UBound(RsltsArr, 2) = 0 Then%>
                                <p>No results available</p>
                            <%Else%>
                                <table style="font-size: 0.97em;">
                                    <tr>
                                        <th>No.</th>
                                        <th style="text-align:left;">Meet</th>
                                        <th style="text-align:left;">Date</th>
                                        <th style="text-align:left;">Site</th>
                                        <th style="text-align:left;">Race</th>
                                        <th style="text-align:left;">Dist</th>
                                        <th>Pl</th>
                                        <th>Time</th>
                                    </tr>
                                    <%For j = 0 To UBound(RsltsArr, 2) - 1%>
                                        <%If j mod 2 = 0 Then%>
                                            <tr>
                                                <td style="text-align:right;"><%=j + 1%>)</td>
                                                <%For k = 0 To 6%>
                                                    <td><%=RsltsArr(k, j)%></td>
                                                <%Next%>
                                            </tr>
                                        <%Else%>
                                            <tr>
                                                <td class="alt" style="text-align:right;"><%=j + 1%>)</td>
                                                <%For k = 0 To 6%>
                                                    <td class="alt"><%=RsltsArr(k, j)%></td>
                                                <%Next%>
                                            </tr>
                                        <%End If%>
                                    <%Next%>
                                </table>
                            <%End If%>
                        </li>
                    <%Next%>
                </ol>
            <%Case "meet"%>
                <ol style="font-size: 0.75em;">
                    <%For i = 0 To UBound(Meets, 2) - 1%>
                        <%Call GetMeetRslts(Meets(0, i))%>
                        <%If sCompOnly = "y" Then%>
                            <%If UBound(MeetRslts, 2) > 1 Then%>
                                <li style="font-weight:bold;margin-top: 10px;">
                                    <%=Meets(1, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>&nbsp; - &nbsp;<%=Meets(3, i)%>)</span>
                        
                                    <table style="font-size: 0.97em;">
                                        <tr>
                                            <th>No.</th>
                                            <th style="text-align:left;">Name</th>
                                            <th style="text-align:left;">Team</th>
                                            <th style="text-align:left;">Race</th>
                                            <th style="text-align:left;">Dist</th>
                                            <th>Pl</th>
                                            <th>Time</th>
                                        </tr>
                                        <%For j = 0 To UBound(MeetRslts, 2) - 1%>
                                            <%If j mod 2 = 0 Then%>
                                                <tr>
                                                    <td style="text-align:right;"><%=j + 1%>)</td>
                                                    <%For k = 0 To 5%>
                                                        <td><%=MeetRslts(k, j)%></td>
                                                    <%Next%>
                                                </tr>
                                            <%Else%>
                                                <tr>
                                                    <td class="alt" style="text-align:right;"><%=j + 1%>)</td>
                                                    <%For k = 0 To 5%>
                                                        <td class="alt"><%=MeetRslts(k, j)%></td>
                                                    <%Next%>
                                                </tr>
                                            <%End If%>
                                        <%Next%>
                                    </table>
                                </li>
                            <%End If%>
                        <%Else%>
                            <li style="font-weight:bold;margin-top: 10px;">
                                <%=Meets(1, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>&nbsp; - &nbsp;<%=Meets(3, i)%></span>)
                        
                                <table style="font-size: 0.97em;">
                                    <tr>
                                        <th>No.</th>
                                        <th style="text-align:left;">Name</th>
                                        <th style="text-align:left;">Team</th>
                                        <th style="text-align:left;">Race</th>
                                        <th style="text-align:left;">Dist</th>
                                        <th>Pl</th>
                                        <th>Time</th>
                                    </tr>
                                    <%For j = 0 To UBound(MeetRslts, 2) - 1%>
                                        <%If j mod 2 = 0 Then%>
                                            <tr>
                                                <td style="text-align:right;"><%=j + 1%>)</td>
                                                <%For k = 0 To 5%>
                                                    <td><%=MeetRslts(k, j)%></td>
                                                <%Next%>
                                            </tr>
                                        <%Else%>
                                            <tr>
                                                <td class="alt" style="text-align:right;"><%=j + 1%>)</td>
                                                <%For k = 0 To 5%>
                                                    <td class="alt"><%=MeetRslts(k, j)%></td>
                                                <%Next%>
                                            </tr>
                                        <%End If%>
                                    <%Next%>
                                </table>
                            </li>
                        <%End If%>
                    <%Next%>
                </ol>
            <%Case "site"%>
                    <ol style="font-size: 0.75em;">
                    <%For i = 0 To UBound(Meets, 2) - 1%>
                        <%bGetRslts = False%>
                        <%If i = 0 Then%>
                            <%sThisSite = Meets(3, i)%>
                            <%bGetRslts = True%>
                        <%ElseIf Not CStr(sThisSite) = CStr(Meets(3, i)) Then%>
                            <%sThisSite = Meets(3, i)%>
                            <%bGetRslts = True%>
                        <%End If%> 
                         
                        <%If bGetRslts = True Then%>   
                            <%Call GetSiteRslts(sThisSite)%>
                            <%If sCompOnly = "y" Then%>
                                <%If UBound(SiteRslts, 2) > 1 Then%>
                                    <li style="font-weight:bold;margin-top: 10px;">
                                        <%=Meets(3, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>)</span>
                        
                                        <table style="font-size: 0.97em;">
                                            <tr>
                                                <th>No.</th>
                                                <th style="text-align:left;">Name</th>
                                                <th style="text-align:left;">Team</th>
                                                <th style="text-align:left;">Meet</th>
                                                <th style="text-align:left;">Race</th>
                                                <th style="text-align:left;">Dist</th>
                                                <th>Pl</th>
                                                <th>Time</th>
                                            </tr>
                                            <%For j = 0 To UBound(SiteRslts, 2) - 1%>
                                                <%If j mod 2 = 0 Then%>
                                                    <tr>
                                                        <td style="text-align:right;"><%=j + 1%>)</td>
                                                        <%For k = 0 To 6%>
                                                            <td><%=SiteRslts(k, j)%></td>
                                                        <%Next%>
                                                    </tr>
                                                <%Else%>
                                                    <tr>
                                                        <td class="alt" style="text-align:right;"><%=j + 1%>)</td>
                                                        <%For k = 0 To 6%>
                                                            <td class="alt"><%=SiteRslts(k, j)%></td>
                                                        <%Next%>
                                                    </tr>
                                                <%End If%>
                                            <%Next%>
                                        </table>
                                    </li>
                                <%End If%>
                            <%Else%>
                                <li style="font-weight:bold;margin-top: 10px;">
                                    <%=Meets(3, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>)</span>
                        
                                    <table style="font-size: 0.97em;">
                                        <tr>
                                            <th>No.</th>
                                            <th style="text-align:left;">Name</th>
                                            <th style="text-align:left;">Team</th>
                                            <th style="text-align:left;">Meet</th>
                                            <th style="text-align:left;">Race</th>
                                            <th style="text-align:left;">Dist</th>
                                            <th>Pl</th>
                                            <th>Time</th>
                                        </tr>
                                        <%For j = 0 To UBound(SiteRslts, 2) - 1%>
                                            <%If j mod 2 = 0 Then%>
                                                <tr>
                                                    <td style="text-align:right;"><%=j + 1%>)</td>
                                                    <%For k = 0 To 6%>
                                                        <td><%=SiteRslts(k, j)%></td>
                                                    <%Next%>
                                                </tr>
                                            <%Else%>
                                                <tr>
                                                    <td class="alt" style="text-align:right;"><%=j + 1%>)</td>
                                                    <%For k = 0 To 6%>
                                                        <td class="alt"><%=SiteRslts(k, j)%></td>
                                                    <%Next%>
                                                </tr>
                                            <%End If%>
                                        <%Next%>
                                    </table>
                                </li>
                            <%End If%>
                        <%End If%>
                    <%Next%>
                </ol>
        <%End Select%>
    </div>
</div>
</body>
<%
conn.close
Set conn = Nothing

conn2.close
Set conn2 = Nothing
%>
</html>
