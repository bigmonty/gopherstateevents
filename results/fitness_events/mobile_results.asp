<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, x, m, n, k
Dim lRaceID, lEventID, lSuppLegID, lMyPartID, lMyRaceID, lFeaturedEventsID
Dim iRaceType, iTtlRcds, iEventType,iBibToFind, iMinPlace, iMaxPlace, iLinkNum, iNumAgeGrps
Dim sEventName, sGender, sSortRsltsBy, sDist, sRaceName, sMF, sTimingMethod, sChipStart, sAllowDuplAwds, sEventRaces, sOrderBy, sErrMsg
Dim sGalleryLink, sTypeFilter, sSuppTime, sOtherTime, sLegName, sOtherName, sHasSplits, sLogo, sShowAge, sIndivRelay, sBannerImage, sClickPage
Dim sngMyTime
Dim dEventDate
Dim BibRslts(6), Events, Races, IndRslts, TempArr(9)
Dim bRsltsOfficial

'Response.Redirect "/misc/taking_break.htm"

sClickPage = Request.ServerVariables("URL")

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lSuppLegID = 0

iEventType = Request.QueryString("event_type")
If CStr(iEventType) = vbNullString Then iEventType = 5
If Not IsNumeric(iEventType) Then Response.Redirect("http://www.google.com")
If CLng(iEventType) < 0 Then Response.Redirect("http://www.google.com")

iLinkNum = 0

Select Case CInt(iEventType)
    Case 46
        sTypeFilter = "' AND EventType IN(4, 6)"
    Case 910
        sTypeFilter = "' AND EventType IN(9, 10)"
    Case Else
        sTypeFilter = "' AND EventType = " & iEventType
End Select

lRaceID = Request.QueryString("race_id")

Response.Write lRaceID

iBibToFind = 0
iTtlRcds = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FeaturedEventsID, BannerImage, Views FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date 
sql = sql & "' AND '" & Date + 360 & "') AND Active = 'y' ORDER BY NewID()"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    lFeaturedEventsID = rs(0).Value
    sBannerImage = rs(1).Value
    rs(2).Value = CLng(rs(2).Value) + 1
    rs.Update
End If
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")

    If CStr(lEventID) = vbNullString Then lEventID = 0
    If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
    If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")
    If CStr(sGender) = vbNullString Then sGender = "M"

    lRaceID = GetFirstRace()
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
	sGender = Request.Form.Item("gender")
ElseIf Request.form.Item("submit_bib") = "submit_bib" Then
    iBibToFind = Request.Form.Item("bib_to_find")
End If

If sGender = vbNullString Then sGender = "M"

'log this user if they are just entering the site
If Session("access_results") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'fitness_results')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate <= '" & Date & sTypeFilter & " ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

If CLng(lEventID) = 0 Then
    ReDim IndRslts(9, 0)
Else
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sEventRaces = sEventRaces & rs(0).Value & ", "
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

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

    If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ")"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iTtlRcds = rs.RecordCount
    rs.Close
    Set rs = Nothing

	'get event information
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventName, EventDate, TimingMethod, Logo FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
    sTimingMethod = rs(2).Value
    sLogo = rs(3).Value
	rs.Close
	Set rs = Nothing

    'get races
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	Races = rs.GetRows()
	rs.Close
	Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventID FROM OfficialRslts WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then bRsltsOfficial = True
	rs.Close
	Set rs = Nothing
	
    If CLng(lRaceID) = 0 Then lRaceID = GetFirstRace()

    'check for team results
    Dim sHasTeams
    sHasTeams = "n"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID FROM Teams WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sHasTeams = "y"
    rs.Close
    Set rs = Nothing

    iNumAgeGrps = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iNumAgeGrps = rs.RecordCount
    rs.Close
    Set rs = Nothing

    sHasSplits = "n"
    sIndivRelay = "indiv"
	sql = "SELECT Dist, RaceName, Type, AllowDuplAwds, ChipStart, SortRsltsBy, NumSplits, ShowAge, IndivRelay FROM RaceData WHERE RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	sDist = rs(0).Value
	sRaceName = rs(1).Value
	iRaceType = rs(2).Value
	sAllowDuplAwds = rs(3).Value
    sChipStart = rs(4).Value
    sSortRsltsBy = rs(5).Value
    If CInt(rs(6).Value) > 0 Then sHasSplits = "y"
    sShowAge = rs(7).Value
    sIndivRelay = rs(8).Value
	Set rs = Nothing

    sSortRsltsBy = "FnlTime"
    If sSortRsltsBy = "FnlTime" Then
        sOrderBy = "ir.FnlScnds"
    Else
        sOrderBy = "ir.EventPl"
    End If

    If sTimingMethod = "RFID" AND sChipStart = "y" Then
        If sGender = "B" Then
            sql = "SELECT pr.Bib, p.LastName, p.FirstName, p.Gender, pr.Age, ir.ChipTime, ir.FnlTime, ir.ChipStart, p.City, p.St "
            sql = sql & "FROM Participant p JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
            sql = sql & "JOIN PartRace pr ON pr.RaceID = ir.RaceID AND pr.ParticipantID = p.ParticipantID "
            sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND ir.FnlTime IS NOT NULL AND ir.FnlTime > '00:00:00.000' ORDER BY " & sOrderBy
        Else
            sql = "SELECT pr.Bib, p.LastName, p.FirstName, p.Gender, pr.Age, ir.ChipTime, ir.FnlTime, ir.ChipStart, p.City, p.St "
            sql = sql & "FROM Participant p JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
            sql = sql & "JOIN PartRace pr ON pr.RaceID = ir.RaceID AND pr.ParticipantID = p.ParticipantID "
            sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND ir.FnlTime IS NOT NULL AND ir.FnlTime > '00:00:00.000' AND p.Gender = '" & sGender 
            sql = sql & "' ORDER BY " & sOrderBy
        End If
    Else
        If sGender = "B" Then
            sql = "SELECT pr.Bib, p.LastName, p.FirstName, p.Gender, pr.Age, ir.FnlTime, ir.ChipTime, ir.ChipStart, p.City, p.St "
            sql = sql & "FROM Participant p JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
            sql = sql & "JOIN PartRace pr ON pr.RaceID = ir.RaceID AND pr.ParticipantID = p.ParticipantID "
            sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND ir.FnlTime IS NOT NULL AND ir.FnlTime > '00:00:00.000' ORDER BY " & sOrderBy
        Else
            sql = "SELECT pr.Bib, p.LastName, p.FirstName, p.Gender, pr.Age, ir.FnlTime, ir.ChipTime, ir.ChipStart, p.City, p.St "
            sql = sql & "FROM Participant p JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
            sql = sql & "JOIN PartRace pr ON pr.RaceID = ir.RaceID AND pr.ParticipantID = p.ParticipantID "
            sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND ir.FnlTime IS NOT NULL AND ir.FnlTime > '00:00:00.000' AND p.Gender = '" 
            sql = sql & sGender & "' ORDER BY " & sOrderBy
        End If
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        IndRslts = rs.GetRows()
    Else
        ReDim IndRslts(9, 0)
    End If
    rs.Close
    Set rs = Nothing

    If sChipStart = "n" Or sTimingMethod = "Conv" Then
		For i = 0 To UBound(IndRslts, 2)
		    If IndRslts(5, i) & "" <> "" Then
                sngMyTime = ConvertToSeconds(IndRslts(5, i))
				IndRslts(6, i) = PacePerMile(sngMyTime, sDist)
				IndRslts(7, i) = PacePerKM(sngMyTime, sDist)
		    End If
        Next
	End If

	For i = 0 To UBound(IndRslts, 2)
        If sShowAge = "n" Then
            IndRslts(4, i) = MyAgeGrp(IndRslts(0, i))
        Else
		    If IndRslts(4, i) = "99" Then IndRslts(4, i) = "0"
        End If
    Next

    Private Function MyAgeGrp(iMyBib)
        If iMyBib & "" = "" Then
            MyAgeGrp = "n/a"
        Else
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT AgeGrp FROM PartRace WHERE Bib = " & iMyBib & " AND RaceID IN (" & sEventRaces & ")"
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then
                If Left(rs(0).Value, 3) = "110" Then
                    MyAgeGrp = "n/a"
                Else
                    MyAgeGrp = rs(0).Value
                End If
            Else
                MyAgeGrp = "n/a"
            End If
            rs.Close
            Set rs = Nothing
        End If
    End Function
    	
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT GalleryLink FROM RaceGallery WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sGalleryLink = rs(0).Value
    rs.Close
    Set rs = Nothing
End If

If Not CInt(iBibToFind) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, RaceID, Age FROM PartRace WHERE Bib = " & iBibToFind & " AND RaceID IN (" & sEventRaces & ")"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        lMyPartID = rs(0).Value
        lMyRaceID = rs(1).Value
        BibRslts(2) = rs(2).Value
    Else
        sErrMsg = "I'm sorry.  That bib number was not found."
    End If
    rs.Close
    Set rs = Nothing

    If sErrMsg = vbNullString Then
	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT LastName, FirstName, Gender FROM Participant WHERE ParticipantID = " & lMyPartID
        rs.Open sql, conn, 1, 2
        BibRslts(1) = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
        rs.Close
        Set rs = Nothing

        k = 1
	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ParticipantID, EventPl, FnlScnds, FnlTime FROM IndResults WHERE RaceID = " & lMyRaceID & " AND FnlScnds > 0 ORDER BY FnlScnds"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If CLng(rs(0).Value) = CLng(lMyPartID) Then
                BibRslts(0) = k
                Exit Do
            Else
                k = k + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lMyRaceID
        rs.Open sql, conn, 1, 2
        BibRslts(3) = Replace(rs(0).Value, "''", "'") 
        rs.Close
        Set rs = Nothing

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ChipTime, FnlTime, ChipStart FROM IndResults WHERE RaceID = " & lMyRaceID & " AND ParticipantID = " & lMyPartID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            BibRslts(4) = rs(0).Value
            BibRslts(5) = rs(1).Value
            BibRslts(6) = rs(2).Value
        Else
            sErrMsg = "I'm sorry.  That bib number was not found."
        End If
        rs.Close
        Set rs = Nothing
    End If

    If CLng(lSuppLegID) > 0 Then
        sSuppTime = "00:00.000"
        sOtherTime = "00:00:00.000"

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT SuppTime, OtherTime FROM SuppLegRslts WHERE SuppLegID = " & lSuppLegID & " AND Bib = " & iBibToFind
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            sSuppTime = rs(0).Value
            sOtherTime = ConvertToMinutes(ConvertToSeconds(BibRslts(4)) - ConvertToSeconds(rs(0).Value))
        End If
        rs.Close
        Set rs = Nothing
    End If

    If CInt(iMinPlace) <= 0 Then 
        iMinPlace = 1
        iMaxPlace = 20
    End If
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

Private Function EventName()
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventName FROM Events WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    EventName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

Private Function PacePerMile(sTime, sThisDist)
    Dim sglDist

    If UCase(Right(sThisDist, 2)) = "MI" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3))
    ElseIf UCase(Right(sThisDist, 4)) = "MILE" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 5))
    ElseIf UCase(sThisDist) = "MARATHON" Then
        sglDist = 26.2
    ElseIf UCase(sThisDist) = "H. MAR" Then
        sglDist = 13.1
    Else
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3)) * 0.6213712
    End If

    'calculate the pace
    PacePerMile = ConvertToMinutes(CSng(ConvertToSeconds(Round(sTime, 2))) / Round(sglDist, 2))
    PacePerMile = Replace(PacePerMile, "-", "")
End Function

Private Function PacePerKM(sTime, sThisDist)
    Dim sglDist

    If UCase(Right(sThisDist, 2)) = "MI" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3)) * 1.609344
    ElseIf UCase(Right(sThisDist, 4)) = "MILE" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 5)) * 1.609344
    ElseIf UCase(sThisDist) = "MARATHON" Then
        sglDist = CSng(26.2) * 1.609344
    ElseIf UCase(sThisDist) = "H. MAR" Then
        sglDist = CSng(13.1) * 1.609344
    Else
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3))
    End If
    
    'calculate the pace
    PacePerKM = ConvertToMinutes(CSng(ConvertToSeconds(Round(sTime, 2))) / Round(sglDist, 2))
    PacePerKM = Replace(PacePerKM, "-", "")
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

Function CleanInput(sInput)
	Dim char, x, bHackFound, sOrigInput
	
	sOrigInput = sInput
	bHackFound = False

	If InStr(UCase(sInput), "DECLARE") > 0 Then bHackFound = True

	If bHackFound = False Then
		If InStr(UCase(sInput), "IFRAME") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "STYLE=") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "HEIGHT=") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(sInput, ":8080") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "/TITLE") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), ".RU") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), ".JS") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "WEBSERVICE") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "SCRIPT SRC=") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "INSERT INTO") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "DELETE FROM") > 0 Then bHackFound = True
	End If
 
	If bHackFound = False Then
		If InStr(UCase(sInput), "NULL") > 0 Then bHackFound = True
	End If
 
	If bHackFound = False Then
		If InStr(UCase(sInput), ".EXE") > 0 Then bHackFound = True
	End If
	
	If bHackFound = True Then
		sHackMsg = "Some of your input appears to be an attempt to compromise the security of this website.  It has been sent to our security department for "
		sHackMsg = sHackMsg & "review.  If this was a genuine attempt to communicate please contact bob.schneider@gopherstateevents.com"
	Else
	    sInput = Replace(lcase(sInput), "http://", "")
	    sInput = Replace(lcase(sInput), "drop", "drp")
	    sInput = Replace(lcase(sInput), "js", "")
	    sInput = Replace(lcase(sInput), "xp_", "")
	    sInput = Replace(lcase(sInput), "CRLF", "")
	    sInput = Replace(lcase(sInput), "%3A", "")';
	    sInput = Replace(lcase(sInput), "%3B", "")':
	    sInput = Replace(lcase(sInput), "%3D", "equals")
	    sInput = Replace(lcase(sInput), "%3E", "grtr than")
	    sInput = Replace(lcase(sInput), "%3F", "")'?
	    sInput = Replace(lcase(sInput), "&quot;", "")
	    sInput = replace(lcase(sInput), "&amp;", "and")
	    sInput = replace(lcase(sInput), "&lt;", "lss than")
	    sInput = replace(lcase(sInput), "&gt;", "grtr than")
	    sInput = replace(lcase(sInput), " exec ", "")
	    sInput = replace(lcase(sInput), "onvarchar", "")
	    sInput = replace(lcase(sInput), "set", "")
	    sInput = replace(lcase(sInput), " cast ", "")
	    sInput = replace(lcase(sInput), "00100111", "")
	    sInput = replace(lcase(sInput), "00100010", "")
	    sInput = replace(lcase(sInput), "00111100", "")
	    sInput = replace(lcase(sInput), "select", "selct")
	    sInput = replace(lcase(sInput), "0x", "")
	    sInput = replace(lcase(sInput), "delete", "delet")
	    sInput = replace(lcase(sInput), "go ", "")
	    sInput = replace(lcase(sInput), "create", "creat")
	    sInput = replace(lcase(sInput), "convert", "cnvrt")
	    sInput = replace(lcase(sInput), "=", "equals")
	    sInput = replace(lcase(sInput), "/", "")
	    sInput = replace(lcase(sInput), "\", "")
	    sInput = replace(lcase(sInput), "?", "")
	    sInput = replace(lcase(sInput), "# ", " ")
	    sInput = replace(lcase(sInput), ";", "")
	    sInput = replace(lcase(sInput), ":", "")
	    sInput = replace(lcase(sInput), "$", "")
	    sInput = replace(lcase(sInput), "<", "lss than")
	    sInput = replace(lcase(sInput), ">", "grtr than")
	    sInput = replace(lcase(sInput), "(", "-")
	    sInput = replace(lcase(sInput), ")", "-")
	    sInput = replace(lcase(sInput), "+ ", "plus")
	    sInput = replace(lcase(sInput), "~", "")
	    sInput = replace(lcase(sInput), "|", "")
	    sInput = replace(lcase(sInput), "$", "")

		If LCase(sOrigInput) = sInput Then
	    	CleanInput = sOrigInput
		Else
			CleanInput = sInput
		End If
	End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events Mobile Results</title>
<meta name="description" content="Fitness Event Results from Gopher State Events, LLC">
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
        <a href="http://www.gopherstateevents.com/featured_events/featured_clicks.asp?featured_events_id=<%=lFeaturedEventsID%>&amp;click_page=<%=sClickPage%>" 
            onclick="openThis(this.href,1024,768);return false;">
            <img src="http://www.gopherstateevents.com/featured_events/images/<%=sBannerImage%>" alt="<%=sBannerImage%>" class="img-responsive">
        </a>

        <img src="/graphics/html_header.png" class="img-responsive" alt="Individual Results">
	    <%If Not CLng(lEventID) = 0 Then%>
            <h4 class="h4"> <%=sEventName%> - <%=sRaceName%> (<%=sDist%>)<br><small> Results</small></h4>
        <%End If%>

        <a href="http://www.gopherstateevents.com" style="font-weight: bold;">Return To Main Site</a>

	    <form role="form" class="form-inline" name="which_event" method="post" action="mobile_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>">
	    <div class="form_group">
            <label>Event:</label>
		    <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
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

    <%If Not CLng(lEventID) = 0 Then%>
        <p>Coming Soon!  The ability to comment on results, individually and overall.  Stay tuned!</p>

        <div class="row">
            <div class="col-sm-9">
		        <form role="form" class="form" name="get_races" method="post" action="mobile_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>">
                <div class="form_group">
			        <label for="races">Race:</label>
			        <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
				        <%For i = 0 to UBound(Races, 2)%>
					        <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
						        <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
					        <%Else%>
						        <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
					        <%End If%>
				        <%Next%>
			        </select>
			        <label>Gender:</label>
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
                                <option value="B">Combined</option>
						        <option value="M">Male</option>
						        <option value="F">Female</option>
				        <%End Select%>
			        </select>
			        <input class="form-control" type="hidden" name="submit_race" id="submit_race" value="submit_race">
			        <input class="form-control" type="submit" name="get_race" id="get_race" value="View">
                </div>
		        </form>
            </div>
            <div class="col-sm-3">
                <%If Not sLogo & "" = "" Then%>
                    <img src="/events/logos/<%=sLogo%>" class="img-responsive"alt="<%=sEventName%>">
                <%End If%>
            </div>
        </div>

	    <%If Not CLng(lRaceID) = 0 Then%>
            <p><span style="font-weight:bold;">Total Finishers:</span>&nbsp;<%=iTtlRcds%></p>

   		    <%If CDate(dEventDate) > Date Then%>
			    <div class="bg-success">This event is currently scheduled for <%=dEventDate%>.  The results will be available on that date.</div>
		    <%Else%>
                <%If CDate(Date) < CDate(dEventDate) + 7 Then%>
			        <%If bRsltsOfficial = False Then%>
				        <div class="bg-warning">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE.  
                            Please report any issues to bob.schneider@gopherstateevents.com.
                        </div>
			        <%Else%>
				        <div class="bg-warning">
                            These results are now official.  If you notice any errors please contact us via 
                            <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by telephone (612.720.8427).
                        </div>
			        <%End If%>
                <%End If%>
		    <%End If%>
	
            <%If sTimingMethod = "RFID" And sChipStart = "y" Then%>
                <div class="bg-info">
                    Note:  This race used a chip start which takes into consideration when you crossed the starting line as well as 
                    when you crossed the finish line.  As a result, people that were close to you at the finish line may not appear that way in the 
                    results.
                </div>
            <%End If%>

			<%If sHasSplits = "y" And sGender <> "B" Then%>
    			<div style="text-align:right;font-size:0.9em;margin:10px 10px 0 0;background-color:#ececd8;">
                    <a href="splits/results_w-splits.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                    onclick="openThis(this.href,1024,768);return false;" style="color: red;">Results With Splits</a>
                    &nbsp;|&nbsp;
                    <a href="splits/rank_by_split.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                    onclick="openThis(this.href,1024,768);return false;" style="color: red;">Rank By Split</a>
                    &nbsp;|&nbsp;
                    <a href="blended_results.asp?event_id=<%=lEventID%>" 
                    onclick="openThis(this.href,1024,768);return false;" style="color: green;">Blended Results</a>
                </div>
            <%End If%>

            <%If Not (CLng(lEventID) = 380 Or CLng(lEventID) = 444) Then%>
                <div class="table bg-warning">
                    <label>Bib To Find:</label>
                    <br>
                    <form name="find_bib" method="post" action="mobile_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                        onsubmit="return chkFlds2();">
                    <div class="form_group">
                        <input class="form-control" type="text" name="bib_to_find" id="bib_to_find" size="3" maxlength="4" value ="<%=iBibToFind%>">
                        <input class="form-control" type="hidden" name="submit_bib" id="submit_bib" value="submit_bib">
                        <input class="form-control" type="submit" name="submit_lookup" id="submit_lookup" value="Find Bib">
                    </div>
                    </form>

                    <%If Not CInt(iBibToFind) = 0 Then%>
                        <%If sErrMsg = vbNullString Then%>
                            <table class="table-condensed">
                                <tr>
                                    <th>Pl</th>
                                    <th>Name</th>
                                    <%If sShowAge = "y" Then%>
                                        <th>Age</th>
                                    <%Else%>
                                        <th>Age Grp</th>
                                    <%End If%>
                                    <th>Race</th>
                                    <th>Chip</th>
                                    <th>Gun</th>
                                    <th>Start</th>
                                </tr>
                                <tr>
                                    <%For i = 0 To 6%>
                                        <td style="text-align:center;"><%=BibRslts(i)%></td>
                                    <%Next%>
                                </tr>
                                <%If CLng(lSuppLegID) > 0 Then%>
                                    <tr>
                                        <th style="text-align: right;" colspan="3"><%=sLegName%></th>
                                        <td><%=sSuppTime%></td>
                                        <th style="text-align: right;" colspan="3"><%=sOtherName%></th>
                                        <td><%=sOtherTime%></td>
                                    </tr>
                                <%End If%>
                            </table>
                        <%End If%>
                    <%End If%>
                </div>
            <%End If%>

            <%If sIndivRelay = "relay" Then%>
                <div class="bg-success">
                    <a href="javascript:pop('relay_by_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)">Results by Split</a>
                    &nbsp;|&nbsp;
                    <a href="javascript:pop('relay_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)" >Results w/Splits</a>
                </div>
            <%End If%>

		    <div>
                <%If Not sGalleryLink = vbNullString Then%>
                    <a href="<%=sGalleryLink%>">Race Pix</a>
                    <%iLinkNum = CInt(iLinkNum) + 1%>
                <%End If%>
                <%If CInt(iEventType) = 5 Then%>
                    <%iLinkNum = CInt(iLinkNum) + 1%>
                    <%If sShowAge = "y" Then%>
                        <%If sShowAge = "n" Then%>
                            <%If CInt(iLinkNum) > 1 Then%>
                                &nbsp;|&nbsp;
                            <%End If%>
				            <a href="age_graded.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                onclick="openThis(this.href,1024,768);return false;" style="font-weight:normal">Age-Graded</a>
                        <%End If%>
                    <%End If%>
                <%End If%>
                <%If CInt(iEventType) = 910 Then%>
                    &nbsp;|&nbsp;
				    <a href="trans_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                    onclick="openThis(this.href,1024,768);return false;" style="font-weight:normal">Trans Data</a>
                    &nbsp;|&nbsp;
				    <a href="multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                    onclick="openThis(this.href,1024,768);return false;" style="font-weight:normal">Results w/Splits</a>
                <%End If%>
			    <%If Not sGender = "B" Then%>
                    <%iLinkNum = CInt(iLinkNum) + 1%>
                    <%If CInt(iNumAgeGrps) > 1 Then%>
                        <%If CInt(iLinkNum) > 1 Then%>
                            &nbsp;|&nbsp;
                        <%End If%>
                        <a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                           onclick="openThis(this.href,1024,768);return false;">Awards</a>
                        <%iLinkNum = CInt(iLinkNum) + 1%>
                        <%If CInt(iLinkNum) > 1 Then%>
                            &nbsp;|&nbsp;
                        <%End If%>
                        <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                           onclick="openThis(this.href,1024,768);return false;" >Age Groups</a>
                    <%End If%>
                <%End If%>
				<%If sHasTeams = "y" Then%>
                    &nbsp;|&nbsp;
                    <a href="team_results.asp?race_id=<%=lRaceID%>" 
                        onclick="openThis(this.href,1024,768);return false;">Team Results</a>
                <%End If%>
				&nbsp;|&nbsp;
                <a href="/records/records.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                onclick="openThis(this.href,1024,768);return false;" style="font-weight:normal">Records</a>
            </div>

            <%If CLng(lEventID) = 380 Or CLng(lEventID) = 444 Then%>
                <div style="background-color:#ececdf;margin:10px 0 0 0;padding:5px;">
                    <p style="color: red;">These results are available through Minnetonka Community Education.</p>
                </div>
            <%Else%>
		        <table class="table-striped">
			        <tr>
                        <%If sTimingMethod = "RFID" And sChipStart = "y" Then%>
				            <th style="width:10px">Pl</th>
				            <th style="text-align:left;">Bib-Name</th>
				            <th>M/F</th>
                            <%If sShowAge = "y" Then%>
                                <th>Age</th>
                            <%Else%>
                                <th>Age Grp</th>
                            <%End If%>
				            <th>Chip</th>
				            <th>Gun</th>
				            <th>Start</th>
				            <th style="text-align:left;">From</th>
                        <%Else%>
				            <th style="width:10px">Pl</th>
				            <th style="text-align:left;">Bib-Name</th>
				            <th>M/F</th>
                            <%If sShowAge = "y" Then%>
                                <th>Age</th>
                            <%Else%>
                                <th>Age Grp</th>
                            <%End If%>
				            <th>Time</th>
				            <th>Per Mi</th>
				            <th>Per Km</th>
				            <th style="text-align:left;">From</th>
                            <th>&nbsp;</th>
                        <%End If%>
			        </tr>
			        <%For i = 0 To UBound(IndRslts, 2)%>
					    <tr>
						    <td style="width:10px;"><%=i + 1%></td>
						    <td>
                                <a href="javascript:pop('ind_rslts.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>&amp;bib=<%=IndRslts(0, i)%>',600,700)">
                                    <%=IndRslts(0, i)%> - <%=IndRslts(2, i)%>&nbsp;<%=IndRslts(1, i)%>
                                </a>
                            </td>
						    <td style="text-align:center;"><%=IndRslts(3, i)%></td>
						    <%If CLng(lRaceID) = 350 Then%>
			                    <td>n/a</td>
		                    <%Else%>
			                    <td><%=IndRslts(4, i)%></td>
		                    <%End If%>
						    <td><%=IndRslts(5, i)%></td>
						    <td><%=IndRslts(6, i)%></td>
						    <td><%=IndRslts(7, i)%></td>
						    <td><%=IndRslts(8, i)%>, <%=IndRslts(9, i)%></td>
					    </tr>
			        <%Next%>
		        </table>
            <%End If%>
	    <%End If%>
    <%End If%>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>