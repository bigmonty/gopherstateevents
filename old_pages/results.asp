<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, x, m, n,  k, j
Dim lRaceID, lEventID, lMyRaceID, lMyPartID, lSuppLegID, lFeaturedEventsID
Dim iRaceType, iBibToFind, iTtlRcds, iEventType, iMinPlace, iMaxPlace, iNumAgeGrps, iNumRace, iNumMAgeGrps, iNumFAgeGrps, iNumLaps
Dim iFPlace, iMPlace
Dim sEventName, sGender, sSortRsltsBy, sDist, sRaceName, sMF, sThisLeg, sThisSplit, sThisTrans, sAllowDuplAwds, sGallery, sChipStart, sLogo
Dim sWeather, sEventRaces, sErrMsg, sEventClass, sLocation, sTypeFilter, sHasSplits, sIndivRelay, sTimed, sSuppTime, sOtherTime, sLegName
Dim sOtherName, sShowAge, sRaceReport, sBannerImage, sClickPage
Dim sngMyTime
Dim dEventDate
Dim BibRslts(6), Events, Races, IndRslts, TempArr(9), RaceGallery(), CustomFields()
Dim bRsltsOfficial, bShowFeatured

'Response.Redirect "/misc/taking_break.htm"

sClickPage = Request.ServerVariables("URL")

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0

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

'mobile results code here

iBibToFind = 0
iTtlRcds = 0
iNumRace = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

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

If Request.Form.Item("submit_order") = "submit_order" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'fitness_results' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("fitness_results") = "y"
	rs.Close
	Set rs = Nothing

    'send email
	If Session("fitness_results") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
        Dim iBibNum
        Dim sEmail, sMsg
        Dim cdoMessage, cdoConfig

        iBibNum = CleanInput(Trim(Request.Form.Item("bib_num")))
        If sHackMsg = vbNullString Then sEmail = CleanInput(Trim(Request.Form.Item("email")))

        If sHackMsg = vbNullString Then
            'write to table
            sql = "INSERT INTO MediaOrder(BibNum, Email, WhenOrdered, IPAddress, EventID, MediaType) VALUES (" & iBibNum & ", '" 
            sql = sql & sEmail & "', '" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', " & lEventID & ", 'both')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing

			sMsg = "Thank you for ordering finish line media from Gopher State Events.  We have already begun processing your order.  The details of "
            sMsg = sMsg & "your order can be found below. Please verify that they are correct:" & vbCrLf & vbCrLf
			
			sMsg = sMsg & "Event Name: " & EventName() & vbCrLf
			sMsg = sMsg & "Bib Number: " & iBibNum & vbCrLf & vbCrLf

            sMsg = sMsg & "You will receive a link for online payment shortly.  Once payment is received your order will be completed and sent to you "
            sMsg = sMsg & "via the email address that you have supplied. " & vbCrLF & vbCrLf

            sMsg = sMsg & "Sincerely; " & vbCrLf
            sMsg = sMsg & "Bob Schneider " & vbCrLf
            sMsg = sMsg & "Gopher State Events, LLC " & vbCrLf
            sMsg = sMsg & "612.720.8427 " & vbCrLf

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = "bob.schneider@gopherstateevents.com;bob.bakken@gopherstateevents.com;" & sEmail
				.From = sEmail
				.Subject = "GSE Media Order"
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
		End If

	    sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'fitness_results'"
	    Set rs = conn.Execute(sql)
	    Set rs = Nothing

	    Session.Contents.Remove("access_results")
	End If
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")

    If CStr(lEventID) = vbNullString Then lEventID = 0
    If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
    If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")
    If CStr(sGender) = vbNullString Then sGender = "B"

    lRaceID = GetFirstRace()
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
	sGender = Request.Form.Item("gender")
ElseIf Request.form.Item("submit_bib") = "submit_bib" Then
    iBibToFind = Request.Form.Item("bib_to_find")
End If

If sGender = vbNullString Then sGender = "B"

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

If CLng(lEventID) = 0 Then
    ReDim IndRslts(17, 0)
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

'    If sChipStart = "y" Then
'        Set rs = Server.CreateObject("ADODB.Recordset")
'        sql = "SELECT ChipTime, FnlScnds FROM IndResults WHERE RaceID IN (" & sEventRaces & ")"
'        rs.Open sql, conn, 1, 2
'        Do While Not rs.EOF
'            rs(1).Value = ConvertToSeconds(rs(0).Value)
'            rs.Update
'            rs.MoveNext
'        Loop
'        rs.Close
'        Set rs = Nothing
'    End If

    If sGender = "B" Then
        If sSortRsltsBy = "place" Then
            Set rs = Server.CreateObject("ADODB.Recordset")  
            sql = "Exec ResultsByPlace @RaceID = " & lRaceID
            Set rs = conn.execute(sql) 
            If rs.BOF and rs.EOF Then
                ReDim IndRslts(17, 0)
            Else
                IndRslts = rs.GetRows()
            End If
            rs.Close
            Set rs = Nothing
        Else
            Set rs = Server.CreateObject("ADODB.Recordset")  
            sql = "Exec OverallResults @RaceID = " & lRaceID
            Set rs = conn.execute(sql) 
            If rs.BOF and rs.EOF Then
                ReDim IndRslts(17, 0)
            Else
                IndRslts = rs.GetRows()
            End If
            rs.Close
            Set rs = Nothing
        End If
    Else
        If sSortRsltsBy = "place" Then
            Set rs = Server.CreateObject("ADODB.Recordset")  
            sql = "Exec GenderByPlace @RaceID = " & lRaceID & ", @Gender = '" & sGender & "'"
            Set rs = conn.execute(sql) 
            If rs.BOF and rs.EOF Then
                ReDim IndRslts(17, 0)
            Else
                IndRslts = rs.GetRows()
            End If
            rs.Close
            Set rs = Nothing
        Else
            Set rs = Server.CreateObject("ADODB.Recordset")  
            sql = "Exec GenderResults @RaceID = " & lRaceID & ", @Gender = '" & sGender & "'"
            Set rs = conn.execute(sql) 
            If rs.BOF and rs.EOF Then
                ReDim IndRslts(17, 0)
            Else
                IndRslts = rs.GetRows()
            End If
            rs.Close
            Set rs = Nothing
        End If
    End If
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
        If rs.RecordCount > 0 Then BibRslts(1) = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
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

If CLng(lEventID) > 0 Then
    iFPlace = 0
    iMPlace = 0

    For i = 0 To UBound(IndRslts, 2) - 1
        IndRslts(5, i) = Replace(IndRslts(5, i), "-", "")
        IndRslts(12, i) = MyOverallPl(IndRslts(0, i))         'overall place

        If UCase(IndRslts(3, i)) = "M" Then       'gender place
            iMPlace = CInt(iMPlace) + 1
            IndRslts(13, i) = iMPlace
        Else
            iFPlace = CInt(iFPlace) + 1
            IndRslts(13, i) = iFPlace
        End If

        IndRslts(14, i) = MyAgeGrpPl(IndRslts(0, i), IndRslts(3, i), IndRslts(11, i))     'age grp place
        If sShowAge = "y" Then
            IndRslts(15, i) = AgeGrdTime(IndRslts(3, i), IndRslts(4, i), IndRslts(5, i))   'age graded time   
        Else
            IndRslts(15, i) = "na"
        End If  
        
        IndRslts(16, i) = PacePerMile(ConvertToSeconds(IndRslts(5, i)), sDist)
        IndRslts(17, i) = PacePerKM(ConvertToSeconds(IndRslts(5, i)), sDist)
    Next
End If

Private Function MyOverallPl(iThisBib)
    MyOverallPl = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT pr.Bib FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID "
    sql2 = sql2 & "WHERE (ir.RaceID = " & lRaceID & " AND pr.RaceID = " & lRaceID  & ") AND ir.FnlScnds > 0 AND ir.FnlTime IS NOT NULL "
	sql2 = sql2 & "AND ir.FnlTime > '00:00:00.000' ORDER BY ir.FnlScnds"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        MyOverallPl = CInt(MyOverallPl) + 1
        If CInt(rs2(0).Value) = CInt(iThisBib) Then Exit Do
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function MyAgeGrpPl(iThisBib, sThisGender, sThisAgeGrp)
    MyAgeGrpPl = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT pr.Bib FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID "
    sql2 = sql2 & "INNER JOIN Participant p ON p.ParticipantID = ir.ParticipantID WHERE (ir.RaceID = " & lRaceID & " AND pr.RaceID = " & lRaceID 
    sql2 = sql2 & ") AND p.Gender = '" & sThisGender & "' AND pr.AgeGrp = '" & sThisAgeGrp & "' AND ir.FnlTime IS NOT NULL "
	sql2 = sql2 & "AND ir.FnlTime > '00:00:00.000' ORDER BY ir.FnlScnds"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        MyAgeGrpPl = CInt(MyAgeGrpPl) + 1
        If CInt(rs2(0).Value) = CInt(iThisBib) Then Exit Do
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function AgeGrdTime(sThisGndr, iThisAge, sThisTime)
    Dim lngAgeGrDistID
    Dim sngThisTime

    sngThisTime = ConvertToSeconds(sThisTime)
    AgeGrdTime = "na"

    lngAgeGrDistID = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT AgeGrDistID FROM AgeGrDist WHERE Distance = '" & sDist & "'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then lngAgeGrDistID = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing

    If CLng(lngAgeGrDistID) > 0 Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT Factor FROM AgeGrFactors WHERE MF = '" & LCase(sThisGndr) & "' AND Age = " & iThisAge & " AND AgeGrDistID = " & lngAgeGrDistID
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then AgeGrdTime = ConvertToMinutes(CSng(sngThisTime)*CSng(rs2(0).Value))
        rs2.Close
        Set rs2 = Nothing
    End If
End Function
%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/pace_per_mile.asp" -->
<!--#include file = "../../includes/pace_per_km.asp" -->
<!--#include file = "../../includes/clean_input.asp" -->
<%
Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events Results</title>
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
<!--
<script src="https://code.jquery.com/jquery.min.js"></script>
<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css"  rel="stylesheet" type="text/css">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>
-->
  
<style>
    .bluebox {
        display: none;
        margin-left: 75px;
        padding: 5px;
        background-color: #fff;
        border: 1px solid #ccc;
    }
    @media print
    {    
        .no-print, .no-print *
        {
            display: none !important;
        }
    }
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="bluebox" style="position: absolute;top: 0;left: 0;z-index: 10000;"></div>

    <div class="row no-print">
		<div class="col-sm-10">

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

		    <%If CLng(lRaceID) = 0 Then%>
                <h3 class="h3 bg-primary">Gopher State Events Results</h3>
            <%Else%>
                <h3 class="h3 bg-primary">Gopher State Events Results: <%=sEventName%> (<%=Year(dEventDate)%>)</h3>
            <%End If%>

            <div class="col-md-6">
			    <form role="form" class="form-inline" name="which_event" method="post" action="results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>">
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
			        <input class="form-control" type="submit" name="get_event" id="get_event" value="Get These" style="font-size:0.9em;">
			    </div>
                </form>
            </div>
            <div class="col-md-6">
                <%If CLng(lEventID) = 0 Then%>
                    &nbsp;
                <%Else%>
				    <form role="form" class="form-inline" name="get_races" method="post" action="results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>">
                    <div class="form_group">
				        <label for="races">Race:</label>
				        <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()" style="font-size:0.9em;">
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
				        <input class="form-control" type="submit" name="get_race" id="get_race" value="View" style="font-size:0.9em;">
                    </div>
				    </form>
                <%End If%>
            </div>

      		<%If Not CLng(lEventID) = 0 Then%>
                <%If sTimed = "y" Then%>
		            <%If Not CLng(lRaceID) = 0 Then%>
                        <ul class="list-inline">
                            <li class="list-group-item">Total Finishers:&nbsp;<%=iTtlRcds%></li>

                            <%If UBound(Races, 2) > 1 Then%>
                                <li class="list-group-item"><%=sRaceName%> Finishers:&nbsp;<%=iNumRace%></li>
                            <%End If%>

                            <li class="list-group-item">Distance: <%=sDist%></li>

		                    <%If Not sLocation = vbNullString Then%>
                                <li class="list-group-item">Location: <%=sLocation%></li>
                            <%End If%>
                        </ul>

   			            <%If CDate(dEventDate) > Date Then%>
				            <div class="bg-info">
                                This event is currently scheduled for <%=dEventDate%>.  The results will be available on that date.
                            </div>
			            <%Else%>
                            <%If CDate(Date) < CDate(dEventDate) + 7 Then%>
			                    <%If bRsltsOfficial = False Then%>
				                    <div class="bg-info">
                                        <span style="color: red;font-size: 1.0em;font-weight: bold;">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE</span><br>
                                        Please report any issues to bob.schneider@gopherstateevents.com.
                                    </div>
			                    <%Else%>
				                    <div class="bg-info">
                                        These results are now official.  If you notice any errors please contact us 
				                        via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by telephone (612.720.8427).
                                    </div>
			                    <%End If%>
                            <%End If%>
			            <%End If%>
                        <%If CLng(lEventID) = 650 Then%>
                            <div class="bg-danger text-danger" style="text-align: right;">
                                <a href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=2&event_id=651">
                                    View Vasaloppet Sunday Results
                                </a>
                            </div>
                        <%ElseIf CLng(lEventID) = 651 Then%>
                            <div class="bg-danger text-danger" style="text-align: right;">
                                <a href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=46&event_id=650">
                                    View Vasaloppet Results
                                </a>
                            </div>
                        <%End If%>
                        <div class="row bg-success" style="margin: 10px 0 10px 0;">
                            <div class="col-xs-4">
                                <br>
                                <form class="form-inline" name="find_bib" method="post" action="results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                                    onsubmit="return chkFlds2();">
                                <div class="form_group">
                                    <label for="bib-to-find">Bib To Find:</label>
                                    <input class="form-control" type="text" name="bib_to_find" id="bib_to_find" size="3" maxlength="4" value ="<%=iBibToFind%>">
                                    <input class="form-control" type="hidden" name="submit_bib" id="submit_bib" value="submit_bib">
                                    <input class="form-control" type="submit" name="submit_lookup" id="submit_lookup" value="Find Bib">
                                </div>
                                </form>
                                <br>
                            </div>
                            <div class="col-xs-8">
                                <%If CInt(iBibToFind) > 0 Then%>
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
                                                <th>Chip Time</th>
                                                <th>Gun Time</th>
                                                <th>Chip Start</th>
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
                                <%Else%>
                                    &nbsp;
                                <%End If%>
                            </div>
                        <%End If%>
                    </div>
                    <%If Not CLng(lRaceID) = 0 Then%>
			            <ul class="list-inline">
                            <%If UBound(CustomFields, 2) > 0 Then%>
                                <%For i = 0 To UBound(CustomFields, 2) - 1%>
                                    <li class="list-group-item-warning">
                                        <a href="javascript:pop('custom_fields_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;custom_fields_id=<%=CustomFields(0, i)%>',1000,700)"><%=CustomFields(1, i)%></a>
                                    </li>
                                <%Next%>
                            <%End If%>

                            <%If sIndivRelay = "relay" Then%>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('relay_by_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)">Results 
                                    by Split</a>
                                </li>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('relay_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)" >Results 
                                    w/Splits</a>
                                </li>
                            <%End If%>

                            <%If CInt(iNumLaps) > 1 Then%>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('rslts_by_lap.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)">Results 
                                    by Lap</a>
                                </li>
                                <li class="list-group-item-success">
                                    <a href="javascript:pop('rslts_w_laps.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>',1000,700)" >Results 
                                    w/Laps</a>
                                </li>
                            <%End If%>

				            <%If sHasSplits = "y" And sGender <> "B" Then%>
                                <li class="list-group-item-warning">
                                    <a href="splits/results_w-splits.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Results With Splits</a>
                                </li>
                                <li class="list-group-item-warning">
                                    <a href="splits/rank_by_split.asp?event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Rank By Split</a>
                                </li>
                            <%End If%>
                            <li class="list-group-item-danger">
                                <a href="javascript:pop('print_rslts.asp?rslts_event_id=<%=lEventID%>&amp;gender=<%=sGender%>&amp;race_id=<%=lRaceID%>&amp;sort_rslts_by=<%=sSortRsltsBy%>',1000,700)">Print</a>
                            </li>
                            <li class="list-group-item-success">
                                <a href="dwnld_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>" 
					            onclick="openThis(this.href,1024,768);return false;">Download</a>
                            </li>
				            <%If Session("role") = "admin" Then%>
                                <li class="list-group-item-warning">
                                    <a href="usatf_results.asp?event_type=<%=iEventType%>&amp;event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
					                onclick="openThis(this.href,1024,768);return false;">USATF Rslts</a>
                                </li>
                            <%End If%>
				            <%If sHasTeams = "y" Then%>
                                <li class="list-group-item-danger">
                                    <a href="team_results.asp?race_id=<%=lRaceID%>" onclick="openThis(this.href,1024,768);return false;">Team Results</a>
                                </li>
                            <%End If%>
                                <li class="list-group-item-info">
                                    <a href="/records/records.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Records</a>
                                </li>
                            <%If CInt(iRaceType) = 5 Then%>
                                <%If sShowAge = "y" Then%>
                                    <li class="list-group-item-success">
                                        <a href="age_graded.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                        onclick="openThis(this.href,1024,768);return false;">Age-Graded</a>
                                    </li>
                                <%End If%>
                            <%End If%>
                            <%If CInt(iRaceType) >= 9 Then%>
                                <li class="list-group-item-warning">
                                    <a href="trans_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Trans Data</a>
                                </li>
                                <li class="list-group-item-warning">
				                    <a href="multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Results w/Splits</a>
                                </li>
                            <%End If%>
				            <%If sGender = "B" Then%>
                                <%If CInt(iNumMAgeGrps) > 1 Then%>
				                    <li class="list-group-item-danger">
                                        <a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=M"
                                        onclick="openThis(this.href,1024,768);return false;">Male Awards</a>
                                    </li>
                                    <li class="list-group-item-danger">
				                        <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=M"
                                        onclick="openThis(this.href,1024,768);return false;">Male Age Groups</a>
                                    </li>
                                <%End If%>
                                <%If CInt(iNumFAgeGrps) > 1 Then%>
				                    <li class="list-group-item-danger">
                                        <a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=F"
                                        onclick="openThis(this.href,1024,768);return false;">Female Awards</a>
                                    </li>
                                    <li class="list-group-item-danger">
				                        <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=F"
                                        onclick="openThis(this.href,1024,768);return false;">Female Age Groups</a>
                                    </li>
                                <%End If%>
                            <%Else%>
                                <%If CInt(iNumAgeGrps) > 1 Then%>
				                    <li class="list-group-item-danger">
                                        <a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                                        onclick="openThis(this.href,1024,768);return false;">Awards</a>
                                    </li>
                                    <li class="list-group-item-danger">
				                        <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>"
                                        onclick="openThis(this.href,1024,768);return false;">Age Groups</a>
                                    </li>
                                <%End If%>
                                <%If CLng(lSuppLegID) > 0 Then%>
                                    <li class="list-group-item-info">
                                        <a href="/results/fitness_events/supp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>" 
                                        onclick="openThis(this.href,1024,768);return false;">Rslts w/Splits</a>
                                    </li>
                                <%End If%>
			                <%End If%>
                            <%If UBound(Races, 2) > 1 And sShowAge = "y" Then%>
                                <li class="list-group-item-success">
                                    <a href="blended_results.asp?event_id=<%=lEventID%>" 
                                    onclick="openThis(this.href,1024,768);return false;">Blended Results</a>
                                </li>
                            <%End If%>
                        </ul>
                    <%End If%>

                    <div class="table-responsive">
		                <table class="table table-striped">
			                <tr>
				                <th style="padding-right: 5px;">Pl</th>
                                <th style="padding-right: 5px;">Bib</th>
				                <th style="padding-right: 5px;">Name</th>
                                <th style="padding-right: 5px;">Certificate</th>
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
			                <%For i = 0 To UBound(IndRslts, 2)%>
                                <%If IndRslts(4, i) = "99" Then%>
                                    <%IndRslts(4, i) = "--"%>
                                    <%IndRslts(11, i) = "na"%>
                                <%ElseIf sShowAge = "n" Then%>
                                    <%IndRslts(4, i) = "--"%>
                                <%End If%>
					            <tr>
						            <td style="padding-right: 5px;"><%=i + 1%>)</td>
                                    <td style="padding-right: 5px;"><%=IndRslts(0, i)%></td>
						            <td style="padding-right: 5px;">
                                        <a class="runnerName" data-bib="<%=IndRslts(0, i)%>" data-event_id="<%=lEventID%>" data-race_id="<%=lRaceID%>" 
                                                data-chip="<%=IndRslts(5, i)%>" data-gun="<%=IndRslts(6, i)%>" data-start="<%=IndRslts(7, i)%>" 
                                                data-location="<%=IndRslts(8, i)%>, <%=IndRslts(9, i)%>" data-age="<%=IndRslts(4, i)%>"  
                                            data-gender="<%=IndRslts(3, i)%>" data-event_name="<%=sEventName%>" data-event_date="<%=dEventDate%>" 
                                            data-race_name="<%=sRaceName%>" data-place="<%=IndRslts(12, i)%>"  data-gender_place="<%=IndRslts(13, i)%>"  
                                            data-age_grp_place="<%=IndRslts(14, i)%>" data-age_graded_time="<%=IndRslts(15, i)%>" 
                                            data-logo="<%=sLogo%>" data-age_group="<%=IndRslts(11, i)%>" data-per_mile="<%=IndRslts(16, i)%>"
                                            data-per_km="<%=IndRslts(17, i)%>"
                                            href="javascript:void(0)"><%=IndRslts(1, i)%>&nbsp;<%=IndRslts(2, i)%></a>
                                        </a>
                                    </td>
						            <td style="text-align:center;">
                                        <a href="javascript:pop('certificate.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>&amp;bib=<%=IndRslts(0, i)%>',1050,775)">
                                            View
                                        </a>
                                    </td>
						            <td style="text-align:center;"><%=IndRslts(3, i)%></td>
                                    <td style="padding-right: 5px;">
                                        <%If sShowAge = "y" Then%>
                                            <%=IndRslts(4, i)%>
                                        <%Else%>
                                            <%=IndRslts(11, i)%>
                                        <%End If%>
                                    </td>
						            <td style="padding-right: 5px;"><%=Replace(IndRslts(5, i), "-", "")%></td>
						            <td style="padding-right: 5px;"><%=IndRslts(6, i)%></td>
						            <td style="padding-right: 5px;"><%=IndRslts(7, i)%></td>
						            <td style="padding-right: 5px;"><%=IndRslts(8, i)%>, <%=IndRslts(9, i)%></td>
					            </tr>
			                <%Next%>
		                </table>
                    </div>
                <%Else%>
                    <p>This was a non-timed race.</p>
                <%End If%>
            <%End If%>
        </div>
		<div class="col-sm-2">
            <br>
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
	<!--#include file = "../../includes/footer.asp" -->
</div>

<!-- START "jQuery" -->
<script>
    $(function () {
        // listen for click of runner name
        $('.runnerName').on('click', function (e) {
            var width = e.pageX ? e.pageX : e.clientX + document.body.scrollLeft + document.documentElement.scrollLeft;
            var height = e.pageY ? e.pageY : e.clientY + document.body.scrollTop + document.documentElement.scrollTop;
            $('.bluebox').css({
                left: width + 60 + 'px',
                top: height - 10 + 'px'
            }).show();
            // get runners name from text in <a> tag
            var runnerName = $(this).text();
            var runnerBib = $(this).data('bib');
            var runnerChip = $(this).data('chip');
            var runnerGun = $(this).data('gun');
            var runnerStart = $(this).data('start');
            var runnerLocation = $(this).data('location');
            var runnerAge = $(this).data('age');
            var runnerGender = $(this).data('gender');
            var runnerLogo = $(this).data('logo');
            var runnerEvent = $(this).data('event_name');
            var runnerDate = $(this).data('event_date');
            var runnerRace = $(this).data('race_name');
            var runnerPlace = $(this).data('place');
            var runnerGenderPlace = $(this).data('gender_place');
            var runnerAgeGrpPlace = $(this).data('age_grp_place');
            var runnerAgeGradedTime = $(this).data('age_graded_time');
            var runnerAgeGroup = $(this).data('age_group');
            var runnerPerMile = $(this).data('per_mile');
            var runnerPerKM = $(this).data('per_km');

            $('.bluebox').html('<a href="javascript:window.print();">Print</a><h4 class="h4 bg-success"><img src="' + runnerLogo 
            + '" width="75">My Results - ' + runnerEvent + '</h4><h5>'+ runnerDate + '<br>' + runnerRace + '</h5><h4>' + runnerName 
            + ' (Bib: ' + runnerBib + ')</h4>Gender: ' + runnerGender + '<br>Age: ' + runnerAge + '<br>Age Group: ' + runnerAgeGroup 
            + '<br>Location: ' + runnerLocation + '<br><br>Time Data<br>Chip Time: ' + runnerChip + '<br>Gun Time: ' + runnerGun 
            + '<br>Start Delay: ' + runnerStart + '<br>Age Graded Time: ' + runnerAgeGradedTime + '<br>Per Mile: ' + runnerPerMile 
            + '<br>Per KM: ' + runnerPerKM + '<br><br>Place Data<br>Overall Place: ' 
            + runnerPlace + '<br>Gender Place: ' + runnerGenderPlace + '<br>Age Group Place: ' + runnerAgeGrpPlace)
        });
    });
</script>
<!-- END -->
<script src="js/ga.js"></script>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>