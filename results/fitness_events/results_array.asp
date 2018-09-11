<%@ Language=VBScript%>
<%
Option Explicit
%>

<!--#include file = "../includes/JSON_2.0.4.asp" -->

<%
Dim conn, rs, sql
Dim i, j

Dim sFilter, sSortBy, sSortDir, sRsltsPage
Dim iLength

Dim IndRslts, Races

Dim lRaceID, lEventID, lSuppLegID, lFeaturedEventsID
Dim iRaceType, iTtlRcds, iEventType, iNumAgeGrps, iNumRace, iNumMAgeGrps, iNumFAgeGrps
Dim sEventName, sGender, sSortRsltsBy, sDist, sRaceName, sGalleryLink, sLogo, sShowAge, sHasTeams
Dim sWeather, sEventRaces, sLocation, sHasSplits, sIndivRelay, sTimed, sRaceReport
Dim dEventDate

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect "http://www.google.com"
If CLng(lEventID) < 0 Then Response.Redirect "http://www.google.com"

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect "http://www.google.com"
If CLng(lRaceID) < 0 Then Response.Redirect "http://www.google.com"

sGender = Request.QueryString("gender")
If sGender = vbNullString Then sGender = "B"
If Len(sGender) > 1 Then Response.Redirect "http://www.google.com"

sFilter = Request.QueryString("results_filter")
If sFilter = "undefined" Then sFilter = vbNullString

iLength = Request.QueryString("results_length")
If CStr(iLength) = "undefined" Then iLength = 100
If CStr(iLength) = vbNullString Then iLength = 100
If CInt(iLength) < 0 Then Response.Redirect "http://www.google.com"

sSortBy = Request.QueryString("results_sort")
sSortDir = Request.QueryString("results_sort_direction")
sRsltsPage = Request.QueryString("results_page")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

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
    sql = "SELECT SuppLegID FROM SuppLeg WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then lSuppLegID = rs(0).Value
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ") AND FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount  > 0 Then iTtlRcds = rs.RecordCount
    rs.Close
    Set rs = Nothing

	'get event information
	sql = "SELECT EventName, EventDate, Location, Logo FROM Events WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
    sLocation = rs(2).Value
    sLogo = rs(3).Value
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
	
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT GalleryLink FROM RaceGallery WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        If Not rs(0).Value & "" = "" Then sGalleryLink = rs(0).Value
    End If
    rs.Close
    Set rs = Nothing
	
    If CLng(lRaceID) = 0 Then lRaceID = GetFirstRace()

    'num race finishers
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    iNumRace = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'check for team results
    sHasTeams = "n"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID FROM Teams WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sHasTeams = "y"
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
	sql = "SELECT Dist, RaceName, Type, NumSplits, IndivRelay, Timed, ShowAge FROM RaceData WHERE RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	sDist = rs(0).Value
	sRaceName = rs(1).Value
	iRaceType = rs(2).Value
    If CInt(rs(3).Value) > 0 Then sHasSplits = "y"
    sIndivRelay = rs(4).Value
    sTimed = rs(5).Value
    sShowAge = rs(6).Value
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

If sSortBy = vbNullString Then
	sql = "SELECT SortRsltsBy FROM RaceData WHERE RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
    sSortBy = rs(0).Value
	Set rs = Nothing

    sSortRsltsBy = "FnlTime"
    If sSortRsltsBy = "FnlTime" Then
        sOrderBy = "ir.FnlScnds"
    Else
        sOrderBy = "ir.EventPl"
    End If
End If

Select Case sSortBy
    Case "Bib"
        sSortBy = "Bib " & sSortDir
    Case "Name"
        sSortBy = "Name " & sSortDir
    Case "Gender"
        sSortBy = "Gender " & sSortDir
    Case "Age"
        sSortBy = "Age " & sSortDir
    Case "ChipTime"
        sSortBy = "ChipTime " & sSortDir
    Case "GunTime"
        sSortBy = "GunTime " & sSortDir
End Select


If sGender = "B" Then
    Set rs = Server.CreateObject("ADODB.Recordset")  
    sql = "Exec OverallResults @RaceID = " & lRaceID
    Set rs = conn.execute(sql) 
    If rs.BOF and rs.EOF Then
        ReDim IndRslts(9, 0)
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
        ReDim IndRslts(9, 0)
    Else
        IndRslts = rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
End If

For i = 0 To UBound(IndRslts, 2)
    If sShowAge = "n" Then
        IndRslts(4, i) = MyAgeGrp(IndRslts(0, i))
    Else
		If IndRslts(4, i) = "99" Then IndRslts(4, i) = "0"
    End If
Next

For i = 0 To UBound(IndRslts, 2)
    IndRslts(0, i) = i + 1
Next

conn.Close
Set conn = Nothing
%>

{
  "data": [
        <%For i = 0 To UBound(IndRslts, 2)%>
            [
            <%For j = 0 To 9%>
                "<%=IndRslts(j, i)%>"
                <%If j < 3 Then %>
                    <%Response.Write ","       '-- don't output a comma on the last element%>
                  <%End If%>
                <%Next%>
            ]

            <%If i < UBound(IndRslts, 2) Then %>
                <%Response.Write ","         '-- don't output a comma on the last section%>
            <%End If%>
        <%Next%>
]
}
