<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j
Dim iThisYear, iThisYear2
Dim iNumFtnsEvnts, iNumFtnsRaces, iNumFtnsParts, iNumFtnsEmail, iNumFtnsFinishers, iNumFtnsMale, iNumFtnsFemale
Dim iNumMaleFinishers, iNumFemaleFinishers, sngPctMaleFinishers
Dim iNumCCMeets, iNumNordMeets, iNumCCRunMeets, iNumCCTeams, iNumCCRunTeams, iNumNordTeams, iNumCCParts, iNumCCRunParts, iNumNordParts
Dim iNumCCRaces, iNumCCRunRaces, iNumNordRaces, iNumCCFinishers, iNumCCRunFinishers, iNumNordFinishers, iNumEvnts, iNumRaces, iNumParts
Dim sngPctFemaleFinishers, sngFtnsInvoice, sngInvoice, sngCCInvoice
Dim sShowWhat, sYrlyBrkdwn, sYrlySumm
Dim YrlyData(), YrlyTtls(5), YrlyPct(4), YrlySumm()

Server.ScriptTimeout = 1200

If Not Session("role") = "admin" Then Response.Redirect "/index.html"

iThisYear = Request.QueryString("this_year")
If CStr(iThisYear) = vbNullString Then iThisYear = 0

sShowWhat = Request.QueryString("show_what")
sYrlyBrkdwn = Request.QueryString("yrly_brkdwn")
sYrlySumm = Request.QueryString("yrly_summ")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'**************************************************************************************************
'upload age graded factors into db
'Dim Filepath
'Dim fs
'Dim file    
'Dim TextStream		
'Dim Line
'Dim sSplit
'Dim field1, field2, field3, field4

'Const ForReading = 1, ForWriting = 2, ForAppending = 3
'Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

'Set fs = Server.CreateObject("Scripting.FileSystemObject")

'Filepath = "C:\inetpub\h51web\gopherstateevents\admin\age_factors.txt"

'Set file = fs.GetFile(Filepath)
'Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)
		
'Do While Not TextStream.AtEndOfStream
'	Line = TextStream.readline
'	sSplit =  Split(Line, vbTab)	

'	field1 = Trim(sSplit(0))		'race
'	field2 = Trim(sSplit(1))		'gender	
'	field3 = Trim(sSplit(2))		'age	
'	field4 = Trim(sSplit(3))		'factor
                   
	'insert into the partdata table
'	sql = "INSERT INTO AgeGrFactors (AgeGrDistID, MF, Age, Factor) VALUES ('" & field1 & "', '"  & field2 & "', '" & field3 & "', '" & field4 & "')"			
'	Set rs=conn.Execute(sql)
'	Set rs=Nothing
'Loop
'Set TextStream = nothing

'Set fs = nothing
'**************************************************************************************************

'insert into db
If sShowWhat = "fitness" Then
    'get fitness invoices
    sngFtnsInvoice = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice FROM Events WHERE Invoice IS NOT NULL"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngFtnsInvoice = CSng(sngFtnsInvoice) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    sngFtnsInvoice = FormatCurrency(sngFtnsInvoice)

    'get num fitness events
    iNumFtnsEvnts = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID FROM Events WHERE EventDate <= '" & Date & "'"
    rs.Open sql, conn, 1, 2
    iNumFtnsEvnts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num fitness races
    iNumFtnsRaces = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT r.RaceID FROM RaceData r INNER JOIN Events e ON r.EventID = e.EventID WHERE e.EventDate <= '" & Date & "'"
    rs.Open sql, conn, 1, 2
    iNumFtnsRaces = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num fitness participants
    iNumFtnsParts = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM PartRace"
    rs.Open sql, conn, 1, 2
    iNumFtnsParts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num male fitness participants
    iNumFtnsFemale = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM Participant WHERE Gender = 'F' OR Gender = 'f'"
    rs.Open sql, conn, 1, 2
    iNumFtnsFemale = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num fitness participants
    iNumFtnsMale = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM Participant WHERE Gender = 'M' OR Gender = 'm'"
    rs.Open sql, conn, 1, 2
    iNumFtnsMale = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num emails
    iNumFtnsEmail = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Email FROM Participant WHERE Email IS NOT NULL"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Not rs(0).Value & "" = "" Then iNumFtnsEmail = CLng(iNumFtnsEmail) + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get num finishers
    iNumFtnsFinishers = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM IndResults"
    rs.Open sql, conn, 1, 2
    iNumFtnsFinishers = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num finishers
    iNumMaleFinishers = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE p.Gender = 'M' "
    sql = sql & "OR p.Gender = 'm'"
    rs.Open sql, conn, 1, 2
    iNumMaleFinishers = rs.RecordCount
    rs.Close
    Set rs = Nothing
	
    sngPctMaleFinishers = Round(CLng(iNumMaleFinishers)/CSng(iNumFtnsFinishers)*100, 2)

    'get num finishers
    iNumFemaleFinishers = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE p.Gender = 'F' "
    sql = sql & "OR p.Gender = 'f'"
    rs.Open sql, conn, 1, 2
    iNumFemaleFinishers = rs.RecordCount
    rs.Close
    Set rs = Nothing
	
    sngPctFemaleFinishers = Round(CLng(iNumFemaleFinishers)/CSng(iNumFtnsFinishers)*100, 2)
ElseIf sShowWhat = "ccmeet" Then
    iNumCCMeets = 0
    iNumCCRunMeets = 0
    iNumNordMeets = 0

    iNumCCParts = 0
    iNumCCRunParts = 0
    iNumNordParts = 0

    iNumCCTeams = 0
    iNumCCRunTeams = 0
    iNumNordTeams = 0

    iNumCCRaces = 0
    iNumCCRunRaces = 0
    iNumNordRaces = 0

    iNumCCFinishers = 0
    iNumCCRunFinishers = 0
    iNumNordFinishers = 0

    sngCCInvoice = 0

    'get num cc meets
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice FROM Meets"
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        iNumCCMeets = CInt(iNumCCMeets) + 1
        If Not rs(0).Value & "" = "" Then sngCCInvoice = CSng(sngCCInvoice) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get num cc runmeets
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetsID FROM Meets WHERE Sport = 'Cross-Country'"
    rs.Open sql, conn2, 1, 2
    iNumCCRunMeets = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num nordic meets
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetsID FROM Meets WHERE Sport = 'Nordic Ski'"
    rs.Open sql, conn2, 1, 2
    iNumNordMeets = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc races
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RacesID FROM Races"
    rs.Open sql, conn2, 1, 2
    iNumCCRaces = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc teams
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID FROM Teams"
    rs.Open sql, conn2, 1, 2
    iNumCCTeams = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc run teams
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID FROM Teams WHERE Sport = 'Cross-Country'"
    rs.Open sql, conn2, 1, 2
    iNumCCRunTeams = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num nordic teams
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID FROM Teams WHERE Sport = 'Nordic Ski'"
    rs.Open sql, conn2, 1, 2
    iNumNordTeams = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc participants
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID FROM Roster"
    rs.Open sql, conn2, 1, 2
    iNumCCParts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc run participants
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT r.RosterID FROM Roster r INNER JOIN Teams t ON r.TeamsID = t.TeamsID  WHERE t.Sport = 'Cross-Country'"
    rs.Open sql, conn2, 1, 2
    iNumCCRunParts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num nordic participants
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT r.RosterID FROM Roster r INNER JOIN Teams t ON r.TeamsID = t.TeamsID  WHERE t.Sport = 'Nordic Ski'"
    rs.Open sql, conn2, 1, 2
    iNumNordParts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc finishers
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM IndRslts WHERE Place <> 0"
    rs.Open sql, conn2, 1, 2
    iNumCCFinishers = rs.RecordCount
    rs.Close
    Set rs = Nothing
Else
    'get num fitness events
    iNumEvnts = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID FROM Events Where EventDate <= '" & Date & "'"
    rs.Open sql, conn, 1, 2
    iNumEvnts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc meets
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetsID FROM Meets WHERE MeetDate <= '" & Date & "'"
    rs.Open sql, conn2, 1, 2
    iNumEvnts = CInt(iNumEvnts) + rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num fitness races
    iNumRaces = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT r.RaceID FROM RaceData r INNER JOIN Events e ON r.EventID = e.EventID"
    rs.Open sql, conn, 1, 2
    iNumRaces = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc races
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT r.RacesID FROM Races r INNER JOIN Meets m ON r.MeetsID = m.MeetsID WHERE m.MeetDate<= '" & Date & "'"
    rs.Open sql, conn2, 1, 2
    iNumRaces =  CInt(iNumRaces) + rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num fitness participants
    iNumParts = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM PartRace"
    rs.Open sql, conn, 1, 2
    iNumParts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get num cc participants
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID FROM Roster"
    rs.Open sql, conn2, 1, 2
    iNumParts =  CLng(iNumParts) + rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get fitness invoices
    sngInvoice = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice FROM Events WHERE Invoice IS NOT NULL"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sngInvoice = CSng(sngInvoice) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get fitness invoices
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice FROM Meets WHERE Invoice IS NOT NULL"
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        sngInvoice = CSng(sngInvoice) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Sub GetYrly(lWhichYear)
    Dim x
    Dim iMTotal, iFTotal

    x = 0
    ReDim YrlyData(8, 0)
    If sShowWhat = "fitness" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventID, EventName, EventDate, Invoice FROM Events WHERE EventDate <= '" & Date & "' ORDER BY EventDate"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Year(CDate(rs(2).Value)) = CInt(lWhichYear) Then
                YrlyData(0, x) = rs(0).Value
                YrlyData(1, x) = Replace(rs(1).Value, "''", "'")
                YrlyData(2, x) = rs(2).Value
                YrlyData(3, x) = NumRaces(rs(0).Value)
                YrlyData(4, x) = NumMale(rs(0).Value)
                YrlyData(5, x) = NumFemale(rs(0).Value)
                YrlyData(6, x) = CInt(YrlyData(4, x)) + Cint(YrlyData(5, x))
                YrlyData(7, x) = NumEmail(rs(0).Value)
                YrlyData(8, x) = rs(3).Value

                YrlyTtls(0) = CInt(YrlyTtls(0)) + CInt(YrlyData(3, x))
                YrlyTtls(1) = CInt(YrlyTtls(1)) + CInt(YrlyData(4, x))
                YrlyTtls(2) = CInt(YrlyTtls(2)) + CInt(YrlyData(5, x))
                YrlyTtls(3) = CInt(YrlyTtls(3)) + CInt(YrlyData(6, x))
                YrlyTtls(4) = CInt(YrlyTtls(4)) + CInt(YrlyData(7, x))
                YrlyTtls(5) = CSng(YrlyTtls(5)) + CSng(YrlyData(8, x))

                iMTotal = CInt(iMTotal) + YrlyData(4, x)
                iFTotal = CInt(iFTotal) + YrlyData(5, x)

                If CInt(YrlyData(4, x)) = 0 Then
                    YrlyData(4, x) = YrlyData(4, x) & " (0%)"
                Else
                    YrlyData(4, x) = YrlyData(4, x) & " (" & Round(CInt(YrlyData(4, x))/CInt(YrlyData(6, x))*100, 2) & "%)"
                End If

                If CInt(YrlyData(5, x)) = 0 Then
                    YrlyData(5, x) = YrlyData(5, x) & " (0%)"
                Else
                    YrlyData(5, x) = YrlyData(5, x) & " (" & Round(CInt(YrlyData(5, x))/CInt(YrlyData(6, x))*100, 2) & "%)"
                End If

                x = x + 1
                ReDim Preserve YrlyData(8, x)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        YrlyPct(0) = Round(CInt(YrlyTtls(0))/(UBound(YrlyData, 2) - 1), 2)
        If Not CInt(iMTotal) = 0 Then YrlyPct(1) = Round(CInt(iMTotal)/(CInt(iMTotal) + CInt(iFTotal))*100, 2)
        If Not CInt(iFTotal) = 0 Then YrlyPct(2) = Round(CInt(iFTotal)/(CInt(iMTotal) + CInt(iFTotal))*100, 2)
        YrlyPct(3) = Round((CInt(iMTotal) + CInt(iFTotal))/(UBound(YrlyData, 2)), 2)
        YrlyPct(4) = Round(CSng(YrlyTtls(5))/(UBound(YrlyData, 2)), 2)
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetsID, MeetName, MeetDate, Invoice FROM Meets WHERE MeetDate <= '" & Date & "' ORDER BY MeetDate"
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If Year(CDate(rs(2).Value)) = CInt(lWhichYear) Then
                YrlyData(0, x) = rs(0).Value
                YrlyData(1, x) = Replace(rs(1).Value, "''", "'")
                YrlyData(2, x) = rs(2).Value
                YrlyData(3, x) = NumRaces(rs(0).Value)
                YrlyData(4, x) = NumMale(rs(0).Value)
                YrlyData(5, x) = NumFemale(rs(0).Value)
                YrlyData(6, x) = CInt(YrlyData(4, x)) + Cint(YrlyData(5, x))
                YrlyData(8, x) = rs(3).Value

                YrlyTtls(0) = CInt(YrlyTtls(0)) + CInt(YrlyData(3, x))
                YrlyTtls(1) = CInt(YrlyTtls(1)) + CInt(YrlyData(4, x))
                YrlyTtls(2) = CInt(YrlyTtls(2)) + CInt(YrlyData(5, x))
                YrlyTtls(3) = CInt(YrlyTtls(3)) + CInt(YrlyData(6, x))
                YrlyTtls(4) = CInt(YrlyTtls(4)) + CInt(YrlyData(7, x))
                YrlyTtls(5) = CSng(YrlyTtls(5)) + CSng(YrlyData(8, x))

                iMTotal = CInt(iMTotal) + YrlyData(4, x)
                iFTotal = CInt(iFTotal) + YrlyData(5, x)

                If CInt(YrlyData(4, x)) = 0 Then
                    YrlyData(4, x) = YrlyData(4, x) & " (0%)"
                Else
                    YrlyData(4, x) = YrlyData(4, x) & " (" & Round(CInt(YrlyData(4, x))/CInt(YrlyData(6, x))*100, 2) & "%)"
                End If

                If CInt(YrlyData(5, x)) = 0 Then
                    YrlyData(5, x) = YrlyData(5, x) & " (0%)"
                Else
                    YrlyData(5, x) = YrlyData(5, x) & " (" & Round(CInt(YrlyData(5, x))/CInt(YrlyData(6, x))*100, 2) & "%)"
                End If

                x = x + 1
                ReDim Preserve YrlyData(8, x)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        YrlyPct(0) = Round(CInt(YrlyTtls(0))/(UBound(YrlyData, 2) - 1), 2)
        If Not CInt(iMTotal) = 0 Then YrlyPct(1) = Round(CInt(iMTotal)/(CInt(iMTotal) + CInt(iFTotal))*100, 2)
        If Not CInt(iFTotal) = 0 Then YrlyPct(2) = Round(CInt(iFTotal)/(CInt(iMTotal) + CInt(iFTotal))*100, 2)
        YrlyPct(3) = Round((CInt(iMTotal) + CInt(iFTotal))/(UBound(YrlyData, 2)), 2)
        YrlyPct(4) = Round(CSng(YrlyTtls(5))/(UBound(YrlyData, 2)), 2)
    End If
End Sub

Private Sub GetYrlySumm(lWhichYear)
    Dim x
    Dim YrlyEvnts()

    If sShowWhat = "fitness" Then
        x = 0
        ReDim YrlyEvnts(0)
        ReDim YrlySumm(10)
        YrlySumm(10) = 0
        YrlySumm(0) = i
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventDate, EventID, Invoice FROM Events WHERE EventDate <= '" & Date & "'"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Year(CDate(rs(0).Value)) = CInt(lWhichYear) Then
                YrlyEvnts(x) = rs(1).Value
                x = x + 1
                ReDim Preserve YrlyEvnts(x)
                YrlySumm(1) = CInt(YrlySumm(1)) + 1
                If CSng(rs(2).Value) > 0 Then
                    YrlySumm(10) = CSng(YrlySumm(10)) + CSng(rs(2).Value)
                Else
                    YrlySumm(10) = CSng(YrlySumm(10)) + CSng(GetInvoice(rs(1).Value, "Fitness Event"))
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        For x = 0 To UBound(YrlyEvnts) - 1
            'get num races
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT RaceID FROM RaceData WHERE EventID = " & YrlyEVnts(x)
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                YrlySumm(2) = CInt(YrlySumm(2)) + 1
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            'get num male parts
            YrlySumm(3) = Cint(YrlySumm(3)) + NumMale(YrlyEVnts(x))
            YrlySumm(5) = Cint(YrlySumm(5)) + NumFemale(YrlyEVnts(x))

            YrlySumm(4) = 0
            YrlySumm(6) = 0
            If Not CInt(YrlySumm(3)) = 0 Then YrlySumm(4) = Round(CInt(YrlySumm(3))/(CInt(YrlySumm(3)) + CInt(YrlySumm(5)))*100, 2)
            If Not CInt(YrlySumm(5)) = 0 Then YrlySumm(6) = Round(CInt(YrlySumm(5))/(CInt(YrlySumm(3)) + CInt(YrlySumm(5)))*100, 2)

            YrlySumm(7) = Cint(YrlySumm(5)) + Cint(YrlySumm(3))
            YrlySumm(8) = Round(Cint(YrlySumm(7))/Cint(YrlySumm(1)), 2)
            YrlySumm(9) = Cint(YrlySumm(9)) + NumEmail(YrlyEVnts(x))
        Next
    Else
        x = 0
        ReDim YrlyEvnts(0)
        ReDim YrlySumm(10)
        YrlySumm(10) = 0
        YrlySumm(0) = i
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetDate, MeetsID, Invoice, Sport FROM Meets WHERE MeetDate <= '" & Date & "'"
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If Year(CDate(rs(0).Value)) = CInt(lWhichYear) Then
                YrlyEvnts(x) = rs(1).Value
                x = x + 1
                ReDim Preserve YrlyEvnts(x)
                YrlySumm(1) = CInt(YrlySumm(1)) + 1
                If CSng(rs(2).Value) > 0 Then
                    YrlySumm(10) = CSng(YrlySumm(10)) + CSng(rs(2).Value)
                Else
                    YrlySumm(10) = CSng(YrlySumm(10)) + CSng(GetInvoice(rs(1).Value, rs(3).Value))
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        For x = 0 To UBound(YrlyEvnts) - 1
            'get num races
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT RacesID FROM Races WHERE MeetsID = " & YrlyEVnts(x)
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                YrlySumm(2) = CInt(YrlySumm(2)) + 1
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            'get num male parts
            YrlySumm(3) = Cint(YrlySumm(3)) + NumMale(YrlyEVnts(x))
            YrlySumm(5) = Cint(YrlySumm(5)) + NumFemale(YrlyEVnts(x))

            YrlySumm(4) = 0
            YrlySumm(6) = 0
            If Not CInt(YrlySumm(3)) = 0 Then YrlySumm(4) = Round(CInt(YrlySumm(3))/(CInt(YrlySumm(3)) + CInt(YrlySumm(5)))*100, 2)
            If Not CInt(YrlySumm(5)) = 0 Then YrlySumm(6) = Round(CInt(YrlySumm(5))/(CInt(YrlySumm(3)) + CInt(YrlySumm(5)))*100, 2)

            YrlySumm(7) = CLng(YrlySumm(5)) + CLng(YrlySumm(3))
            YrlySumm(8) = Round(CLng(YrlySumm(7))/CLng(YrlySumm(1)), 2)
        Next
    End If
End Sub

Private Function NumRaces(lThisEvent)
    NumRaces = 0

    If sShowWhat = "fitness" Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT RaceID FROM RaceData WHERE EventID = " & lThisEvent
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then NumRaces = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    Else
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT RacesID FROM Races WHERE MeetsID = " & lThisEvent
        rs2.Open sql2, conn2, 1, 2
        If rs2.RecordCount > 0 Then NumRaces = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

Private Function NumFemale(lThisEvent)
    NumFemale = 0

    If sShowWhat = "fitness" Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT p.ParticipantID FROM Participant p INNER JOIN PartRace pr on p.ParticipantID = pr.ParticipantID INNER JOIN RaceData rd "
        sql2 = sql2 & "ON pr.RaceID = rd.RaceID WHERE rd.EventID = " & lThisEvent & " AND (p.Gender = 'F' OR p.Gender = 'f')"
        rs2.Open sql2, conn, 1, 2
        NumFemale = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    Else
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT r.RosterID FROM Roster r INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID WHERE ir.MeetsID = " & lThisEvent 
        sql2 = sql2 & " AND (r.Gender = 'F' OR r.Gender = 'f')"
        rs2.Open sql2, conn2, 1, 2
        NumFemale = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

Private Function NumMale(lThisEvent)
    NumMale = 0

    If sShowWhat = "fitness" Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT p.ParticipantID FROM Participant p INNER JOIN PartRace pr on p.ParticipantID = pr.ParticipantID INNER JOIN RaceData rd "
        sql2 = sql2 & "ON pr.RaceID = rd.RaceID WHERE rd.EventID = " & lThisEvent & " AND (p.Gender = 'M' OR p.Gender = 'm')"
        rs2.Open sql2, conn, 1, 2
        NumMale = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    Else
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT r.RosterID FROM Roster r INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID WHERE ir.MeetsID = " & lThisEvent 
        sql2 = sql2 & " AND (r.Gender = 'M' OR r.Gender = 'm')"
        rs2.Open sql2, conn2, 1, 2
        NumMale = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

Private Function NumEmail(lThisEvent)
    NumEmail = 0

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT p.Email FROM Participant p INNER JOIN PartRace pr on p.ParticipantID = pr.ParticipantID INNER JOIN RaceData rd "
    sql2 = sql2 & "ON pr.RaceID = rd.RaceID WHERE rd.EventID = " & lThisEvent & " AND p.Email IS NOT NULL"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        If Not rs2(0).Value & "" = "" Then NumEmail = CInt(NumEmail) + 1
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetInvoice(lThisEvent, sThisSport)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Invoice FROM FinanceEvents WHERE EventID = " & lThisEvent & " AND Sport = '" & sThisSport & "'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetInvoice = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing

'Response.Write GetInvoice & "<br>"

    If CStr(GetInvoice) = vbNullString Then GetInvoice = 0
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE&copy; Admin Home</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		
		<div class="col-md-10">
            <div style="text-align:right;">
                <%If sShowWhat = "fitness" Then%>
                    <a href="admin.asp">Hide</a>
                    &nbsp;|&nbsp;
                    <a href="admin.asp?show_what=ccmeet">Show CCMeet</a>
                <%ElseIf sShowWhat = "ccmeet" Then%>
                    <a href="admin.asp?show_what=fitness">Show Fitness</a>
                    &nbsp;|&nbsp;
                <a href="admin.asp">Hide</a>
                <%Else%>
                    <a href="admin.asp?show_what=fitness">Show Fitness</a>
                    &nbsp;|&nbsp;
                    <a href="admin.asp?show_what=ccmeet">Show CCMeet</a>
                <%End If%>
            </div>

			<h3 class="h3">GSE Admin Portal</h3>
			
            <%If sShowWhat = "fitness" Then%>
                <div>
                    <h4 class="h4">Fitness Events To Date</h4>

                    <table class="table">
                        <tr>
                            <th>Events:</th>
                            <td><%=iNumFtnsEvnts%></td>
                            <th>Races:</th>
                            <td><%=iNumFtnsRaces%></td>
                            <th>Parts:</th>
                            <td><%=iNumFtnsParts%></td>
                            <th>Finishers:</th>
                            <td><%=iNumFtnsFinishers%></td>
                            <th>w/Email:</th>
                            <td><%=iNumFtnsEmail%></td>
                        </tr>
                       <tr>
                            <th>M Parts:</th>
                            <td><%=iNumFtnsMale%></td>
                            <th>F Parts:</th>
                            <td><%=iNumFtnsFemale%></td>
                            <th>M Fin:</th>
                            <td><%=iNumMaleFinishers%> (<%=sngPctMaleFinishers%>%)</td>
                            <th>F Fin:</th>
                            <td><%=iNumFemaleFinishers%> (<%=sngPctFemaleFinishers%>%)</td>
                            <th>Invoice:</th>
                            <td><%=sngFtnsInvoice%></td>
                      </tr>
                    </table>

                    <div class="bg-info">
                        <%If sYrlySumm = "y" Then%>
                             <a href="admin.asp?show_what=fitness&amp;yrly_summ=n&amp;yrly_brkdwn=<%=sYrlyBrkdwn%>">Hide Year-by-Year Summary</a>
                        <%Else%>
                            <a href="admin.asp?show_what=fitness&amp;yrly_summ=y&amp;yrly_brkdwn=<%=sYrlyBrkdwn%>">Show Year-by-Year Summary</a>
                        <%End If%>
                        &nbsp;|&nbsp;
                        <%If sYrlyBrkdwn = "y" Then%>
                             <a href="admin.asp?show_what=fitness&amp;yrly_summ=<%=sYrlySumm%>&amp;yrly_brkdwn=n">Hide Yearly Breakdown</a>
                        <%Else%>
                            <a href="admin.asp?show_what=fitness&amp;yrly_summ=<%=sYrlySumm%>y&amp;yrly_brkdwn=y">Show Yearly Breakdown</a>
                        <%End If%>
                    </div>

                    <%If sYrlySumm = "y" Then%>
                        <h5 class="h5">Summary by Year</h5>
                        <table class="table table-striped">
                            <tr>
                                <th>Year</th>
                                <th>Events</th>
                                <th>Races</th>
                                <th>Male</th>
                                <th>Male %</th>
                                <th>Female</th>
                                <th>Female %</th>
                                <th>Total</th>
                                <th>Evnt Size</th>
                                <th>Email</th>
                                <th>Invoice</th>
                                <th>Avg</th>
                            </tr>
                             <%For i = Year(Date) To 2002 Step -1%>
                                <%Call GetYrlySumm(i)%>
                                <tr>
                                    <td><%=YrlySumm(0)%></td>
                                    <td><%=YrlySumm(1)%></td>
                                    <td><%=YrlySumm(2)%></td>
                                    <td><%=YrlySumm(3)%></td>
                                    <td><%=YrlySumm(4)%>%</td>
                                    <td><%=YrlySumm(5)%></td>
                                    <td><%=YrlySumm(6)%>%</td>
                                    <td><%=YrlySumm(7)%></td>
                                    <td><%=YrlySumm(8)%></td>
                                    <td><%=YrlySumm(9)%></td>
                                    <td><%=FormatCurrency(YrlySumm(10))%></td>
                                        <%If YrlySumm(10) & "" = "" Then%>
                                        <td>n/a</td>
                                    <%ElseIf CSng(YrlySumm(1)) = 0 Then%>
                                        <td>n/a</td>
                                    <%Else%>
                                        <td><%=FormatCurrency(YrlySumm(10)/YrlySumm(1))%></td>
                                    <%End If%>
                                </tr>
                            <%Next%>
                        </table>
                    <%End If%>

                    <%If sYrlyBrkdwn = "y" Then%>
                        <h5 class="h5">Yearly Breakdown</h5>
                        <%For i = Year(Date) To 2002 Step -1%>
                            <%If CInt(iThisYear) = CInt(i) Then%>
                                <%Call GetYrly(iThisYear)%>
                                <h5 class="h5">
                                    <a href="admin.asp?show_what=fitness&amp;yrly_summ=<%=sYrlySumm%>y&amp;yrly_brkdwn=y"><%=iThisYear%> (click to hide)</a>
                                </h5>
                               <table class="table table-striped">
                                    <tr>
                                        <th>No.</th>
                                        <th>Event</th>
                                        <th>Date</th>
                                        <th>Races</th>
                                        <th>Male (%)</th>
                                        <th>Female (%)</th>
                                        <th>Total</th>
                                        <th>Email</th>
                                        <th>Invoice</th>
                                    </tr>
                                    <%For j = 0 To UBound(YrlyData, 2) - 1%>
                                        <tr>
                                            <td><%=j + 1%>)</td>
                                            <td><%=YrlyData(1, j)%></td>
                                            <td><%=YrlyData(2, j)%></td>
                                            <td><%=YrlyData(3, j)%></td>
                                            <td><%=YrlyData(4, j)%></td>
                                            <td><%=YrlyData(5, j)%></td>
                                            <td><%=YrlyData(6, j)%></td>
                                            <td><%=YrlyData(7, j)%></td>
                                            <td><%=FormatCurrency(YrlyData(8, j))%></td>
                                        </tr>
                                   <%Next%>
                                   <tr>
                                        <th class="totals" colspan="3">Totals:</th>
                                        <th class="totals"><%=YrlyTtls(0)%></th>
                                        <th class="totals"><%=YrlyTtls(1)%></th>
                                        <th class="totals"><%=YrlyTtls(2)%></th>
                                        <th class="totals"><%=YrlyTtls(3)%></th>
                                        <th class="totals"><%=YrlyTtls(4)%></th>
                                        <th class="totals"><%=FormatCurrency(YrlyTtls(5))%></th>
                                    </tr>
                                   <tr>
                                        <th class="totals" colspan="3">Percent/Average:</th>
                                        <th class="totals"><%=YrlyPct(0)%></th>
                                        <th class="totals"><%=YrlyPct(1)%>%</th>
                                        <th class="totals"><%=YrlyPct(2)%>%</th>
                                        <th class="totals"><%=YrlyPct(3)%></th>
                                        <th>&nbsp;</th>
                                        <th class="totals"><%=FormatCurrency(YrlyPct(4))%></th>
                                    </tr>
                                </table>
                            <%Else%>
                                <h5 class="h5">
                                    <a href="admin.asp?show_what=fitness&amp;this_year=<%=i%>&amp;yrly_summ=<%=sYrlySumm%>y&amp;yrly_brkdwn=y"><%=i%> (click to view)</a>
                                </h5>
                            <%End If%>
                        <%Next%>
                    <%End If%>
			    </div>
            <%ElseIf sShowWhat = "ccmeet" Then%>
                <div style="margin: 10px;">
                    <h4 class="h4">Cross-Country Running/Nordic Ski (School-based events) To Date</h4>

                    <table class="table table-striped">
                         <tr>
                            <th>Meets:</th>
                            <td><%=iNumCCMeets%></td>
                            <th>Races:</th>
                            <td><%=iNumCCRaces%></td>
                            <th>Parts:</th>
                            <td><%=iNumCCParts%></td>
                            <th>Finishers:</th>
                            <td><%=iNumCCFinishers%></td>
                            <th>Teams:</th>
                            <td><%=iNumCCTeams%></td>
                            <th>Invoices:</th>
                            <td><%=FormatCurrency(sngCCInvoice)%></td>
                        </tr>
                         <tr>
                            <th>CC Meets:</th>
                            <td><%=iNumCCRunMeets%></td>
                            <th>CC Tms:</th>
                            <td><%=iNumCCRunTeams%></td>
                            <th>CC Parts:</th>
                            <td><%=iNumCCRunParts%></td>
                            <th>Nord Meets:</th>
                            <td><%=iNumNordMeets%></td>
                            <th>Nord Tms:</th>
                            <td><%=iNumNordTeams%></td>
                            <th>Nord Parts:</th>
                            <td><%=iNumNordParts%></td>
                        </tr>
                    </table>

                    <div style="text-align:left;margin: 0;padding: 0;font-size: 0.8em;">
                        <%If sYrlySumm = "y" Then%>
                             <a href="admin.asp?show_what=ccmeet&amp;yrly_summ=n&amp;yrly_brkdwn=<%=sYrlyBrkdwn%>&amp;this_year=<%=iThisYear%>">Hide Year-by-Year Summary</a>
                        <%Else%>
                            <a href="admin.asp?show_what=ccmeet&amp;yrly_summ=y&amp;yrly_brkdwn=<%=sYrlyBrkdwn%>&amp;this_year=<%=iThisYear%>">Show Year-by-Year Summary</a>
                        <%End If%>
                        &nbsp;|&nbsp;
                        <%If sYrlyBrkdwn = "y" Then%>
                             <a href="admin.asp?show_what=ccmeet&amp;yrly_summ=<%=sYrlySumm%>&amp;yrly_brkdwn=n&amp;this_year=<%=iThisYear%>">Hide Yearly Breakdown</a>
                        <%Else%>
                            <a href="admin.asp?show_what=ccmeet&amp;yrly_summ=<%=sYrlySumm%>y&amp;yrly_brkdwn=y&amp;this_year=<%=iThisYear%>">Show Yearly Breakdown</a>
                        <%End If%>
                    </div>

                    <%If sYrlySumm = "y" Then%>
                        <h5 class="h5">Summary by Year</h5>
                        <table class="table table-striped">
                            <tr>
                                <th>Year</th>
                                <th>Meets</th>
                                <th>Races</th>
                                <th>Male</th>
                                <th>Male %</th>
                                <th>Female</th>
                                <th>Female %</th>
                                <th>Total</th>
                                <th>Meet Size</th>
                                <th>Invoice</th>
                                <th>Avg</th>
                            </tr>
                             <%For i = Year(Date) To 2005 Step -1%>
                                <%Call GetYrlySumm(i)%>
                                <tr>
                                    <td><%=YrlySumm(0)%></td>
                                    <td><%=YrlySumm(1)%></td>
                                    <td><%=YrlySumm(2)%></td>
                                    <td><%=YrlySumm(3)%></td>
                                    <td><%=YrlySumm(4)%>%</td>
                                    <td><%=YrlySumm(5)%></td>
                                    <td><%=YrlySumm(6)%>%</td>
                                    <td><%=YrlySumm(7)%></td>
                                    <td><%=YrlySumm(8)%></td>
                                    <td><%=FormatCurrency(YrlySumm(10))%></td>
                                    <%If CLng(YrlySumm(10)) > 0 Then%>
                                        <td><%=FormatCurrency(YrlySumm(10)/YrlySumm(1))%></td>
                                    <%Else%>
                                        <td>n/a</td>
                                    <%End If%>
                                </tr>
                            <%Next%>
                        </table>
                    <%End If%>


                    <%If sYrlyBrkdwn = "y" Then%>
                        <h5 class="h5">Yearly Breakdown</h5>
                        <%For i = Year(Date) To 2002 Step -1%>
                            <%If CInt(iThisYear) = CInt(i) Then%>
                                <%Call GetYrly(iThisYear)%>
                                <h5 class="h5"><a href="admin.asp?show_what=ccmeet&amp;yrly_summ=<%=sYrlySumm%>y&amp;yrly_brkdwn=y"><%=iThisYear%> (click to hide)</a></h5>
                               <table class="table table-striped">
                                    <tr>
                                        <th>No.</th>
                                        <th>Event</th>
                                        <th>Date</th>
                                        <th>Races</th>
                                        <th>Male (%)</th>
                                        <th>Female (%)</th>
                                        <th>Total</th>
                                        <th>Invoice</th>
                                    </tr>
                                    <%For j = 0 To UBound(YrlyData, 2) - 1%>
                                        <tr>
                                            <td><%=j + 1%>)</td>
                                            <td><%=YrlyData(1, j)%></td>
                                            <td><%=YrlyData(2, j)%></td>
                                            <td><%=YrlyData(3, j)%></td>
                                            <td><%=YrlyData(4, j)%></td>
                                            <td><%=YrlyData(5, j)%></td>
                                            <td><%=YrlyData(6, j)%></td>
                                            <td><%=FormatCurrency(YrlyData(8, j))%></td>
                                        </tr>
                                   <%Next%>
                                   <tr>
                                        <th class="totals" colspan="3">Totals:</th>
                                        <th class="totals"><%=YrlyTtls(0)%></th>
                                        <th class="totals"><%=YrlyTtls(1)%></th>
                                        <th class="totals"><%=YrlyTtls(2)%></th>
                                        <th class="totals"><%=YrlyTtls(3)%></th>
                                        <th class="totals"><%=FormatCurrency(YrlyTtls(5))%></th>
                                    </tr>
                                   <tr>
                                        <th class="totals" colspan="3">Percent/Average:</th>
                                        <th class="totals"><%=YrlyPct(0)%></th>
                                        <th class="totals"><%=YrlyPct(1)%>%</th>
                                        <th class="totals"><%=YrlyPct(2)%>%</th>
                                        <th class="totals"><%=YrlyPct(3)%></th>
                                        <th class="totals"><%=FormatCurrency(YrlyPct(4))%></th>
                                    </tr>
                                </table>
                            <%Else%>
                                <h5 class="h5"><a href="admin.asp?show_what=ccmeet&amp;this_year=<%=i%>&amp;yrly_summ=<%=sYrlySumm%>y&amp;yrly_brkdwn=y"><%=i%> (click to view)</a>
                                </h5>
                            <%End If%>
                        <%Next%>
                    <%End If%>
                </div>
            <%Else%>
                <h4 class="h4">All Events (Fitness, Cross-Country Running, and Nordic Ski)</h4>

                <table class="table">
                    <tr>
                        <th>Events:</th>
                        <td><%=iNumEvnts%></td>
                        <th>Races:</th>
                        <td><%=iNumRaces%></td>
                        <th>Parts:</th>
                        <td><%=iNumParts%></td>
                        <th>Invoices:</th>
                        <td><%=FormatCurrency(sngInvoice)%></td>
                    </tr>
                 </table>
            <%End If%>
		</div>
	</div>
</div>	

<!--#include file = "../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
