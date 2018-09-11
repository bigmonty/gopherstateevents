<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i,m, n, x
Dim lEventID, lRaceID
Dim sRaceName, sGender, sMF, sThisDist, sThisLeg, sThisSplit, sThisTrans, sTimingMethod, sChipStart, sEventClass, sDist, sAllowDuplAwds
Dim sShowAge, sEventName, sSortRsltsBy, sEventRaces
Dim iEventType, iRaceType
Dim sngMyTime
Dim IndRslts, TempArr(9)
Dim fs, fname, sFileName
Dim sDwnldName, sDwnldCity, sDwnldPerMi, sDwnldPerKM
Dim dEventDate

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

iEventType = Request.QueryString("event_type")
lRaceID = Request.QueryString("race_id")
sGender = Request.QueryString("gender")

If sGender = "M" Then
	sMF = "Male"
ElseIf sGender = "F" Then
	sMF = "Female"
Else
	sMF = "Both"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDate, TimingMethod, EventClass FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sTimingMethod = rs(2).Value
sEventClass = rs(3).Value
Set rs = Nothing

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

sql = "SELECT Dist, RaceName, Type, AllowDuplAwds, ChipStart, SortRsltsBy, ShowAge FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sDist = rs(0).Value
sRaceName = rs(1).Value
iRaceType = rs(2).Value
sAllowDuplAwds = rs(3).Value
sChipStart = rs(4).Value
sSortRsltsBy = rs(5).Value
sShowAge = rs(6).Value
Set rs = Nothing

If sGender = "B" Then
    If sSortRsltsBy = "place" Then
        Set rs = Server.CreateObject("ADODB.Recordset")  
        sql = "Exec ResultsByPlace @RaceID = " & lRaceID
        Set rs = conn.execute(sql) 
        If rs.BOF and rs.EOF Then
            ReDim IndRslts(10, 0)
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
            ReDim IndRslts(10, 0)
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
            ReDim IndRslts(10, 0)
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
            ReDim IndRslts(10, 0)
        Else
            IndRslts = rs.GetRows()
        End If
        rs.Close
        Set rs = Nothing
    End If
End If

For i = 0 To UBound(IndRslts, 2)
    If sShowAge = "n" Then
        IndRslts(4, i) = MyAgeGrp(IndRslts(0, i))
    Else
		If IndRslts(4, i) = "99" Then IndRslts(4, i) = "0"
    End If
Next

Private Function MyAgeGrp(iMyBib)
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
End Function

Private Function GetThisLeg(lThisRace, lThisPart, iMmbrNum)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MmbrName FROM TeamMmbrs WHERE RaceID = " & lThisRace & " AND ParticipantID = " & lThisPart & " AND MmbrNum = "
    sql = sql & iMmbrNum & " ORDER BY MmbrNum"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetThisLeg = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

Private Function GetDOB(lThisPart)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT DOB FROM Participant WHERE ParticipantID = " & lThisPart
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetDOB = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/pace_per_mile.asp" -->
<!--#include file = "../../includes/pace_per_km.asp" -->
<%

Private Function GetThisSplit(lThisRace, lThisPart, iSplitNum)
    Set rs = Server.CreateObject("ADODB.Recordset")
    Select Case iSplitNum
        Case 1
            sql = "SELECT rd.RaceDelay, pr.IndDelay, pr.Trans1Out FROM IndResults ir INNER JOIN RaceData rd ON rd.RaceID = ir.RaceID "
            sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID WHERE pr.RaceID = " & lThisRace
            sql = sql & " AND pr.ParticipantID = " & lThisPart
        Case 2
            sql = "SELECT pr.Trans1Out, pr.Trans2Out FROM IndResults ir INNER JOIN RaceData rd ON rd.RaceID = ir.RaceID "
            sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID WHERE pr.RaceID = " & lThisRace
            sql = sql & " AND pr.ParticipantID = " & lThisPart
        Case 3
            sql = "SELECT pr.Trans2Out, ir.ElpsdTime FROM IndResults ir INNER JOIN RaceData rd ON rd.RaceID = ir.RaceID "
            sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID WHERE pr.RaceID = " & lThisRace
            sql = sql & " AND pr.ParticipantID = " & lThisPart
    End Select
    rs.Open sql, conn, 1, 2
    Select Case iSplitNum
        Case 1
            If ConvertToSeconds(rs(2).Value) = 0 Then
                GetThisSplit = "unavail"
            Else
                GetThisSplit = ConvertToMinutes(ConvertToSeconds(rs(2).Value) - ConvertToSeconds(rs(1).Value) - ConvertToSeconds(rs(0).Value))
            End If
        Case Else
            If ConvertToSeconds(rs(0).Value) = 0 Or ConvertToSeconds(rs(1).Value) = 0 Then
                GetThisSplit = "unavail"
            Else
                GetThisSplit = ConvertToMinutes(ConvertToSeconds(rs(1).Value) - ConvertToSeconds(rs(0).Value))
            End If
    End Select
    rs.Close
    Set rs = Nothing
End Function

Private Function GetThisTrans(lThisRace, lThisPart, iTransNum)
    Set rs = Server.CreateObject("ADODB.Recordset")
    Select Case iTransNum
        Case 1
            sql = "SELECT pr.Trans1In, pr.Trans1Out FROM IndResults ir INNER JOIN RaceData rd "
            sql = sql & "ON rd.RaceID = ir.RaceID INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID "
            sql = sql & "WHERE pr.RaceID = " & lThisRace & " AND pr.ParticipantID = " & lThisPart
        Case 2
            sql = "SELECT pr.Trans2In, pr.Trans2Out FROM IndResults ir INNER JOIN RaceData rd "
            sql = sql & "ON rd.RaceID = ir.RaceID INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID "
            sql = sql & "WHERE pr.RaceID = " & lThisRace & " AND pr.ParticipantID = " & lThisPart
    End Select
    rs.Open sql, conn, 1, 2
    If ConvertToSeconds(rs(0).Value) = 0 Or ConvertToSeconds(rs(1).Value) = 0 Then
        GetThisTrans = "unavail"
    Else
        GetThisTrans = ConvertToMinutes(ConvertToSeconds(rs(1).Value) - ConvertToSeconds(rs(0).Value))
    End If
    rs.Close
    Set rs = Nothing
End Function

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\" & sEventName & "_" & sRaceName & "-" & sMF & "_"
sFileName = sFileName & Year(dEventDate) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE Race Results for " & sEventName & " " & sRaceName & "-" & sMF & "  on " & dEventDate)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
If sShowAge = "y" Then
    fname.WriteLine("PL" & vbTab & " BIB" & vbTab & "NAME" & Space(14) & vbTab & "AGE" & vbTab & "MF" & vbTab & "CHIP  " & vbTab & "GUN  " & vbTab & "START " & vbTab & "CITY" & Space(14) & vbTab & "ST")
Else
    fname.WriteLine("PL" & vbTab & " BIB" & vbTab & "NAME" & Space(14) & vbTab & "AGE GRP" & vbTab & "MF" & vbTab & "CHIP  " & vbTab & "GUN  " & vbTab & "START " & vbTab & "CITY" & Space(14) & vbTab & "ST")
End If

For i = 0 to UBound(IndRslts, 2)
	sDwnldName = IndRslts(2, i) & " " & IndRslts(1, i)
	If Len(sDwnldName) < 18 Then
		sDwnldName = sDwnldName & Space(18 - Len(sDwnldName))
	Else
		sDwnldName = Left(sDwnldName, 18)
	End If
	
'	sDwnldPerMi = IndRslts(6, i)
'	If Len(sDwnldPerMi) < 8 Then
'		sDwnldPerMi = sDwnldPerMi & Space(8 - Len(sDwnldPerMi))
'	Else
'		sDwnldPerMi = Left(sDwnldPerMi, 8)
'	End If
	
	sDwnldPerKM = IndRslts(7, i)
	If Len(sDwnldPerKM) < 8 Then
		sDwnldPerKM = sDwnldPerKM & Space(8 - Len(sDwnldPerKM))
	Else
		sDwnldPerKM = Left(sDwnldPerKM, 8)
	End If
	
	sDwnldCity = IndRslts(8, i)
	If Len(sDwnldCity) < 18 Then
		sDwnldCity = sDwnldCity & Space(18 - Len(sDwnldCity))
	Else
		sDwnldCity = Left(sDwnldCity, 18)
	End If
	fname.WriteLine(i + 1 & vbTab & IndRslts(0, i) & vbTab & sDwnldName & vbTab & IndRslts(4, i) & vbTab & IndRslts(3, i) & vbTab & IndRslts(5, i) & vbTab & IndRslts(6, i) & vbTab & sDwnldPerKM & vbTab & sDwnldCity & vbTab & IndRslts(9, i))

	If iRaceType = 10 Then
	    For x = 0 To 2
	        sThisLeg = GetThisLeg(lRaceID, IndRslts(8, i), x + 1)
	        sThisSplit = GetThisSplit(lRaceID, IndRslts(8, i), x + 1)
	        If Not x = 2 Then sThisTrans = GetThisTrans(lRaceID, IndRslts(8, i), x + 1)
	        
	        If x = 2 Then
	            fname.writeline vbTab & vbTab & sThisLeg & vbTab & sThisSplit
	        Else
	            fname.writeline vbTab & vbTab & sThisLeg & vbTab & sThisSplit & " (Trans: " & sThisTrans & ")"
	        End If
	    Next
	End If
Next

fname.Close
Set fname=nothing
Set fs=nothing

conn.Close
Set conn = Nothing

'begin download
Response.Redirect "/dwnlds/" & sEventName & "_" & sRaceName & "-" & sMF & "_" & Year(dEventDate) & ".txt"
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Download Results</title>
<meta name="description" content="Gopher State Events (GSE) results download.">
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
</body>
</html>
