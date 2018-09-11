<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim iRaceType
Dim lRaceID
Dim sRaceName, sGender, sMF, sThisLeg, sThisSplit, sThisTrans, sTimingMethod, sChipStart, sOrderBy, sSortRsltsBy, sPartName, sPerMile, sPerKm, sShowAge
Dim i, x, j, n, m, k
Dim sngMyTime
Dim AgeGrps(), iBegAge, IndRslts(), TempArr(6)
Dim lEventID, sEventName, dEventDate
Dim fs, fname, sFileName

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

sGender = Request.QueryString("gender")

If sGender = "M" Then
	sMF = "Male"
Else
	sMF = "Female"
End If

Dim sWhichTime
sWhichTime = Request.QueryString("which_time")
If sWhichTime = vbNullString Then sWhichTime = "chip"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDate, TimingMethod FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sTimingMethod = rs(2).Value
Set rs = Nothing

sql = "SELECT Dist, Type, ChipStart, SortRsltsBy FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
iRaceType = rs(1).Value
sChipStart = rs(2).Value
sSortRsltsBy = rs(3).Value
Set rs = Nothing
	
If sSortRsltsBy = "FnlTime" Then
    sOrderBy = "ir.FnlScnds"
Else
    sOrderBy = "ir.EventPl"
End If

i = 0
iBegAge = 0
ReDim AgeGrps(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EndAge, AgeGrpName FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = '" & LCase(sGender) & "' ORDER BY EndAge"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    AgeGrps(0, i) = iBegAge
    AgeGrps(1, i) = rs(0).Value
    AgeGrps(2, i) = rs(1).Value

    iBegAge = CInt(rs(0).Value) + 1
    i = i + 1
    ReDim Preserve AgeGrps(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetTheseRslts(iBegAge, iEndAge)
	i = 0
	ReDim IndRslts(7, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
    If sWhichTime = "chip" Then
	    sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlScnds, p.City, p.St, p.ParticipantID, pr.Bib, ir.ChipStart "
        sql = sql & "FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
	    sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
        sql = sql & "INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
	    sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND pr.Age >= " & iBegAge & " AND pr.Age <= " & iEndAge 
	    sql = sql & " AND pr.Age <> 99 AND ir.Eligible = 'y' AND ir.FnlScnds > 0 ORDER BY " & sOrderBy
    Else
	    sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlTime, p.City, p.St, p.ParticipantID, pr.Bib, ir.ChipStart "
        sql = sql & "FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
	    sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
        sql = sql & "INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
	    sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND pr.Age >= " & iBegAge & " AND pr.Age <= " & iEndAge 
	    sql = sql & " AND pr.Age <> 99 AND ir.Eligible = 'y' AND ir.FnlScnds > 0 ORDER BY " & sOrderBy
    End If
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
	    If rs(2).Value >= iBegAge Then
	        If rs(2).Value <= iEndAge Then
	            IndRslts(0, i) = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
				
				If CLng(lRaceID) = 350 Then
					IndRslts(1, i) = "na"
				Else
					IndRslts(1, i) = rs(2).Value
				End If

	            If rs(3).Value & "" = "" Then
	                IndRslts(2, i) = "00:00"
	                IndRslts(3, i) = PacePerMile(ConvertToSeconds("00:00"), sRaceName)
	                IndRslts(4, i) = PacePerKM(ConvertToSeconds("00:00"), sRaceName)
	            Else
				    If sWhichTime = "chip" Then
                        IndRslts(2, i) = ConvertToMinutes(CSng(rs(3).Value))
				        IndRslts(3, i) = PacePerMile(rs(3).Value, sRaceName)
				        IndRslts(4, i) = PacePerKM(rs(3).Value, sRaceName)
                    Else
	                    IndRslts(2, i) = rs(3).Value
	                    IndRslts(3, i) = PacePerMile(ConvertToSeconds(rs(3).Value), sRaceName)
	                    IndRslts(4, i) = PacePerKM(ConvertToSeconds(rs(3).Value), sRaceName)
                    End If
	            End If
	            If rs(4).Value & "" = "" Then
	                If rs(5).Value & "" = "" Then
	                    IndRslts(5, i) = "--"
	                Else
	                    IndRslts(5, i) = rs(5).Value
	                End If
	            Else
	                If rs(5).Value & "" = "" Then
	                    IndRslts(5, i) = Replace(rs(3).Value, "''", "'")
	                Else
	                    IndRslts(5, i) = Replace(rs(4).Value, "''", "'") & ", " & rs(5).Value
	               End If
	            End If
				
	            IndRslts(6, i) = rs(6).Value
		        IndRslts(7, i) = rs(7).Value
                                 
	            i = i + 1
	            ReDim Preserve IndRslts(7, i)
	        End If
	    End If
		                     
	    If rs.RecordCount > 0 Then rs.MoveNext
	Loop
End Sub

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/pace_per_mile.asp" -->
<!--#include file = "../../includes/pace_per_km.asp" -->
<%

Private Function GetThisLeg(lThisRace, lThisPart, iMmbrNum)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MmbrName FROM TeamMmbrs WHERE RaceID = " & lThisRace & " AND ParticipantID = " & lThisPart & " AND MmbrNum = "
    sql = sql & iMmbrNum & " ORDER BY MmbrNum"
    rs.Open sql, conn, 1, 2
    GetThisLeg = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

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
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\" & sEventName & "_" & sRaceName & "_Age-Grps_"
sFileName = sFileName & Year(dEventDate) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE Age Group Results for " & sEventName & " " & sRaceName & "-Age Groups  on " & dEventDate)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(2)

For j = 0 to UBound(AgeGrps, 2) - 1
    Call GetTheseRslts(AgeGrps(0, j), AgeGrps(1, j))

 	fname.WriteLine(AgeGrps(2, j))

    fname.WriteLine("PL" & vbTab & " BIB" & vbTab & "NAME" & Space(14) & vbTab & "AGE" & vbTab & "FINAL TIME" & vbTab & "PER MI  " & vbTab & "PER KM  " & vbTab & "LOCATION")

    For i = 0 To UBound(IndRslts, 2) - 1
        'size name field
        If Len(IndRslts(0, i)) < 18 Then
            sPartName = IndRslts(0, i) & Space(18 - Len(IndRslts(0, i)))
        Else
            sPartName = Left(IndRslts(0, i), 18)
        End If

        'size per mile field
        If Len(IndRslts(3, i)) < 8 Then
            sPerMile = IndRslts(3, i) & Space(8 - Len(IndRslts(3, i)))
        Else
            sPerMile = Left(IndRslts(3, i), 8)
        End If

        'side per km field
        If Len(IndRslts(4, i)) < 8 Then
            sPerKm = IndRslts(4, i) & Space(8 - Len(IndRslts(4, i)))
        Else
            sPerKm = Left(IndRslts(4, i), 8)
        End If

        fname.WriteLine(i + 1 & vbTab & IndRslts(7, i) & vbTab & sPartName & vbTab & IndRslts(1, i) & vbTab & IndRslts(2, i) & vbTab & sPerMile & vbTab & sPerKm & vbTab & IndRslts(5, i))
    Next
        
    fname.WriteBlankLines(1)
Next

fname.Close
Set fname=nothing
Set fs=nothing

conn.Close
Set conn = Nothing

'begin download
Response.Redirect "/dwnlds/" & sEventName & "_" & sRaceName & "_Age-Grps_" & Year(dEventDate) & ".txt"
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Download Age Group Results</title>
<meta name="description" content="Gopher State Events (GSE) Age Group Results download.">
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
</body>
</html>
