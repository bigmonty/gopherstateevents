<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, sql2, rs2
Dim lRaceID, lThisRace, lMeetID, lTeamA, lTeamB
Dim i, j, k
Dim sMeetName, dMeetDate, sMeetSite, sWeather, sRaceDist, sOrderResultsBy, sScoreMethod, sRaceName
Dim iNumScore
Dim TmRslts()
Dim fs, fname, sFileName

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetName, MeetDate, MeetSite, Weather FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value
If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
If Not rs(3).Value & "" = "" Then sWeather = Replace(rs(3).Value, "''", "'")
Set rs = Nothing

sql = "SELECT RaceDesc, ScoreMethod, NumScore, RaceDist, RaceUnits, OrderBy FROM Races WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
sScoreMethod = rs(1).Value
iNumScore = rs(2).Value
sRaceDist = rs(3).Value & " " & rs(4).Value
sOrderResultsBy = rs(5).Value
Set rs = Nothing
        
j = 0
ReDim Preserve TmRslts(0)
sql = "SELECT tr.TeamsID FROM TmRslts tr INNER JOIN Teams t ON tr.TeamsID = t.TeamsID WHERE RacesID = " & lRaceID & "ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    TmRslts(j) = rs(0).Value
    j = j + 1
    ReDim Preserve TmRslts(j)
    rs.MoveNext
Loop
Set rs = Nothing
 
Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\dual_reslts_" & sMeetName & "_" & sRaceName & "_" & Year(dMeetDate) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE Dual Meet Results for " & sMeetName & " (" & sRaceName & ")  on " & dMeetDate)
fname.WriteLine("Meet Site: " & sMeetSite)
fname.WriteLine("Distance: " & sRaceDist)
If Not sWeather = vbNullString Then fname.WriteLine("Weather: " & sWeather)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(2)
        
'now get the dual meet breakdown
k = 0
For j = 0 To UBound(TmRslts) - 2
    For k = j + 1 To UBound(TmRslts) - 1
        'get results for these two teams
        lTeamA = TmRslts(j)
        lTeamB = TmRslts(k)
        Call TheseDualRslts(lTeamA, lTeamB)
    Next
Next

Private Sub TheseDualRslts(lTeamA, lTeamB)
    Dim iDualPl, iTeamAPl, iTeamBPl, iThisPl
    Dim sql2, rs2
    Dim TheseRslts(), TempArr(), TmARslts(8), TmBRslts(8)
    Dim x, y, z
    Dim sTeamA, sTeamB
    
    x = 0
    ReDim TheseRslts(1, 0)
    
    'determines the order that the team scores are calculated
    Select Case sScoreMethod
        Case "Points"
            sql2 = "SELECT r.TeamsID, ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
            sql2 = sql2 & "WHERE ir.RacesID = " & lRaceID & " AND ir.Place <> 0 ORDER BY ir.Excludes, ir.Place"
        Case Else
            If sOrderResultsBy = "place" Then
                sql2 = "SELECT r.TeamsID, ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
                sql2 = sql2 & "WHERE ir.RacesID = " & lRaceID & " AND ir.Place <> 0 ORDER BY ir.Excludes, ir.Place"
            Else
                sql2 = "SELECT r.TeamsID, ir.RaceTime FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
                sql2 = sql2 & "WHERE ir.RacesID = " & lRaceID & " AND ir.Place <> 0 ORDER BY ir.Excludes, ir.FnlScnds, ir.Place"
            End If
    End Select

    Set rs2 = conn.Execute(sql2)
    Do While Not rs2.EOF
        If CLng(rs2(0).Value) = CLng(lTeamA) Or CLng(rs2(0).Value) = CLng(lTeamB) Then
            TheseRslts(0, x) = rs2(0).Value
            TheseRslts(1, x) = ConvertToSeconds(rs2(1).Value)
            x = x + 1
            ReDim Preserve TheseRslts(1, x)
        End If
        
        rs2.MoveNext
    Loop
    Set rs2 = Nothing
    
    'team name
    TmARslts(0) = GetTeamName(lTeamA)
    TmBRslts(0) = GetTeamName(lTeamB)
    
    'team score
    TmARslts(1) = "0"
    TmBRslts(1) = "0"
    
    iDualPl = 1
    iTeamAPl = 1
    iTeamBPl = 1
    iThisPl = 1
    
    'now seperate array by team
    For x = 0 To UBound(TheseRslts, 2) - 1
         If CLng(TheseRslts(0, x)) = CLng(lTeamA) Then
            For y = 2 To 8
                If TmARslts(y) = vbNullString Then
                    TmARslts(y) = iThisPl
            
                    If iTeamAPl <= iNumScore And y <= iNumScore + 2 Then
                        Select Case sScoreMethod
                            Case "Points"
                                TmARslts(1) = CSng(TmARslts(1)) + iNumScore * 2 - iDualPl + 1
                            Case "Place"
                                    TmARslts(1) = CSng(TmARslts(1)) + iDualPl
                            Case Else
                                    TmARslts(1) = CSng(TmARslts(1)) + iNumScore * 2 - iDualPl + 1
                        End Select
     
                        iDualPl = iDualPl + 1
                        iTeamAPl = iTeamAPl + 1
                    End If
                   
                    Exit For
                End If
            Next
        Else
            For y = 2 To 8
                If TmBRslts(y) = vbNullString Then
                    TmBRslts(y) = iThisPl
                
                    If iTeamBPl <= iNumScore And y <= iNumScore + 2 Then
                        Select Case sScoreMethod
                            Case "Points"
                                TmBRslts(1) = CSng(TmBRslts(1)) + iNumScore * 2 - iDualPl + 1
                            Case "Place"
                                 TmBRslts(1) = CSng(TmBRslts(1)) + iDualPl
                            Case Else
                                 TmBRslts(1) = CSng(TmBRslts(1)) + iNumScore * 2 - iDualPl + 1
                        End Select
      
                        iDualPl = iDualPl + 1
                        iTeamBPl = iTeamBPl + 1
					End If
                     
                    Exit For
                End If
            Next
        End If
        
        iThisPl = iThisPl + 1
    Next
    
    'sort the array
    ReDim TempArr(8)
    Select Case sScoreMethod
        Case "Points"
            If CSng(TmARslts(1)) < CSng(TmBRslts(1)) Then
                For x = 0 To 8
                    TempArr(x) = TmARslts(x)
                    TmARslts(x) = TmBRslts(x)
                    TmBRslts(x) = TempArr(x)
                Next
            End If
        Case Else
            If CSng(TmARslts(1)) > CSng(TmBRslts(1)) Then
                For x = 0 To 8
                    TempArr(x) = TmARslts(x)
                    TmARslts(x) = TmBRslts(x)
                    TmBRslts(x) = TempArr(x)
                Next
            End If
    End Select
    
    sTeamA = TmARslts(0)
    If Len(sTeamA) < 18 Then
		sTeamA = sTeamA & Space(18 - Len(sTeamA))
    Else
		sTeamA = Left(sTeamA, 18)
    End If
    
    sTeamB = TmBRslts(0)
    If Len(sTeamB) < 18 Then
		sTeamB = sTeamB & Space(18 - Len(sTeamB))
    Else
		sTeamB = Left(sTeamB, 18)
    End If
	
    fname.WriteLine (TmARslts(0) & " vs. " & TmBRslts(0))
    fname.WriteLine ("PL" & vbTab & "TEAM" & Space(14) & vbTab & "PTS" & vbTab & "R1" & _
                 Space(5) & vbTab & "R2" & Space(4) & vbTab & "R3" & Space(5) & vbTab & "R4" & _
                 Space(5) & vbTab & "R5" & Space(5) & vbTab & "R6" & vbTab & "R7")
    fname.WriteLine ("1)" & vbTab & sTeamA & vbTab & TmARslts(1) & vbTab & TmARslts(2) & vbTab & TmARslts(3) &  vbTab & TmARslts(4) & vbTab & TmARslts(5) & vbTab & TmARslts(6) & vbTab & TmARslts(7) & vbTab & TmARslts(8))
    fname.WriteLine ("2)" & vbTab & sTeamB & vbTab & TmBRslts(1) & vbTab & TmBRslts(2) & vbTab & TmBRslts(3) &  vbTab & TmBRslts(4) & vbTab & TmBRslts(5) & vbTab & TmBRslts(6) & vbTab & TmBRslts(7) & vbTab & TmBRslts(8))
      
    fname.WriteBlankLines (1)
End Sub

fname.Close
Set fname=nothing
Set fs=nothing

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<%

Private Function GetTeamName(lTeamID)
    sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
    Set rs2 = conn.Execute(sql2)
    GetTeamName = Replace(rs2(0).Value, "''", "'")
    Set rs2 = Nothing
End Function

'begin download
Response.Redirect "/dwnlds/dual_reslts_" & sMeetName & "_" & sRaceName & "_" & Year(dMeetDate) & ".txt"
%>
<!DOCTYPE html>
<html>
<head>
<title>GSE Download Dual Meet Results</title>
<meta name="description" content="GSE dual-meet formatted results for cross-country running and nordic skiing.">

</head>
<body>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
