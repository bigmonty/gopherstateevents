<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID, lRaceID, lSuppLegID
Dim i,m, n, x
Dim sRaceName, sGender, sMF, sThisDist, sLegName, sOtherName, sChipStart, sDist, sEventName, sDwnldName, sSuppTime, sOtherTime, sShowAge
Dim sngMyTime, sngRaceDelay
Dim dEventDate
Dim IndRslts(), TempArr(9)
Dim fs, fname, sFileName

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

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

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SuppLegID FROM SuppLeg WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then lSuppLegID = rs(0).Value
rs.Close
Set rs = Nothing

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
Set rs = Nothing

sql = "SELECT Dist, RaceName, ChipStart, RaceDelay FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sDist = rs(0).Value
sRaceName = rs(1).Value
sChipStart = rs(2).Value
sngRaceDelay = ConvertToSeconds(rs(3).Value)
Set rs = Nothing

i = 0
ReDim IndRslts(5, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.ChipTime, ir.ChipStart, p.City, p.St, pr.Bib "
sql = sql & "FROM Participant p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID INNER JOIN PartRace pr "
sql = sql & "ON pr.ParticipantID = p.ParticipantID INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID "
sql = sql & "AND rd.RaceID = ir.RaceID WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender
sql = sql & "' AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    IndRslts(0, i) = rs(7).Value & "-" & Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
    IndRslts(1, i) = rs(2).Value
    IndRslts(2, i) = rs(3).Value
    IndRslts(3, i) = rs(4).Value
    If rs(5).Value & "" = "" Or rs(5).Value = Null Then
        If rs(6).Value & "" = "" Or rs(6).Value = Null Then
            IndRslts(4, i) = "--"
        Else
            IndRslts(4, i) = rs(6).Value
        End If
    Else
        If rs(6).Value & "" = "" Or rs(6).Value = Null Then
            IndRslts(4, i) = Replace(rs(5).Value, "''", "'")
        Else
            IndRslts(4, i) = Replace(rs(5).Value, "''", "'") & ", " & rs(6).Value
        End If
    End If
                    
    IndRslts(5, i) = rs(7).Value
                    
    i = i + 1
    ReDim Preserve IndRslts(5, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT LegName, OtherName FROM SuppLeg WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sLegName = Replace(rs(0).Value, "''", "'")
If Not rs(1).Value & "" = "" Then sOtherName = Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<%
Private Sub GetSplits(iBib, sChipTime)
    Dim sngChipTime

    sSuppTime = "00:00.000"
    sOtherTime = "00:00:00.000"
    
    sngChipTime = ConvertToSeconds(sChipTime)

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SuppTime, OtherTime FROM SuppLegRslts WHERE SuppLegID = " & lSuppLegID & " AND Bib = " & iBib
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sSuppTime = rs(0).Value
        sOtherTime = ConvertToMinutes(sngChipTime - ConvertToSeconds(rs(0).Value))
    End If
    rs.Close
    Set rs = Nothing
End Sub

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\" & sEventName & "_" & sRaceName & "-" & sMF & "_"
sFileName = sFileName & Year(dEventDate) & "_splits.txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE  Results with Splits for " & sEventName & " " & sRaceName & "-" & sMF & "  on " & dEventDate)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("PL" & vbTab & "BIB-NAME" & Space(20) & vbTab & "AGE" & vbTab & "TOTAL TIME" & vbTab & "START" & vbTab & vbTab & "FROM")

For i = 0 to UBound(IndRslts, 2) - 1
	sDwnldName = IndRslts(5, i) & "-" & IndRslts(1, i) & ", " & IndRslts(0, i)
	If Len(sDwnldName) < 28 Then
		sDwnldName = sDwnldName & Space(28 - Len(sDwnldName))
	Else
		sDwnldName = Left(sDwnldName, 28)
	End If
	
	fname.WriteLine(i + 1 & vbTab & sDwnldName & vbTab & IndRslts(1, i) & vbTab & IndRslts(2, i) & vbTab & IndRslts(3, i) & vbTab & vbTab & IndRslts(4, i))

    Call GetSplits(IndRslts(5, i), IndRslts(2, i))

	fname.writeline vbTab & vbTab & sLegName & ": " & sSuppTime
	fname.writeline vbTab & vbTab & sOtherName & ": " & sOtherTime
    fname.writeblanklines(1)
Next

fname.Close
Set fname=nothing
Set fs=nothing

conn.Close
Set conn = Nothing

'begin download
Response.Redirect "/dwnlds/" & sEventName & "_" & sRaceName & "-" & sMF & "_" & Year(dEventDate) & "_splits.txt"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events (GSE) Download Results with Supplemental Leg Splits</title>
<meta name="description" content="Gopher State Events (GSE) Download Results with Supplemental Leg Splits">
<!--#include file = "../../includes/js.asp" --> 
</head>

<body>
</body>
</html>
