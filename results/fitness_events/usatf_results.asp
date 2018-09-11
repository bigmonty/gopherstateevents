<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lEventID, lRaceID
Dim sEventName, sRaceName, sTimingMethod, sChipStart, sShowAge
Dim sngMyTime
Dim dEventDate
Dim iEventType, iRaceType
Dim RsltsArray()
Dim fs, fname, sFileName

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

iEventType = Request.QueryString("event_type")
lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sEventName = rs(0).Value
    dEventDate = rs(1).Value
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Dist, Type FROM RaceData WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sRaceName = rs(0).Value
    iRaceType = rs(1).Value
End If
rs.Close
Set rs = Nothing

i = 0
ReDim RsltsArray(13, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, pr.Age, p.Gender, ir.FnlTime, p.City, p.St, p.Country, p.Email, pr.Bib, "
sql = sql & "ir.FnlScnds, p.DOB FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
sql = sql & "INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.eof
	RsltsArray(0, i) = i + 1								'place
	RsltsArray(1, i) = rs(0).Value   						'part id
	RsltsArray(2, i) = Replace(rs(1).Value, "''", "'") 		'first name
	RsltsArray(3, i) = Replace(rs(2).Value, "''", "'")		'last name
	
	If CLng(lRaceID) = 350 Then
		RsltsArray(4, i) = "na"							'age
	Else
		RsltsArray(4, i) = rs(3).Value							'age
	End If

	RsltsArray(5, i) = rs(4).Value							'gender
	RsltsArray(6, i) = AgeGrp(rs(4).Value, rs(3).Value)		'age grp
	RsltsArray(7, i) = rs(5).Value							'gun time
	RsltsArray(8, i) = ConvertToMinutes(rs(11).Value)		'chip time
	RsltsArray(9, i) = Replace(rs(6).Value, "''", "'")		'city
	RsltsArray(10, i) = rs(7).Value							'state
	RsltsArray(11, i) = rs(8).Value   						'country
	RsltsArray(12, i) = rs(9).Value   						'email
    RsltsArray(13, i) = rs(12).Value   						'dob
	i = i + 1
	ReDim Preserve RsltsArray(13, i)
	rs.MoveNext
Loop
rs.Close
Set rs=Nothing

Private Function AgeGrp(sMF, iMyAge)
	Dim iBegAge, iEndAge
	
	AgeGrp = "na"
	
	iEndAge = 0
	iBegAge = 0
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sMF & "' AND RaceID = " & lRaceID & " ORDER BY EndAge"
	rs2.Open sql2, conn, 1, 2
	Do While Not rs2.EOF
		If CInt(rs2(0).Value) >= CInt(iMyAge) Then
			iEndAge = rs2(0).Value
			Exit Do
		Else
			iBegAge = CInt(rs2(0).Value) + 1
		End If
		rs2.MoveNext
	Loop
	rs2.Close
	Set rs2 = Nothing
	
	If CInt(iBegAge) > 0 Then
		If CInt(iEndAge) = 0 Then
			AgeGrp = iBegAge & " and Over"
		Else
			AgeGrp = iBegAge & " - " & iEndAge
		End If
	End if
End Function

%>
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<%
Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\" & sEventName & "_" & sRaceName & "_" & Year(dEventDate) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE USATF-Formatted Results for " & sEventName & " " & sRaceName & "  on " & dEventDate)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(1)
fname.WriteLine("PL" & vbTab & "FIRST NAME" & vbTab & "LAST NAME" & vbTab & "AGE" & vbTab & "MF" & vbTab & "AGE GRP" & vbTab & "GUN TIME" & vbTab & "CHIP TIME" & vbTab & "CITY" & vbTab & "ST" & vbTab & "CTRY" & vbTab & "EMAIL" & vbTab & "DOB")
For i = 0 to UBound(RsltsArray, 2) - 1
	fname.WriteLine(i + 1 & vbTab & RsltsArray(2, i) & vbTab & RsltsArray(3, i) & vbTab & RsltsArray(4, i) & vbTab & RsltsArray(5, i) & vbTab & RsltsArray(6, i) & vbTab & RsltsArray(7, i) & vbTab & RsltsArray(8, i) & vbTab & RsltsArray(9, i) & vbTab & RsltsArray(10, i) & vbTab & RsltsArray(11, i) & vbTab & RsltsArray(12, i) & vbTab & RsltsArray(13, i))
Next

fname.Close
Set fname=nothing
Set fs=nothing

conn.Close
Set conn = Nothing

'begin download
Response.Redirect "/dwnlds/" & sEventName & "_" & sRaceName & "_" & Year(dEventDate) & ".txt"
%>

<!DOCTYPE html>
<html>
<head>
<title>GSE USATF Results</title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">
<meta name="description" content="GSE USATF-formatted results page.">

<link rel="icon" href="favicon.ico" type="image/x-icon"> 
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon"> 


 



<script language="javascript" src="/misc/scripts.js"></script>


</head>

<body>
</body>
</html>
