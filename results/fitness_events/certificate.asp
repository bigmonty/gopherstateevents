<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID, lRaceID, lPartID
Dim iBib, iYear
Dim sEventName

lEventID = Request.QueryString("event_id")
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

iBib = Request.QueryString("bib")
If Not IsNumeric(iBib) Then Response.Redirect("http://www.google.com")
If CLng(iBib) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
iYear = Year(rs(1).Value)
Set rs = Nothing

sql = "SELECT ParticipantID FROM PartRace WHERE RaceID = " & lRaceID & " AND Bib = " & iBib
Set rs = conn.Execute(sql)
lPartID = rs(0).Value
Set rs = Nothing

conn.Close
Set conn = Nothing

'delete the file if it exists
Dim fs
Dim MyFile
Set fs = Server.CreateObject("Scripting.FileSystemObject")
MyFile = Server.mappath("/results/fitness_events/certificates/" & sEventName & "_" & iYear & "_" & lPartID & ".pdf")
If fs.FileExists(MyFile) Then fs.DeleteFile(MyFile) 
Set fs = Nothing

Dim Pdf, Doc, FileName
Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument
Doc.ImportFromUrl "http://www.gopherstateevents.com/results/fitness_events/certif_table.asp?race_id=" & lRaceID & "&event_id=" & lEventID & "&bib=" & iBib, "scale=0.8; hyperlinks=true; drawbackground=true; landscape=true"
Filename = Doc.Save("C:\inetpub\h51web\gopherstateevents\results\fitness_events\certificates\" & sEventName & "_" & iYear & "_" & lPartID & ".pdf", False )
Response.ContentType = "application/pdf"

Response.AddHeader "content-disposition", "Filename=" & Request.QueryString("File") & ".PDF"
Const adTypeBinary = 1

Dim strFilePath
Dim objStream

strFilePath = "C:\inetpub\h51web\gopherstateevents\results\fitness_events\certificates\" & sEventName & "_" & iYear & "_" & lPartID & ".pdf" 

Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type = adTypeBinary
objStream.LoadFromFile strFilePath

Response.BinaryWrite objStream.Read

objStream.Close
Set objStream = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Create GSE Finishers Certificate</title>
<meta name="description" content="Create Gopher State Events (GSE) Finisher Certificate.">
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<a href="/results/fitness_events/certificates/<%=sEventName%>_<%=iYear%>_<%=lPartID%>.pdf">Open Certificate</a>
</body>
</html>
