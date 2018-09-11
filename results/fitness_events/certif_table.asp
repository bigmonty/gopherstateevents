<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID, lRaceID, lPartID
Dim sEventName, sRaceName, sMyTime, sMyName, sLogo
Dim i, iMyPlace, iBib
Dim dEventDate

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

sql = "SELECT EventName, EventDate, Logo FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sLogo = rs(2).Value
Set rs = Nothing

If sLogo & "" = "" Then 
    sLogo = "/graphics/gopher.jpg"
Else
    sLogo = "/events/logos/" & sLogo
End If

sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
Set rs = Nothing

sql = "SELECT ParticipantID FROM PartRace WHERE RaceID = " & lRaceID & " AND Bib = " & iBib
Set rs = conn.Execute(sql)
lPartID = rs(0).Value
Set rs = Nothing

sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
Set rs = conn.Execute(sql)
sMyName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
Set rs = Nothing

iMyPlace = 1
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ParticipantID, FnlScnds FROM IndResults WHERE RaceID = " & lRaceID & " ORDER BY FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If CLng(rs(0).Value) = CLng(lPartID) Then
        sMyTime = ConvertToMInutes(rs(1).Value)
        Exit Do
    Else
        iMyPlace = CInt(iMyPlace) + 1
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If CSng(sglScnds) <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(CSng(sglScnds) \ 60)
    sSecondPart = CStr(((CSng(sglScnds) / 60) - (CSng(sglScnds) \ 60)) * 60)
    
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
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finishers Certificate</title>
<meta name="description" content="Gopher State Events (GSE) Finisher Certificate.">
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    @import url(https://fonts.googleapis.com/css?family=Libre+Baskerville:400,400italic);

.fancyTxt {
  font-family: 'Libre Baskerville', serif;
  font-style: italic;
  font-size: 25px
}
.small {
  font-size: 20px
}
</style>
</head>
<body>
<table style="border-collapse:collapse;width: 1030px;height:775px;background-image: url('/graphics/finisher_certificate.png');background-repeat: no-repeat;">
    <tr>
        <td style="text-align: center;padding: 300px 20px 0 0;height: 300px;" colspan="2">
            <span class="fancyTxt">This is to certify that <%=sMyName%> finished the
            <br>
            <%=sEventName%> - <%=sRaceName%> on <%=dEventDate%>
            <br><br>
            <span class="fancyTxt small">Time:&nbsp;<%=sMyTime%></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <span class="fancyTxt small">Overall Place:&nbsp;<%=iMyPlace%></span></span>
        </td>
    </tr>
    <tr>
        <td style="width: 50%;">&nbsp;</td>
        <td style="text-align: right;padding:0 100px 0 0;"><img src="<%=sLogo%>" alt="logo" style="height: 85px;margin: 0;"></td>
    </tr>
</table>

<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
