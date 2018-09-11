<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lRaceID, lEventID, lRaceSplitsID
Dim sRaceName, sGender, sMF, sEventName
Dim iNumSplits
Dim dEventDate
Dim SplitRslts(), SortArr(2), LegNames(3)

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
sGender = Request.QueryString("gender")

If sGender = "M" Then
	sMF = "Male"
Else
	sMF = "Female"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = rs(0).Value
dEventDate = rs(1).Value
Set rs = Nothing

sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
Set rs = Nothing
    
Set rs = Server.CreateObject("ADODB.REcordset")
sql = "SELECT RaceSplitsID FROM RaceSplits WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
lRaceSplitsID = rs(0).Value
rs.Close
Set rs = Nothing
    
'get num splits
Set rs = Server.CreateObject("ADODB.REcordset")
sql = "SELECT NumSplits FROM RaceData WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
iNumSplits = rs(0).Value
rs.Close
Set rs = Nothing
        
'get leg headers ready for display
LegNames(0) = "Leg 1"
LegNames(1) = "Leg 2"
LegNames(2) = "Leg 3"
LegNames(3) = "Leg 4"
        
Set rs = Server.CreateObject("ADODB.REcordset")
sql = "SELECT Leg1Name, Leg2Name, Leg3Name, Leg4Name FROM RaceSplits WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    If Not rs(0).Value & "" = "" Then LegNames(0) = Replace(rs(0).Value, "''", "'")
    If Not rs(1).Value & "" = "" Then LegNames(1) = Replace(rs(1).Value, "''", "'")
    If Not rs(2).Value & "" = "" Then LegNames(2) = Replace(rs(2).Value, "''", "'")
    If Not rs(3).Value & "" = "" Then LegNames(3) = Replace(rs(3).Value, "''", "'")
End If
rs.Close
Set rs = Nothing

%>
<!--#include file = "../../../includes/convert_to_seconds.asp" -->
<%            
'get participants
i = 0
ReDim SplitRslts(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT pr.Bib, p.FirstName, p.LastName FROM Participant p INNER JOIN IndResults ir "
sql = sql & "ON p.ParticipantID = ir.ParticipantID INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND pr.RaceID = " & lRaceID & " AND p.Gender = '" & sGender
sql = sql & "' ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    SplitRslts(0, i) = rs(0).Value
    SplitRslts(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    i = i + 1
    ReDim Preserve SplitRslts(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetSplitRslts(iThisSplit)
    Dim x, y, z
     
    For x = 0 To UBound(SplitRslts, 2) - 1
        SplitRslts(2, x) = "99:99:99.999"
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Leg" & iThisSplit & "Time FROM SplitTimes WHERE Bib = " & SplitRslts(0, x) & " AND RaceSplitsID = " & lRaceSplitsID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then SplitRslts(2, x) = rs(0).Value
        rs.Close
        Set rs = Nothing
    Next
               
    For x = 0 To UBound(SplitRslts, 2) - 2
        For y = x + 1 To UBound(SplitRslts, 2) - 1
            If CSng(ConvertToSeconds(SplitRslts(2, x))) > CSng(ConvertToSeconds(SplitRslts(2, y))) Then
                For z = 0 To 2
                    SortArr(z) = SplitRslts(z, x)
                    SplitRslts(z, x) = SplitRslts(z, y)
                    SplitRslts(z, y) = SortArr(z)
                Next
            End If
        Next
    Next
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Rank By Split</title>
<meta name="description" content="Gopher State Events (GSE) Results ranked by split.">
<!--#include file = "../../../includes/js.asp" --> 

<style type="text/css">
<!--
th{
	padding:0 5px 0 5px;
	}

td{
	padding:0 5px 0 5px;
    text-align: left;
	}
-->
</style>
</head>

<body>
<img src="/graphics/html_header.png" alt="Results">
<div style="text-align:right;margin-right:10px;">
	<a href="javascript:window.print()">Print This</a>
</div>
<h1 class="h1">Rank By Split for <%=sEventName%></h1>
<h2 class="h2"><%=sRaceName%>  (<%=sMF%>) <%=dEventDate%></h2>

    <%For j = 1 To 4%>
        <%If j > iNumSplits + 1 Then Exit For%>

        <%Call GetSplitRslts(j)%>
        <h4 class="h4"><%=LegNames(j - 1)%></h4>
        <table class="table-striped">
            <tr>
                <th>Rank</th><th>Bib</th><th style="text-align: left;">Participant</th><th>Split Time</th>
            </tr>
            <%For i = 0 To UBound(SplitRslts, 2) - 1%>
                <tr>
                    <%If i mod 2 = 0 Then%>
                        <td class="alt"><%=i + 1%>)</td>
                        <td class="alt" style="text-align: left;"><%=SplitRslts(0, i)%></td>
                        <td class="alt" style="text-align: left;"><%=SplitRslts(1, i)%></td>
                        <td class="alt" style="text-align: left;">
                            <%If SplitRslts(2, i) ="99:99:99.999" Then%>
                                Missing
                            <%Else%>
                                <%=SplitRslts(2, i)%>
                            <%End If%>
                        </td>
                    <%Else%>
                        <td><%=i + 1%>)</td>
                        <td style="text-align: left;"><%=SplitRslts(0, i)%></td>
                        <td style="text-align: left;"><%=SplitRslts(1, i)%></td>
                        <td style="text-align: left;">
                            <%If SplitRslts(2, i) ="99:99:99.999" Then%>
                                Missing
                            <%Else%>
                                <%=SplitRslts(2, i)%>
                            <%End If%>
                        </td>
                    <%End If%>
                </tr>
            <%Next%>
        </table>
     <%Next%>               
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
