<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lRaceID, lEventID, lRaceSplitsID
Dim sRaceName, sGender, sMF, sEventName
Dim iNumSplits
Dim dEventDate
Dim sThisLeg, sMyName
Dim SplitRslts(), SortArr(2), LegNames(3), LegTimes(3)

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
            
'get splits
i = 0
ReDim SplitRslts(2 + iNumSplits, 0)
Set rs = Server.CreateObject("ADODB.REcordset")
sql = "SELECT pr.Bib, p.FirstName, p.LastName FROM Participant p INNER JOIN IndResults ir "
sql = sql & "ON p.ParticipantID = ir.ParticipantID INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND pr.RaceID = " & lRaceID & " AND p.Gender = '" & sGender
sql = sql & "' ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Call GetLegTimes(rs(0).Value, lRaceID)
    SplitRslts(0, i) = rs(0).Value
    SplitRslts(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    For j = 0 To iNumSplits
        SplitRslts(2 + j, i) = LegTimes(j)
    Next
    i = i + 1
    ReDim Preserve SplitRslts(2 + iNumSplits, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetLegTimes(iThisBib, lThisRace)
    Dim rs2, sql2
    
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Leg1Time, Leg2Time, Leg3Time, Leg4Time FROM SplitTimes st INNER JOIN RaceSplits rs ON st.RaceSplitsID = rs.RaceSplitsID "
    sql2 = sql2 & "WHERE rs.RaceID = " & lThisRace & " AND st.Bib = " & iThisBib
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then
        LegTimes(0) = rs2(0).Value
        LegTimes(1) = rs2(1).Value
        LegTimes(2) = rs2(2).Value
        LegTimes(3) = rs2(3).Value
    End If
    rs2.Close
    Set rs2 = Nothing
    
    If Len(LegTimes(0)) < 12 Then LegTimes(0) = LegTimes(0) & Space(12 - Len(LegTimes(0)))
    If Len(LegTimes(1)) < 12 Then LegTimes(1) = LegTimes(1) & Space(12 - Len(LegTimes(1)))
    If Len(LegTimes(2)) < 12 Then LegTimes(2) = LegTimes(2) & Space(12 - Len(LegTimes(2)))
    If Len(LegTimes(3)) < 12 Then LegTimes(3) = LegTimes(3) & Space(12 - Len(LegTimes(3)))
End Sub

Private Function MyFinalTime(iThisBib)
    Dim sngMyFnlTime
    Dim sngMySplitSum

    MyFinalTime = "00:00:00.000"
    sngMyFnlTime = "0"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.FnlScnds FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID WHERE pr.RaceID = " & lRaceID 
    sql = sql & " AND pr.Bib = " & iThisBib
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sngMyFnlTime = rs(0).Value
    rs.Close
    Set rs = Nothing
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Leg1Time, Leg2Time, Leg3Time, Leg4Time FROM SplitTimes st INNER JOIN RaceSplits rs ON st.RaceSplitsID = rs.RaceSplitsID "
    sql = sql & "WHERE rs.RaceID = " & lRaceID & " AND st.Bib = " & iThisBib
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sngMySplitSum = ConvertToSeconds(rs(0).Value) + ConvertToSeconds(rs(1).Value) + ConvertToSeconds(rs(2).Value) + ConvertToSeconds(rs(3).Value)
    End If
    rs.Close
    Set rs = Nothing

    If ABS(CSng(sngMyFnlTime) - CSng(sngMySplitSum)) > 0.5 Then
        MyFinalTime = ConvertToMinutes(sngMySplitSum)
    Else
        MyFinalTime = ConvertToMinutes(sngMyFnlTime)
    End If
End Function

%>
<!--#include file = "../../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../../includes/convert_to_minutes.asp" -->
<%            

%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Results With Splits</title>
<meta name="description" content="Gopher State Events (GSE) Results w/Splits.">
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
<h1 class="h1">Results With Splits for <%=sEventName%></h1>
<h2 class="h2"><%=sRaceName%>  (<%=sMF%>) <%=dEventDate%></h2>
                    
<table class="table-striped">
    <tr>
        <th>Pl</th><th>Bib</th><th style="text-align: left;">Participant</th>
        <%For i = 0 to CInt(iNumSplits)%>
            <th><%=LegNames(i)%></th>
        <%Next%>
        <th>Final Time</th>
    </tr>
    <%For i = 0 To UBound(SplitRslts, 2) - 1%>
        <tr>
            <%If i mod 2 = 0 Then%>
                <td class="alt"><%=i + 1%>)</td>
                <td class="alt" style="text-align: left;"><%=SplitRslts(0, i)%></td>
                <%For j = 1 To 2 + iNumSplits%>
                    <td class="alt"><%=SplitRslts(j, i)%></td>
                <%Next%>
                <td class="alt"><%=MyFinalTime(SplitRslts(0, i))%></td>
            <%Else%>
                <td><%=i + 1%>)</td>
                <td style="text-align: left;"><%=SplitRslts(0, i)%></td>
                <%For j = 1 To 2 + iNumSplits%>
                    <td><%=SplitRslts(j, i)%></td>
                <%Next%>
                <td><%=MyFinalTime(SplitRslts(0, i))%></td>
            <%End If%>
        </tr>
    <%Next%>
</table>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
