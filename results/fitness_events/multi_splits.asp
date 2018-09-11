<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k, m
Dim lRaceID, lEventID
Dim sEventName, sGender, sRaceName, sLeg1Name, sLeg2Name, sLeg3Name, sStartTime, sRaceDelay, sShowAge, sMmbrName, sMmbrDist, sLegName, sLegDist, sMmbrGender
Dim sSortBy
Dim iNumLegs, iMmbrAge
Dim sngMyStart, sngLeg1Dist, sngLeg2Dist, sngLeg3Dist
Dim dEventDate
Dim Races, MultiRslts(), SortArr(16)

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")

sSortBy = Request.QueryString("sort_by")
If sSortBy = vbNullString Then sSortBy = "elapsed"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
        
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT StartTime FROM RFIDSettings Where EventID = " & lEventID
rs.Open sql, conn, 1, 2
sStartTime = rs(0).Value
rs.Close
Set rs = Nothing
        
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ShowAge FROM RaceData WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
sShowAge = rs(0).Value
rs.Close
Set rs = Nothing

'get num legs
iNumLegs = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT NumLegs FROM MultiSettingsChip WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iNumLegs = rs(0).Value
rs.Close
Set rs = Nothing
    
'get let names
sLeg3Name = vbNullString
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Leg1Name, Leg2Name, Leg3Name FROM MultiSettingsChip WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sLeg1Name = Replace(rs(0).Value, "''", "'")
    sLeg2Name = Replace(rs(1).Value, "''", "'")
    If iNumLegs > 2 Then sLeg3Name = Replace(rs(2).Value, "''", "'")
End If
rs.Close
Set rs = Nothing

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

'get races
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Races = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
End If

sql = "SELECT RaceName, RaceDelay FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
sRaceDelay = rs(1).Value
Set rs = Nothing

iNumLegs = 3

sLeg1Name = "Swim"
sLeg2Name = "Bike"
sLeg3Name = "Run"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT NumLegs, Leg1Name, Leg2Name, Leg3Name, Leg1Dist, Leg2Dist, Leg3Dist FROM MultiSettingsChip WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then 
    iNumLegs = rs(0).Value

    If Not rs(1).Value & "" = "" Then sLeg1Name = rs(1).Value
    If Not rs(2).Value & "" = "" Then sLeg2Name = rs(2).Value
    If Not rs(3).Value & "" = "" Then sLeg3Name = rs(3).Value

    sngLeg1Dist = rs(4).Value
    sngLeg2Dist = rs(5).Value
    sngLeg3Dist = rs(6).Value
End If
rs.Close
Set rs = Nothing

If sngLeg1Dist & "" = "" Then sngLeg1Dist = "unknown"
If sngLeg2Dist & "" = "" Then sngLeg2Dist = "unknown"
If sngLeg3Dist & "" = "" Then sngLeg3Dist = "unknown"

'get results by gender
Private Sub RsltsByMF(sThisMF)
    Dim x, y, z

    x = 0
    ReDim MultiRslts(16, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If sThisMF = "X" Then
        sql = "SELECT pr.Bib, p.FirstName, p.LastName, p.Gender, pr.Age, ir.FnlScnds, p.ParticipantID, ir.ElpsdTime, pr.IndDelay, p.TeamInd "
        sql = sql & "FROM Participant p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
        sql = sql & "INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID WHERE ir.RaceID = " & lRaceID
        sql = sql & " AND pr.RaceID = " & lRaceID & " AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
    Else
        sql = "SELECT pr.Bib, p.FirstName, p.LastName, p.Gender, pr.Age, ir.FnlScnds, p.ParticipantID, ir.ElpsdTime, pr.IndDelay, p.TeamInd "
        sql = sql & "FROM Participant p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
        sql = sql & "INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID WHERE ir.RaceID = " & lRaceID
        sql = sql & " AND pr.RaceID = " & lRaceID & " AND ir.FnlScnds > 0 AND p.Gender = '" & sThisMF & "' ORDER BY ir.FnlScnds"
    End If
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MultiRslts(0, x) = rs(0).Value & "-" & Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
        MultiRslts(1, x) = rs(3).Value
        MultiRslts(2, x) = rs(4).Value
        If MultiRslts(2, x) = "99" Then MultiRslts(2, x) = "--"
        MultiRslts(8, x) = rs(5).Value
        MultiRslts(9, x) = rs(0).Value
        MultiRslts(10, x) = rs(8).Value
        MultiRslts(15, x) = rs(6).Value
        MultiRslts(16, x) = rs(9).Value
        x = x + 1
        ReDim Preserve MultiRslts(16, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For x = 0 To UBound(MultiRslts, 2) - 1
        sngMyStart = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ActualStart FROM StartData WHERE EventID = " & lEventID & " AND Bib = " & MultiRslts(9, x)
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then sngMyStart = rs(0).Value
        rs.Close
        Set rs = Nothing

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Trans1In, Trans1Out, Trans1Time, Trans2In, Trans2Out, Trans2Time FROM TransData WHERE Bib = "
        sql = sql & MultiRslts(9, x) & " AND RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            MultiRslts(3, x) = ConvertToMinutes(Round(ConvertToSeconds(rs(0).Value) - ConvertToSeconds(sStartTime) - ConvertToSeconds(sRaceDelay) - ConvertToSeconds(MultiRslts(10, x)) - CSng(sngMyStart), 3))
            MultiRslts(4, x) = rs(2).Value
            If iNumLegs > 2 Then
                MultiRslts(5, x) = ConvertToMinutes(Round(ConvertToSeconds(rs(3).Value) - ConvertToSeconds(rs(1).Value), 3))
                MultiRslts(6, x) = rs(5).Value
                MultiRslts(7, x) = ConvertToMinutes(Round(CSng(MultiRslts(8, x)) + ConvertToSeconds(sStartTime) + ConvertToSeconds(sRaceDelay) + ConvertToSeconds(MultiRslts(10, x)) - ConvertToSeconds(rs(4).Value), 3))
                MultiRslts(14, x) = ConvertToMinutes(Round(ConvertToSeconds(MultiRslts(3, x)) + ConvertToSeconds(MultiRslts(5, x)) + ConvertToSeconds(MultiRslts(7, x)), 3))
            Else
                MultiRslts(5, x) = ConvertToMinutes(CSng(MultiRslts(8, x)) + ConvertToSeconds(sStartTime) + ConvertToSeconds(sRaceDelay) + ConvertToSeconds(MultiRslts(10, x)) + CSng(sngMyStart) - ConvertToSeconds(rs(1).Value))
                MultiRslts(14, x) = ConvertToMinutes(Round(ConvertToSeconds(MultiRslts(3, x)) + ConvertToSeconds(MultiRslts(5, x)), 3))
            End If
        End If
        rs.Close
        Set rs = Nothing
    Next
        
    'get leg 1 rank
    For x = 0 To UBound(MultiRslts, 2) - 2
        For z = x + 1 To UBound(MultiRslts, 2) - 1
            If ConvertToSeconds(MultiRslts(3, x)) > ConvertToSeconds(MultiRslts(3, z)) Then
                For y = 0 To 16
                    SortArr(y) = MultiRslts(y, z)
                    MultiRslts(y, z) = MultiRslts(y, x)
                    MultiRslts(y, x) = SortArr(y)
                Next
            End If
        Next
    Next
        
    'enter leg 1 rank
    y = 1
    For x = 0 To UBound(MultiRslts, 2) - 1
        If ConvertToSeconds(MultiRslts(3, x)) = 0 Then
            MultiRslts(11, x) = "---"
        Else
            MultiRslts(11, x) = y
            y = y + 1
        End If
    Next
        
    'get leg 2 rank
    For x = 0 To UBound(MultiRslts, 2) - 2
        For z = x + 1 To UBound(MultiRslts, 2) - 1
            If ConvertToSeconds(MultiRslts(5, x)) > ConvertToSeconds(MultiRslts(5, z)) Then
                For y = 0 To 16
                    SortArr(y) = MultiRslts(y, z)
                    MultiRslts(y, z) = MultiRslts(y, x)
                    MultiRslts(y, x) = SortArr(y)
                Next
            End If
        Next
    Next
        
    'enter leg 2 rank
    y = 1
    For x = 0 To UBound(MultiRslts, 2) - 1
        If ConvertToSeconds(MultiRslts(5, x)) = 0 Then
            MultiRslts(12, x) = "---"
        Else
            MultiRslts(12, x) = y
            y = y + 1
        End If
    Next
        
    If iNumLegs > 2 Then
        'get leg 3 rank
        For x = 0 To UBound(MultiRslts, 2) - 2
            For z = x + 1 To UBound(MultiRslts, 2) - 1
                If ConvertToSeconds(MultiRslts(7, x)) > ConvertToSeconds(MultiRslts(7, z)) Then
                    For y = 0 To 16
                        SortArr(y) = MultiRslts(y, z)
                        MultiRslts(y, z) = MultiRslts(y, x)
                        MultiRslts(y, x) = SortArr(y)
                    Next
                End If
            Next
        Next
            
        'enter leg 3 rank
        y = 1
        For x = 0 To UBound(MultiRslts, 2) - 1
            If ConvertToSeconds(MultiRslts(7, x)) = 0 Then
                MultiRslts(13, x) = "---"
            Else
                MultiRslts(13, x) = y
                y = y + 1
            End If
        Next
    End If

    'sort by elapsed time
    For y = 0 To UBound(MultiRslts, 2) - 2
        For x = y + 1 To UBound(MultiRslts, 2) - 1
            If sSortBy = "active" Then
                If ConvertToSeconds(MultiRslts(14, y)) > ConvertToSeconds(MultiRslts(14, x)) Then
                    For z = 0 To 16
                        SortArr(z) = MultiRslts(z, x)
                        MultiRslts(z, x) = MultiRslts(z, y)
                        MultiRslts(z, y) = SortArr(z)
                    Next
                End If
            Else
                If ConvertToSeconds(MultiRslts(8, y)) > ConvertToSeconds(MultiRslts(8, x)) Then
                    For z = 0 To 16
                        SortArr(z) = MultiRslts(z, x)
                        MultiRslts(z, x) = MultiRslts(z, y)
                        MultiRslts(z, y) = SortArr(z)
                    Next
                End If
            End If
        Next
    Next
End Sub

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<%
    
Private Sub MultiTmData(iThisBib, iThisLeg)
    Dim lThisPart

    sMmbrName = "Unknown"

    Select Case CInt(iThisLeg)
        Case 1
            sLegName = sLeg1Name
        Case 2
            sLegName = sLeg2Name
        Case 3
            sLegName = sLeg3Name
    End Select

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM PartRace WHERE RaceID = " & lRaceID & " AND Bib = " & iThisBib
    rs.Open sql, conn, 1, 2
    lThisPart = rs(0).Value
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MmbrName, Age, Gender FROM MultiTmMmbrs WHERE RaceID = " & lRaceID & " AND ParticipantID = " & lThisPart & " AND MmbrNum = " & iThisLeg
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sMmbrName = rs(0).Value
        iMmbrAge = rs(1).Value
        sMmbrGender = rs(2).Value
    End If
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Multi-Sport Results with Splits</title>
<meta name="description" content="Gopher State Events Multi-Event Results/Results With Splits.">
 <!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Transition Results">

	<h1 class="h1">Gopher State Events Multi-Sport Results with Splits:&nbsp;<%=sEventName%>&nbsp;On&nbsp;<%=dEventDate%></h1>

    <%If UBound(Races, 2) > 0 Then%>
		<form class="form-inline" name="get_races" method="post" action="multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;sort_by=<%=sSortBy%>">
        <div class="form-group">
		    <label for="races">Race:</label>
		    <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
			    <%For i = 0 to UBound(Races, 2)%>
				    <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
					    <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
				    <%Else%>
					    <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
				    <%End If%>
			    <%Next%>
		    </select>
		    <input type="hidden" class="form-control" name="submit_race" id="submit_race" value="submit_race">
		    <input type="submit" class="form-control" name="get_race" id="get_race" value="View">
        </div>
		</form>
    <%Else%>
        <h2 class="h2"><%=sRaceName%></h2>
    <%End If%>

    <div class="bg-warning">
        Note: ELAPSED TIME, commonly used for triathlons and duathlons to determine order of finish, is a participant's time from the start of the race 
        until the participant or team finishes.  ACTIVE TIME, commonly used in stair climbs with multiple legs, eliminates time in transition and is the 
        actual time spent "on the course."
    </div>

    <ul class="list-inline">
        <li class="list-group-item"><a href="javascript:window.print();">Print</a></li>
        <li class="list-group-item"><a href="trans_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">Transition Data</a></li>
        <li class="list-group-item">Number of Legs: <%=iNumLegs%></li>
        <li class="list-group-item">Leg 1: <%=sLeg1Name%> (<%=sngLeg1Dist%>)</li>
        <li class="list-group-item">Leg 2: <%=sLeg2Name%> (<%=sngLeg2Dist%>)</li>
        <%If iNumLegs = "3" Then%>
            <li class="list-group-item">Leg 3: <%=sLeg3Name%> (<%=sngLeg3Dist%>)</li>
        <%End If%>
        <li class="list-group-item"><a href="multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;sort_by=elapsed">Sort By Elapsed Time</a></li>
        <li class="list-group-item"><a href="multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;sort_by=active">Sort By Active Time</a></li>
    </ul>

    <div class="row">
        <%For j = 0 To 2%>
            <%If j = 0 Then%>
                <%Call RsltsByMF("X")%>
                <h4 class="h4">Combined Results</h4>
            <%ElseIf j = 1 Then%>
                <%Call RsltsByMF("M")%>
                <h4 class="h4">Male Results</h4>
            <%ElseIf j = 2 Then%>
                <%Call RsltsByMF("F")%>
                <h4 class="h4">Female Results</h4>
            <%End If%>
          
            <%If CInt(iNumLegs) > 2 Then%>  
                <table class="table table-striped">
                    <tr>
                        <th rowspan="2">PL</th>
                        <th rowspan="2">BIB-PARTICIPANT</th>
                        <th rowspan="2">AGE</th>
                        <th rowspan="2">GENDER</th>
                        <th style="text-align: center;" colspan="2"><%=UCase(sLeg1Name)%></th>
                        <th rowspan="2">TRANSITION</th>
                        <th style="text-align: center;" colspan="2"><%=UCase(sLeg2Name)%></th>
                        <th rowspan="2">TRANSITION</th>
                        <th style="text-align: center;" colspan="2"><%=UCase(sLeg3Name)%></th>
                        <th style="text-align: center;" rowspan="2">ELAPSED</th>
                        <th style="text-align: center;" rowspan="2">ACTIVE</th>
                    </tr>
                    <tr>
                        <th style="text-align: center;">TIME</th>
                        <th>RANK</th>
                        <th style="text-align: center;">TIME</th>
                        <th>RANK</th>
                        <th style="text-align: center;">TIME</th>
                        <th>RANK</th>
                    </tr>
                    <%For i = 0 To UBound(MultiRslts, 2) - 1%>
                        <tr>
                            <td><%=i + 1%></td>
                            <td><%=MultiRslts(0, i)%></td>
                            <td><%=MultiRslts(2, i)%></td>
                            <td><%=MultiRslts(1, i)%></td>
                            <td><%=MultiRslts(3, i)%></td>
                            <td><%=MultiRslts(11, i)%></td>
                            <td><%=MultiRslts(4, i)%></td>
                            <td><%=MultiRslts(5, i)%></td>
                            <td><%=MultiRslts(12, i)%></td>
                            <td><%=MultiRslts(6, i)%></td>
                            <td><%=MultiRslts(7, i)%></td>
                            <td><%=MultiRslts(13, i)%></td>
                            <td><%=ConvertToMinutes(CSng(MultiRslts(8, i)))%></td>
                            <td><%=MultiRslts(14, i)%></td>
                        </tr>

                        <%If MultiRslts(16, i) = "team" Then%>
                            <tr>
                                <td colspan="14">
                                    <ul class="list-inline"  style="margin-left: 50px;">
                                        <%For m = 0 To CInt(iNumLegs) - 1%>
                                            <%Call MultiTmData(MultiRslts(9, i), m + 1)%>
                                        
                                            <li style="font-weight: bold;"><%=sLegName%>:</li>
                                            <li><%=sMmbrName%></li>
                                            <li><%=iMmbrAge%></li>
                                            <li><%=sMmbrGender%></li>
                                        <%Next%>
                                    </ul>
                                </td>
                            </tr>
                        <%End If%>
                    <%Next%>
                </table>                    
            <%Else%>
                <table class="table table-striped">
                    <tr>
                        <th rowspan="2">PL</th>
                        <th rowspan="2">BIB-PARTICIPANT</th>
                        <%If sShowAge = "y" Then%>
                            <th rowspan="2">AGE</th>
                        <%End If%>
                        <th style="text-align: center;border-bottom:1px solid #555;" colspan="2"><%=UCase(sLeg1Name)%></th>
                        <th rowspan="2">TRANSITION</th>
                        <th style="text-align: center;border-bottom:1px solid #555;" colspan="2"><%=UCase(sLeg2Name)%></th>
                        <th style="text-align: center;" rowspan="2">ELAPSED</th>
                        <th style="text-align: center;" rowspan="2">ACTIVE</th>
                    </tr>
                    <tr>
                        <th style="text-align: center;">TIME</th>
                        <th>RANK</th>
                        <th style="text-align: center;">TIME</th>
                        <th>RANK</th>
                    </tr>
                    <%m = 0%>
                    <%For i = 0 To UBound(MultiRslts, 2) - 1%>
                        <%If ConvertToSeconds(MultiRslts(5, i)) > 0 And ConvertToSeconds(MultiRslts(4, i)) > 0 Then%>
                            <%m = m + 1%>
                            <tr>
                                <td><%=m%></td>
                                <td><%=MultiRslts(0, i)%></td>
                                <%If sShowAge = "y" Then%>
                                    <td><%=MultiRslts(2, i)%></td>
                                <%End If%>
                                <td style="text-align: center;"><%=MultiRslts(3, i)%></td>
                                <td style="text-align: center;"><%=MultiRslts(11, i)%></td>
                                <td style="text-align: center;"><%=MultiRslts(4, i)%></td>
                                <td style="text-align: center;"><%=MultiRslts(5, i)%></td>
                                <td style="text-align: center;"><%=MultiRslts(12, i)%></td>
                                <td style="text-align: center;"><%=ConvertToMinutes(CSng(MultiRslts(8, i)))%></td>
                                <td style="text-align: center;"><%=MultiRslts(14, i)%></td>
                            </tr>
                        <%End If%>

                        <%If MultiRslts(16, i) = "team" Then%>
                            <tr>
                                <td colspan="13">
                                    <ul class="list-inline"  style="margin-left: 50px;">
                                        <%For m = 0 To CInt(iNumLegs) - 1%>
                                            <%Call MultiTmData(MultiRslts(15, i), m + 1)%>
                                        
                                            <li style="font-weight: bold;"><%=sLegName%>:</li>
                                            <li><%=sMmbrName%></li>
                                            <li><%=iMmbrAge%></li>
                                            <li><%=sMmbrGender%></li>
                                        <%Next%>
                                    </ul>
                                </td>
                            </tr>
                        <%End If%>
                    <%Next%>
                </table>                    
            <%End If%>
            <br>
        <%Next%>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>