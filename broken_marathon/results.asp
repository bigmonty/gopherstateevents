<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k
Dim lOCMen, lOCWomen
Dim sRaceTime, sAGTime, sAGFactor, sAGPct
Dim BMParts(), BMRaces(2, 2), SortArray(14)

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

BMRaces(0, 0) = "1052"
BMRaces(1, 0) = "Run New Prague"
BMRAces(2, 0) = "607"
BMRaces(0, 1) = "1216"
BMRaces(1, 1) = "Mora"
BMRAces(2, 1) = "696"
BMRaces(0, 2) = "1133"
BMRaces(1, 2) = "Gandy Dancer"
BMRAces(2, 2) = "645"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT OCMen, OCWomen FROM AgeGrOCTime WHERE AgeGrDistID = 11"
rs.Open sql, conn, 1, 2
lOCMen = rs(0).Value
lOCWomen = rs(1).Value
rs.Close
Set rs = Nothing

i = 0
j = 0
ReDim BMParts(14, 0)
For i = 0 To UBound(BMRaces, 2)
    Set rs= Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.ParticipantID, p.FirstName, p.LastName, p.Gender, pr.Age FROM IndResults ir INNER JOIN Participant p "
    sql = sql & "ON ir.ParticipantID = p.ParticipantID INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID AND ir.RaceID = pr.RaceID "
    sql = sql & "WHERE ir.RaceID = " & BMRaces(0, i) & " AND (p.Gender IN ('M', 'F')) AND ir.FnlTime IS NOT NULL AND ir.FnlTime > '00:00:00.000' "
    sql = sql & "AND pr.Age < 99 ORDER BY p.LastName, p.FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Call GetMyData(rs(0).Value, BMRaces(0, i), rs(4).Value, rs(3).Value)
        BMParts(0, j) = rs(0).Value                                                                     'participantID
        BMParts(1, j) = Replace(rs(2).Value, "''", "'") & ", " &   Replace(rs(1).Value, "''", "'")      'name
        BMParts(2, j) = rs(3).Value                                                                     'gender
        BMParts(3, j) = rs(4).Value                                                                     'age
        Select Case i
            Case 0
                BMParts(4, j) = sRaceTime
                BMParts(5, j) = sAGTime
                BMParts(6, j) = sAGPct
            Case 1
                BMParts(7, j) = sRaceTime
                BMParts(8, j) = sAGTime
                BMParts(9, j) = sAGPct
            Case 2
                BMParts(10, j) = sRaceTime
                BMParts(11, j) = sAGTime
                BMParts(12, j) = sAGPct
        End Select

        'this will have to be changed to show the sum of top two, excluding those with only one race after Mora
        BMParts(13, j) = BMParts(5, j)
        BMParts(14, j) = BMParts(6, j)

        j = j + 1
        ReDim Preserve BMParts(14, j)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
Next

'sort by combined time
For i = 0 To UBound(BMParts, 2) - 2
    For j = i + 1 To UBound(BMParts, 2) - 1
        If ConvertToSeconds(BMParts(14, i)) < ConvertToSeconds(BMParts(14, j)) Then
            For k = 0 To 14
                SortArray(k) = BMParts(k, i)
                BMParts(k, i) = BMParts(k, j)
                BMParts(k, j) = SortArray(k)
            Next
        End If
    Next
Next

Private Sub GetMyData(lPartID, lRaceID, iAge, sGender)
    Dim sngRaceTime, sngAGTime

    sRaceTime = vbNullString
    sAGTime = vbNullString
    sAGFactor = vbNullString
    sAGPct = 0
    sngRaceTime = 0

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FnlScnds FROM IndResults WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then 
        sngRaceTime = rs2(0).Value
        sRaceTime = ConvertToMinutes(rs2(0).Value)
    End If
    rs2.Close
    Set rs2 = Nothing

    If CSng(sngRaceTime) > 0 Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT Factor FROM AgeGrFactors WHERE MF = '" & LCase(sGender) & "' AND Age = " & iAge & " AND AgeGrDistID = 11"
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then 
            sAGFactor = rs2(0).Value
            sngAGTime = CSng(sngRaceTime)*CSng(rs2(0).Value)
            sAGTime = ConvertToMinutes(sngAGTime)
            If UCASE(sGender) = "M" Then
                sAGPct = Round(CLng(lOCMen)/CSng(sngAGTime), 4)*100
            ElseIf UCASE(sGender) = "F" Then
                sAGPct = Round(CLng(lOCWomen)/CSng(sngAGTime), 4)*100
            End If
        End If
        rs2.Close
        Set rs2 = Nothing
    End If
End Sub
%>

<!--#include file = "../includes/convert_to_minutes.asp" -->
<!--#include file = "../includes/convert_to_seconds.asp" -->

<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events Broken Marathon Standings</title>
 
<link href="//cdn.datatables.net/1.10.2/css/jquery.dataTables.css" rel="stylesheet" type="text/css">
    
<script src="//code.jquery.com/jquery-2.1.4.min.js"></script>
<script src="//cdn.datatables.net/1.10.2/js/jquery.dataTables.min.js"></script>

<!-- bootstrap JavaScript & CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
</head>

<body>
<div class="container">
    <div class="row" style="padding: 5px;">
        <div class="col-xs-6">
            <img src="/graphics/html_header.png" alt="Series Header" class="img-responsive">
        </div>
        <div class="col-xs-6">
            <img src="broken_marathon_banner.png" alt="Broken Marathon" class="img-responsive" style="float: right;">
        </div>
    </div>

    <div class="row">
        <div class="bg-info col-sm-3" style="text-align: center;">
            <a href="http://www.gopherstateevents.com/misc/broken_marathon.asp">What is a "Broken Marathon"?</a>
        </div>
        <div class="bg-warning col-sm-3" style="text-align: center;">
            <a href="https://vasaloppet.us/event/2017-mora-half-marathon-5k/">Register for Mora</a>
        </div>
        <div class="bg-danger col-sm-3" style="text-align: center;">
            <a href="https://www.zapevent.com/reg/event/12626">Register for Gandy Dancer</a>
        </div>
        <div class="bg-info col-sm-3" style="text-align: center;">
            <a href="http://www.tempotickets.com/BrokenMarathon17">Register for the Broken Marathon</a>
        </div>
    </div>

    <div class="row">
        <h2 class="h2">Broken Marathon Standings 2017</h2>

        <div class="bg-danger text-danger">
            IMPORTANT NOTE:  While the goal, in theory, is to sort by the two fastest combined times regardless of age or gender, this is actually not possible
            to do across genders.  As a result, the order of finish will be in terms of Age Graded % which fairly indicates best performance across genders. 
        </div>
    </div>

    <div class="row">
        <div class="table-responsive">
            <table class="table table-striped">
                <tr>
                    <th rowspan="2">Pl</th>
                    <th rowspan="2">Name</th>
                    <th rowspan="2">Gender</th>
                    <th rowspan="2">Age</th>
                    <%For i = 0 To UBound(BMRaces, 2)%>
                        <th colspan="3" style="text-align:center;">
                            <a href="/results/fitness_events/results.asp?event_type=5&event_id=<%=BMRaces(2, i)%>&amp;race_id=<%=BMRaces(0, i)%>"><%=BMRaces(1, i)%></a>
                        </th>
                    <%Next%>
                    <th colspan="2" style="text-align:center;">Top 2 Combined</th>
                </tr>
                <tr>
                    <%For i = 0 To UBound(BMRaces, 2)%>
                        <th style="text-align:center;">Time</th>
                        <th style="text-align:center;">AGTime</th>
                        <th style="text-align:center;">AG%</th>
                    <%Next%>
                    <th style="text-align:center;">AGTime</th>
                    <th style="text-align:center;">AG%</th>
                </tr>
                <%For i = 0 To UBound(BMParts, 2) - 1%>
                    <tr>
                        <td><%=i + 1%></td>
                        <%For j = 1 To 14%>
                            <%Select Case j%>
                                <%Case "6"%>
                                    <td><%=BMParts(j, i)%>%</td>
                                <%Case "9"%>
                                    <td><%=BMParts(j, i)%>%</td>
                                <%Case "12"%>
                                    <td><%=BMParts(j, i)%>%</td>
                                <%Case "14"%>
                                    <td><%=BMParts(j, i)%>%</td>
                                <%Case Else%>
                                    <td><%=BMParts(j, i)%></td>
                            <%End Select%>
                        <%Next%>
                    </tr>
                <%Next%>
            </table>
        </div>
    </div>
</div>
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
