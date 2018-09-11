<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID
Dim iThisBib, iThisLap, iMyAge, iNumLaps, iPl
Dim sEventName, sRaceName, sGender, sMyName, sMyGender
Dim dEventDate
Dim Results(), SortArr(5)

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

sql = "SELECT RaceName, Numlaps FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
iNumLaps = rs(1).Value
Set rs = Nothing

i = 0
ReDim Results(5, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Bib, ReadTime, RaceLapsID FROM RaceLaps WHERE RaceID = " & lRaceID & " ORDER BY Bib, RaceLapsID"
rs.Open sql, conn, 1, 2     'use race laps id as placeholder for lap num
Do While Not rs.EOF
    Call MyData(rs(0).Value)

    Results(0, i) = rs(0).Value
    Results(1, i) = rs(1).Value
    Results(2, i) = rs(2).Value
    Results(3, i) = sMyName
    Results(4, i) = sMyGender
    Results(5, i) = iMyAge
    i = i + 1
    ReDim Preserve Results(5, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
    
If UBound(Results, 2) > 0 Then
    'go through and assign lap time for each bib
    For i = 0 To UBound(Results, 2) - 1
        If i = 0 Then           'this is just for the very first record in the array
            iThisLap = 1
            iThisBib = Results(0, i)
            Results(2, i) = iThisLap
        Else
            If CInt(Results(0, i)) = iThisBib Then 'if it's the same bib then just increment the lap and record that
                iThisLap = iThisLap + 1
                Results(2, i) = iThisLap
            Else
                iThisLap = 1
                iThisBib = Results(0, i)
                Results(2, i) = iThisLap
            End If
        End If
    Next
        
    're-sort by time
    For i = 0 To UBound(Results, 2) - 2
        For j = i + 1 To UBound(Results, 2) - 1
            If CLng(Results(1, i)) > CLng(Results(1, j)) Then
                For k = 0 To 5
                    SortArr(k) = Results(k, i)
                    Results(k, i) = Results(k, j)
                    Results(k, j) = SortArr(k)
                Next
            End If
        Next
    Next
End If

Private Sub MyData(iMyBib)
    Dim rs2, sql2
          
    sMyName = vbNullString
    sMyGender = vbNullString
    iMyAge = 0
    
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT p.FirstName, p.LastName, p.Gender, pr.Age FROM Participant p INNER JOIN PartRace pr "
    sql2 = sql2 & "ON p.ParticipantID = pr.ParticipantID WHERE pr.Bib = " & iMyBib & " AND pr.RaceID = " & lRaceID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then
        sMyName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
        sMyGender = rs2(2).Value
        iMyAge = rs2(3).Value
        If rs2(3).Value = "99" Then
            iMyAge = "na"
        Else
            iMyAge = rs2(3).Value
        End IF
   End If
    rs2.Close
    Set rs2 = Nothing
End Sub

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
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Individual Finish By Lap</title>
<meta name="description" content="Gopher State Events Individual Finish By Lap.">
 <!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Results by Lap">

    <div class="bg-danger">
        <a href="javascript:window.print();" style="color:#fff;">Print</a>
        &nbsp;|&nbsp;
        <a href="rslts_w_laps.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" style="color:#fff;">Results with Laps</a>
    </div>
	<h1 class="h1">Gopher State Events Individual Order By Lap</h1>
    <h2 class="h2"><%=sEventName%>&nbsp;On&nbsp;<%=dEventDate%></h2>
    <h3 class="h3"><%=sRaceName%></h3>

    <div class="row">
        <%For j = 0 To 1%>
            <%For k = 1 To iNumLaps%>
                <%Select Case j%>
                    <%Case "0"%>
                        <h4 class="h4">MALE RESULTS LAP <%=k%></h4>
                        <%sGender = "M"%>
                    <%Case "1"%>
                        <h4 class="h4">FEMALE RESULTS LAP <%=k%></h4>
                         <%sGender = "F"%>
                <%End Select%>

                <table class="table table-striped">
                    <tr>
                        <th>PL</th>
                        <th>BIB</th>
                        <th>NAME</th>
                        <th>AGE</th>
                        <th>LAP TIME</th>
                    </tr>
                    <%iPl = 1%>
                    <%For i = 0 To UBound(Results, 2) - 1%>
                        <%If Results(4, i) = sGender Then            'make sure it is the correct gender%>
                            <%If CInt(Results(2, i)) = k Then        'make sure it is the correct lap time%>
                                <tr>
                                    <td><%=iPl%>)</td>
                                    <td><%=Results(0, i)%></td>
                                    <td><%=Results(3, i)%></td>
                                    <td><%=Results(5, i)%></td>
                                    <td><%=ConvertToMinutes(Results(1, i))%></td>
                                </tr>
                                <%iPl = CInt(iPl) + 1%>
                            <%End If%>
                        <%End If%>
                    <%Next%>
                </table>
            <%Next%>
        <%Next%>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>