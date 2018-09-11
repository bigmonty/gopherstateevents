<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k
Dim lRaceID, lEventID
Dim iThisBib, iThisLap, iNumLaps, iPl, iArrayDim, iLap
Dim sEventName, sRaceName, sGender
Dim dEventDate
Dim SortArr(), LapMatrix()

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

iArrayDim = 2*iNumLaps + 2

i = 0
ReDim LapMatrix(iArrayDim, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT pr.Bib, p.FirstName, p.LastName, p.Gender, pr.Age FROM PartRace pr INNER JOIN Participant p ON pr.ParticipantID = p.ParticipantID "
sql = sql & "WHERE pr.RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    LapMatrix(0, i) = rs(0).Value
    LapMatrix(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
    LapMatrix(2, i) = rs(3).Value
    If rs(4).Value = "99" Then
        LapMatrix(3, i) = "na"
    Else
        LapMatrix(3, i) = rs(4).Value
    End IF
    LapMatrix(4, i) = MySplit(1, rs(0).Value)
    LapMatrix(5, i) = MySplit(2, rs(0).Value)

    Select Case CInt(iNumLaps)
        Case 2
            LapMatrix(6, i) = CSng(LapMatrix(4, i)) + CSng(LapMatrix(5, i))
        Case 3
            LapMatrix(6, i) = MySplit(3, rs(0).Value)

            LapMatrix(7, i) = CSng(LapMatrix(4, i)) + CSng(LapMatrix(5, i))
            LapMatrix(8, i) = CSng(LapMatrix(7, i)) + CSng(LapMatrix(6, i))
        Case 4
            LapMatrix(6, i) = MySplit(3, rs(0).Value)
            LapMatrix(7, i) = MySplit(4, rs(0).Value)

            LapMatrix(8, i) = CSng(LapMatrix(4, i)) + CSng(LapMatrix(5, i))
            LapMatrix(9, i) = CSng(LapMatrix(8, i)) + CSng(LapMatrix(6, i))
            LapMatrix(10, i) = CSng(LapMatrix(9, i)) + CSng(LapMatrix(7, i))
        Case 5
            LapMatrix(6, i) = MySplit(3, rs(0).Value)
            LapMatrix(7, i) = MySplit(4, rs(0).Value)
            LapMatrix(8, i) = MySplit(5, rs(0).Value)

            LapMatrix(9, i) = CSng(LapMatrix(4, i)) + CSng(LapMatrix(5, i))
            LapMatrix(10, i) = CSng(LapMatrix(9, i)) + CSng(LapMatrix(6, i))
            LapMatrix(11, i) = CSng(LapMatrix(10, i)) + CSng(LapMatrix(7, i))
            LapMatrix(12, i) = CSng(LapMatrix(11, i)) + CSng(LapMatrix(8, i))
    End Select

    i = i + 1
    ReDim Preserve LapMatrix(iArrayDim, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
        
're-sort by last time
ReDim SortArr(iArrayDim)
For i = 0 To UBound(LapMatrix, 2) - 2
    For j = i + 1 To UBound(LapMatrix, 2) - 1
        If CLng(LapMatrix(iArrayDim, i)) > CLng(LapMatrix(iArrayDim, j)) Then
            For k = 0 To iArrayDim
                SortArr(k) = LapMatrix(k, i)
                LapMatrix(k, i) = LapMatrix(k, j)
                LapMatrix(k, j) = SortArr(k)
            Next
        End If
    Next
Next

Private Function MySplit(iWhichSplit, iMyBib)
    Dim iWhichRcd

    iWhichRcd = 1
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT ReadTime FROM RaceLaps WHERE RaceID = " & lRaceID & " AND Bib = " & iMyBib & " ORDER BY RaceLapsID"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        If CInt(iWhichSplit) = CInt(iWhichRcd) Then
            MySplit = rs2(0).Value
            Exit Do
        Else
            iWhichRcd = CInt(iWhichRcd) + 1
        End If
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End Function

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
<title>GSE Results with Laps</title>
<meta name="description" content="Gopher State Events Results With Lap Times.">
 <!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Relay Results">

    <div class="bg-danger">
        <a href="javascript:window.print();" style="color:#fff;">Print</a>
        &nbsp;|&nbsp;
        <a href="rslts_by_lap.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" style="color:#fff;">Individual Finish By Lap</a>
    </div>
	<h1 class="h1">Gopher State Events Results with Lap Times</h1>
    <h2 class="h2"><%=sEventName%>&nbsp;On&nbsp;<%=dEventDate%></h2>
    <h3 class="h3"><%=sRaceName%></h3>

    <div class="row">
        <%For k = 0 To 1%>
            <%Select Case k%>
                <%Case "0"%>
                    <h4 class="h4">Male</h4>
                    <%sGender = "M"%>
                <%Case "1"%>
                    <h4 class="h4">Female</h4>
                    <%sGender = "F"%>
            <%End Select%>

            <%iLap = 1%>
            <table class="table table-striped">
                <tr>
                    <th>Pl</th>
                    <th>Bib</th>
                    <th>Name</th>
                    <th>Age</th>
                    <th>Lap <%=iLap%></th>
                    <%For j = 5 To iNumLaps + 3%>
                        <%iLap = CInt(iLap) + 1%>
                        <th>Lap <%=iLap%></th>
                        <th>Combined</th>
                    <%Next%>
                </tr>
                <%iPl = 1%>
                <%For i = 0 To UBound(LapMatrix, 2) - 1%>
                    <%If LapMatrix(2, i) = sGender Then%>
                        <%If Not CSng(LapMatrix(iArrayDim, i)) = 0 Then%>
                            <tr>
                                <td><%=iPl%></td>
                                <td><%=LapMatrix(0, i)%></td>
                                <td><%=LapMatrix(1, i)%></td>
                                <td><%=LapMatrix(3, i)%></td>
                                <td><%=ConvertToMinutes(LapMatrix(4, i))%></td>
                                <%For j = 2 To iNumLaps%>
                                    <td><%=ConvertToMinutes(LapMatrix(j + 3, i))%></td>
                                    <td><%=ConvertToMinutes(LapMatrix(iNumLaps + j + 2, i))%></td>
                                <%Next%>
                            </tr>
                            <%iPl = CInt(iPl) + 1%>
                        <%End If%>
                    <%End If%>
                <%Next%>
            </table>
        <%Next%>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>