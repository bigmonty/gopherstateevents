<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j
Dim iNumYears
Dim sShowWhat, sLineColor, sThisDay, sThisMonth, sShowAvg, sThisEvent
Dim sngMaxY, sngMinY, sngNumParts
Dim Years(), Events()
Dim Chart

sShowAvg = Request.QueryString("show_avg")
sShowWhat = Request.QueryString("show_what")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sngMaxY = 0
sngMinY = 100000

iNumYears = Year(Date) - CInt(2013) + 1

ReDim Years(iNumYears)

For i = 0 To UBound(Years) - 1
    Years(i) = i + 2013 'which year
Next

sThisDay = Day(Date)
sThisMonth = Month(Date)

Private Sub GetThisData(iThisMonth, iThisYear)
    Call GetEvents(iThisMonth, iThisYear)

    Select Case sShowWhat
        Case "all"
            sngNumParts = CSng(sngNumParts) + CSng(GetParts(iThisMonth, iThisYear, "Fitness Event"))
            sngNumParts = CSng(sngNumParts) + CSng(GetParts(iThisMonth, iThisYear, "Cross-Country"))
            sngNumParts = CSng(sngNumParts) + CSng(GetParts(iThisMonth, iThisYear, "Nordic Ski"))
        Case "fitness"
            sngNumParts = CSng(sngNumParts) + CSng(GetParts(iThisMonth, iThisYear, "Fitness Event"))
        Case "cc"
            sngNumParts = CSng(sngNumParts) + CSng(GetParts(iThisMonth, iThisYear, "Cross-Country"))
        Case "nordic"
            sngNumParts = CSng(sngNumParts) + CSng(GetParts(iThisMonth, iThisYear, "Nordic Ski"))
    End Select
End Sub

Private Sub GetEvents(sCurrMonth, sCurrYear)
    Dim y

    ReDim Events(2, 0)

    If sShowWhat = "fitness" or sShowWhat = "all" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & sCurrYear 
        sql = sql & "' AND EventDate <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "') ORDER BY EventDate"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Month(rs(1).Value) = sCurrMonth Then
                Events(0, y) = rs(0).Value
                Events(1, y) = "Fitness Event"
                Events(2, y) = rs(1).Value
                y = y + 1
                ReDim Preserve Events(2, y)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
        
    If sShowWhat = "cc" or sShowWhat = "all" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetsID, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & sCurrYear 
        sql = sql & "' AND MeetDate <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "') AND Sport = 'Cross-Country' "
        sql = sql & "ORDER BY MeetDate"
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If Month(rs(1).Value) = sCurrMonth Then
                Events(0, y) = rs(0).Value
                Events(1, y) = rs(2).Value
                Events(2, y) = rs(1).Value
                y = y + 1
                ReDim Preserve Events(2, y)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
        
    If sShowWhat = "nordic" or sShowWhat = "all" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetsID, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & sCurrYear 
        sql = sql & "' AND MeetDate <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "') AND Sport = 'Nordic Ski' "
        sql = sql & "ORDER BY MeetDate"
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If Month(rs(1).Value) = sCurrMonth Then
                Events(0, y) = rs(0).Value
                Events(1, y) = rs(2).Value
                Events(2, y) = rs(1).Value
                y = y + 1
                ReDim Preserve Events(2, y)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Function GetParts(sCurrMonth, sCurrYear, sThisSport)
    Dim x
    Dim sEventRaces

    GetParts = 0

    For x = 0 To UBound(Events, 2) - 1
        If sThisSport = "Fitness Event" Then
            sEventRaces = GetEventRaces(Events(0, x))
    
            If Len(sEventRaces) > 0 Then
                Set rs2 = Server.CreateObject("ADODB.Recordset")
                sql2 = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ")"
                rs2.Open sql2, conn, 1, 2
                If rs2.RecordCount > 0 Then GetParts = CInt(GetParts) + rs2.RecordCount
                rs2.Close
                Set rs2 = Nothing
            End If
        Else
            If sThisSport = Events(1, x) Then
                Set rs2 = Server.CreateObject("ADODB.Recordset")
                sql2 = "SELECT IndRsltsID FROM IndRslts WHERE MeetsID = " & Events(0, x)
                rs2.Open sql2, conn2, 1, 2
                If rs2.RecordCount > 0 Then GetParts = CInt(GetParts) + rs2.RecordCount
                rs2.Close
                Set rs2 = Nothing
            End If
        End If
    Next
End Function

Private Function GetEventRaces(lThisEvent)
    GetEventRaces = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT RaceID FROM RaceData WHERE EventID = " & lThisEvent
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        GetEventRaces = GetEventRaces & rs2(0).Value & ", "
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If Not GetEventRaces = vbNullString Then GetEventRaces = Left(GetEventRaces, Len(GetEventRaces) - 2)
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Finance Graphs: Monthly Graph</title>
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<%
'now begin setting up the display
Response.Expires = 0
Response.Buffer = true
Response.Clear

Set Chart = Server.CreateObject("csDrawGraph64.Draw")

Chart.Width = 775
Chart.Height = 350
Chart.OriginX = 60
Chart.OriginY = 300
Chart.MaxX = 500
Chart.MaxY = 275
Chart.XOffset = 1
Chart.XTop = 12
Chart.YOffset = 0
Chart.YTop = sngMaxY
Chart.ShowGrid = true
Chart.XGrad = 1
Chart.YGrad = 0
Chart.XAxisText = "Month"
Chart.YAxisText = "Value"
Chart.LegendX = 650
Chart.LegendY = 50
Chart.ShowLegend = true
Chart.ShowLegendBox = true
Chart.LegendTextSize = 9
Chart.Title = sShowWhat
Chart.TitleX = 75
Chart.TitleY = 50
Chart.TitleSize = 10
Chart.TitleBold = true
Chart.TitleColor = "000000"

'show ind graph
For i = 0 To UBound(Years) - 1
    Select Case i
        Case 0
            sLineColor = "86194d"
        Case 1
            sLineColor = "009933"
        Case 2
            sLineColor = "008877"
        Case 3
            sLineColor = "3c38e8"
        Case 4
            sLineColor = "f45628"
        Case 5
            sLineColor = "000000"
    End Select

    For j = 1 To 12
        Call GetThisData(j, Years(i))
        If j <= CInt(sThisMonth) Then
	        Chart.AddPoint j, sngNumParts, sLineColor, Years(i)
'            Chart.AddLineGraphText sngNumParts, j, sngNumParts, 0
        End If

        If j = 12 Then sngNumParts = 0
    Next
Next

Response.ContentType = "Image/Gif"
Response.BinaryWrite Chart.GIFLine

Response.End

conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>