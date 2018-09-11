<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j
Dim iNumFitness, iNumNordic, iNumCC, iNumYrs
Dim sEventRaces, sShowAvg
Dim sngMaxY
Dim Fitness(), Nordic(), CC(), Total(), Years(), Sports(2)
Dim Chart

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

sShowAvg = Request.QueryString("show_avg")
If sShowAvg = vbNullString Then sShowAvg = "n"

Sports(0) = "Fitness Event"
Sports(1) = "Nordic Ski"
Sports(2) = "Cross-Country"

j = 0
ReDim Years(0)
For i = 2013 To Year(Date)
    Years(j) = i
    j = j + 1
    ReDim Preserve Years(j)
Next

iNumYrs = UBound(Years) - 1

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

ReDim Fitness(1, iNumYrs) 
ReDim Nordic(1, iNumYrs)
ReDim CC(1, iNumYrs)
ReDim Total(1, iNumYrs)

sngMaxY = 0

'get event data by year
For j = 0 To iNumYrs
    Fitness(0, j) = Years(j)
    Fitness(1, j) = "0"

    Nordic(0, j) = Years(j)
    Nordic(1, j) = "0"

    CC(0, j) = Years(j)
    CC(1, j) = "0"

    Total(0, j) = Years(j)
    Total(1, j) = "0"

    iNumFitness = 0
    iNumNordic = 0
    iNumCC = 0

    'just get num events
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Date < CDate("12/31/" & Years(j)) Then   'to prevent events later this year from skewing data
       sql = "SELECT EventID FROM Events WHERE (EventDate >= '1/1/" & Years(j) & "' AND EventDate <= '" & Date & "')"
    Else
        sql = "SELECT EventID FROM Events WHERE (EventDate >= '1/1/" & Years(j) & "' AND EventDate <= '12/31/" & Years(j) & "')"
    End If
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iNumFitness = rs.RecordCount
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    If Date < CDate("12/31/" & Years(j)) Then
        sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & Years(j) & "' AND EventDate <= '" & Date & "')"
    Else
        sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & Years(j) & "' AND EventDate <= '12/31/" & Years(j) & "')"
    End If
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sEventRaces = GetEventRaces(rs(0).Value)
        If Not sEventRaces = vbNullString Then
            Fitness(1, j) = CSng(Fitness(1, j)) + NumFitness(rs(0).Value, rs(1).Value, sEventRaces)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Total(1, j) = CSng(Fitness(1, j))

    If sShowAvg = "y" Then 
        If CInt(iNumFitness) > 0 Then Fitness(1, j) = Round(CSng(Fitness(1, j))/CSng(iNumFitness), 2)
    End If

    'just get num events
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Date < CDate("12/31/" & Years(j)) Then   'to prevent events later this year from skewing data
       sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '" & Date & "') AND Sport = 'Nordic Ski'"
    Else
       sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Nordic Ski'"
    End If
    rs.Open sql, conn2, 1, 2
    If rs.RecordCount > 0 Then iNumNordic = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get data per event
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Date < CDate("12/31/" & Years(j)) Then
        sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '" & DAte & "') AND Sport = 'Nordic Ski'"
    Else
        sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Nordic Ski'"
    End If
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        Nordic(1, j) = CSng(Nordic(1, j)) + NumSchool(rs(0).Value, rs(1).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Total(1, j) = CSng(Total(1, j)) + CSng(Nordic(1, j))

    If sShowAvg = "y" Then 
        If CInt(iNumNordic) > 0 Then Nordic(1, j) = Round(CSng(Nordic(1, j))/CSng(iNumNordic), 2)
    End If

    'just get num events
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Date < CDate("12/31/" & Years(j)) Then   'to prevent events later this year from skewing data
       sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '" & Date & "') AND Sport = 'Cross-Country'"
    Else
       sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Cross-Country'"
    End If
    rs.Open sql, conn2, 1, 2
    If rs.RecordCount > 0 Then iNumCC = rs.RecordCount
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    If Date < CDate("12/31/" & Years(j)) Then
        sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '" & DAte & "') AND Sport = 'Cross-Country'"
    Else
        sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Cross-Country'"
    End If
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        CC(1, j) = CSng(CC(1, j)) + NumSchool(rs(0).Value, rs(1).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Total(0, j) = Years(j)
    Total(1, j) = CSng(Total(1, j)) + CSng(CC(1, j))

    If sShowAvg = "y" Then 
        If CInt(iNumCC) > 0 Then CC(1, j) = Round(CSng(CC(1, j))/CSng(iNumCC), 2)
        Total(1, j) = Round(CSng(Total(1, j))/(CInt(iNumFitness) + CInt(iNumNordic) + CInt(iNumCC)), 2)
    End If
Next

Private Function GetEventRaces(lThisEvent)
    Dim rs2, sql2
    
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

Private Function NumFitness(lThisEvent, dEventDate, sEventRaces)
    Dim rs2, sql2
    
    NumFitness = 0

    If CDate(dEventDate) < Date Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ")"
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then NumFitness = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

Private Function NumSchool(lThisEvent, dEventDate)
    Dim rs2, sql2
    
    NumSchool = 0

    If CDate(dEventDate) < Date Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT IndRsltsID FROM IndRslts WHERE MeetsID = " & lThisEvent
        rs2.Open sql2, conn2, 1, 2
        If rs2.RecordCount > 0 Then NumSchool = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Participation Graph</title>
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
Chart.XOffset = 2013
Chart.XTop = 2018
Chart.YOffset = 0
Chart.YTop = sngMaxY
Chart.ShowGrid = true
Chart.XGrad = 1
Chart.YGrad = 0
Chart.XAxisText = "Year"
Chart.YAxisText = "Participants"
Chart.LegendX = 650
Chart.LegendY = 50
Chart.ShowLegend = true
Chart.ShowLegendBox = true
Chart.LegendTextSize = 9
Chart.Title = "Participation Trends"
Chart.TitleX = 75
Chart.TitleY = 50
Chart.TitleSize = 10
Chart.TitleBold = true
Chart.TitleColor = "000000"

'show ind graph
For i = 0 To UBound(Fitness, 2)
	Chart.AddPoint Fitness(0, i), Fitness(1, i), "ff0000", "Fitness Events"
    Chart.AddLineGraphText Fitness(1, i), Fitness(0, i), Fitness(1, i), 0

    Chart.AddPoint Nordic(0, i), Nordic(1, i), "003399", "Nordic Ski"
    Chart.AddLineGraphText Nordic(1, i), Nordic(0, i), Nordic(1, i), 0

    Chart.AddPoint CC(0, i), CC(1, i), "009933", "Cross-Country"
    Chart.AddLineGraphText CC(1, i), CC(0, i), CC(1, i), 0

    Chart.AddPoint Total(0, i), Total(1, i), "000000", "Total"
    Chart.AddLineGraphText Total(1, i), Total(0, i), Total(1, i), 0
Next

Response.ContentType = "Image/Gif"
Response.BinaryWrite Chart.GIFLine

Response.End
%>
</body>
</html>