<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID, lEventGrp, lThisEvent
Dim sEventRaces
Dim sngMaxY
Dim Finishers()
Dim Chart

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
If lEventID = vbNullString Then lEventID = "0"

j = 0
ReDim Years(0)
For i = 2002 To Year(Date)
    Years(j) = i
    j = j + 1
    ReDim Preserve Years(j)
Next

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sngMaxY = 0

ReDim Finishers(1, UBound(Years) - 1)

If CLng(lEventID) > 0 Then
    'get eventgrp
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventGrp FROM Events WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    lEventGrp = rs(0).Value
    rs.Close
    Set rs = Nothing

    'get event data by year
    For j = 0 To UBound(Years) - 1
        Finishers(0, j) = Years(j)
        Finishers(1, j) = "0"

        lThisEvent = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventID, EventDate FROM Events WHERE EventGrp = " & lEventGrp
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If rs.RecordCount > 0 Then 
                If Year(CDate(rs(1).Value)) = Years(j) Then
                    sEventRaces = GetEventRaces(rs(0).Value)
                    Finishers(1, j) = NumFinishers(rs(0).Value, rs(1).Value, sEventRaces)
                    Exit Do 
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next
End If

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

Private Function NumFinishers(lThisEvent, dEventDate, sEventRaces)
    Dim rs2, sql2
    
    NumFinishers = 0

    If CDate(dEventDate) < Date Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ")"
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then NumFinishers = rs2.RecordCount
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Event Trends Graph</title>
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
Chart.XOffset = 2002
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
Chart.Title = "Event Trends"
Chart.TitleX = 75
Chart.TitleY = 50
Chart.TitleSize = 10
Chart.TitleBold = true
Chart.TitleColor = "000000"

'show ind graph
If UBound(Finishers, 2) > 0 Then
    For i = 0 To UBound(Finishers, 2)
        Chart.AddPoint Finishers(0, i), Finishers(1, i), "ff0000", "Finishers"
        Chart.AddLineGraphText Finishers(1, i), Finishers(0, i), Finishers(1, i), 0
    Next
End If

Response.ContentType = "Image/Gif"
Response.BinaryWrite Chart.GIFLine

Response.End
%>
</body>
</html>