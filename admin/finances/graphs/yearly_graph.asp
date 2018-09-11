<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j
Dim iNumFitness, iNumNordic, iNumCC, iNumEvents, iNumYrs
Dim sWhichGraph, sShowAvg
Dim sngMaxY
Dim Fitness(), Nordic(), CC(), Total(), Other(), Events(), Years(), Sports(3)
Dim Chart

sWhichGraph = Request.QueryString("which_graph")
sShowAvg = Request.QueryString("show_avg")

Sports(0) = "Fitness Event"
Sports(1) = "Nordic Ski"
Sports(2) = "Cross-Country"
Sports(3) = "Other"

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
ReDim Other(1, iNumYrs)
ReDim Events(1, iNumYrs)

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

    Other(0, j) = Years(j)
    Other(1, j) = "0"

    iNumEvents = 0
    iNumFitness = 0
    iNumNordic = 0
    iNumCC = 0

    If sWhichGraph = "Margin" Then
        Fitness(1, j) = GetMargin("Fitness Event", Years(j))
        Nordic(1, j) = GetMargin("Nordic Ski", Years(j))
        CC(1, j) = GetMargin("Cross-Country", Years(j))
        Total(1, j) = GetMargin("Total", Years(j))

        sngMaxY = "50"
    ElseIf sWhichGraph = "Staff" Then
        Fitness(1, j) = GetStaff("Fitness Event", Years(j))
        If CSng(Fitness(1, j)) > CSng(sngMaxY) Then sngMaxY = Fitness(1, j)

        Nordic(1, j) = GetStaff("Nordic Ski", Years(j))
        If CSng(Nordic(1, j)) > CSng(sngMaxY) Then sngMaxY = Nordic(1, j)

        CC(1, j) = GetStaff("Cross-Country", Years(j))
        If CSng(CC(1, j)) > CSng(sngMaxY) Then sngMaxY = CC(1, j)

        Total(1, j) = GetStaff("Total", Years(j))
        If CSng(Total(1, j)) > CSng(sngMaxY) Then sngMaxY = Total(1, j)
    Else
        'just get num events
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventID FROM Events WHERE (EventDate >= '1/1/" & Years(j) & "' AND EventDate <= '12/31/" & Years(j) & "')"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then iNumFitness = rs.RecordCount
        rs.Close
        Set rs = Nothing

        Set rs = Server.CreateObject("ADODB.Recordset")
 '       If Date < CDate("12/31/" & Years(j)) Then
 '           sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & Years(j) & "' AND EventDate <= '" & Date & "')"
 '       Else
            sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & Years(j) & "' AND EventDate <= '12/31/" & Years(j) & "')"
 '       End If
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Not sWhichGraph = "Events" Then 'don't call EventVal if we just want a count of the number of events...because that includes through the year
                Fitness(1, j) = CSng(Fitness(1, j)) + EventVal(rs(0).Value, "Fitness Event", rs(1).Value)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If sWhichGraph = "Events" Then
            Fitness(1, j) = iNumFitness
        ELse
            If sShowAvg = "y" AND iNumFitness > 0 Then Fitness(1, j) = Round(CSng(Fitness(1, j))/CInt(iNumFitness), 2)
        End If

        Total(1, j) = CSng(Fitness(1, j))

        'just get num events
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Nordic Ski'"
        rs.Open sql, conn2, 1, 2
        If rs.RecordCount > 0 Then iNumNordic = rs.RecordCount
        rs.Close
        Set rs = Nothing

        'get data per event
        Set rs = Server.CreateObject("ADODB.Recordset")
'        If Date < CDate("12/31/" & Years(j)) Then
 '           sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '" & DAte & "') AND Sport = 'Nordic Ski'"
 '       Else
            sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Nordic Ski'"
 '       End If
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If Not sWhichGraph = "Events" Then 'don't call EventVal if we just want a count of the number of events
                Nordic(1, j) = CSng(Nordic(1, j)) + EventVal(rs(0).Value, "Nordic Ski", rs(1).Value)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If sWhichGraph = "Events" Then
            Nordic(1, j) = iNumNordic
        ELse
            If sShowAvg = "y" AND iNumNordic > 0 Then Nordic(1, j) = Round(CSng(Nordic(1, j))/CInt(iNumNordic), 2)
        End If

        Total(1, j) = CSng(Total(1, j)) + CSng(Nordic(1, j))


        If sShowAvg = "y" AND iNumNordic > 0 Then Nordic(1, j) = Round(CSng(Nordic(1, j))/CInt(iNumNordic), 2)

        'just get num events
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetsID FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Cross-Country'"
        rs.Open sql, conn2, 1, 2
        If rs.RecordCount > 0 Then iNumCC = rs.RecordCount
        rs.Close
        Set rs = Nothing

        Set rs = Server.CreateObject("ADODB.Recordset")
 '       If Date < CDate("12/31/" & Years(j)) Then
 '           sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '" & DAte & "') AND Sport = 'Cross-Country'"
'        Else
            sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & Years(j) & "' AND MeetDate <= '12/31/" & Years(j) & "') AND Sport = 'Cross-Country'"
'        End If
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If Not sWhichGraph = "Events" Then 'don't call EventVal if we just want a count of the number of events
                CC(1, j) = CSng(CC(1, j)) + EventVal(rs(0).Value, "Cross-Country", rs(1).Value)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If sWhichGraph = "Events" Then
            CC(1, j) = iNumCC
        ELse
            If sShowAvg = "y" AND iNumCC > 0 Then CC(1, j) = Round(CSng(CC(1, j))/CInt(iNumCC), 2)
        End If

        Total(1, j) = CSng(Total(1, j)) + CSng(CC(1, j))

        If sShowAvg = "y" AND iNumCC > 0 Then CC(1, j) = Round(CSng(CC(1, j))/CInt(iNumCC), 2)

        iNumEvents = CInt(iNumFitness) + CInt(iNumNordic) + CInt(iNumCC)

        Total(0, j) = Years(j)
        If sShowAvg = "y" And iNumEvents > 0 Then Total(1, j) = Round(CSng(Total(1, j))/CInt(iNumEvents), 2)

        If sWhichGraph = "Income" Then
        'now get other income for the year
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT AmtRcvd FROM FinanceIncome WHERE IncomeType IN ('Crowd Torch', 'AdSense', 'Tempo Events', 'Misc Income') AND (WhenRcvd >= '1/1/" 
            sql = sql & Years(j) & "' AND WhenRcvd <= '12/31/" & Years(j) & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                Other(1, j) = CSng(Other(1, j)) + CSng(rs(0).Value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If

        'reset max y for graph axis
        If CSng(Total(1, j)) > CSng(sngMaxY) Then sngMaxY = Total(1, j)
'        If CSng(CC(1, j)) > CSng(sngMaxY) Then sngMaxY = CC(1, j)
'        If CSng(Nordic(1, j)) > CSng(sngMaxY) Then sngMaxY = Nordic(1, j)
'        If CSng(Fitness(1, j)) > CSng(sngMaxY) Then sngMaxY = Fitness(1, j)
'        If CSng(Other(1, j)) > CSng(sngMaxY) Then sngMaxY = Other(1, j)
    End If
Next

Private Function EventVal(lThisEvent, sThisSport, dEventDate)
    Dim rs2, sql2
    Dim sngIncome, sngExpense, sngProfit
    
    EventVal = 0
    sngIncome = 0
    sngExpense = 0    
    sngProfit = 0    

    If CDate(dEventDate) <= Date Then
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT Invoice, Staffing, MiscCost, PartCost, LaborCost, Mileage FROM FinanceEvents WHERE Sport = '" & sThisSport 
        sql2 = sql2 & "' AND EventID = " & lThisEvent
        rs2.Open sql2, conn, 1, 2
        If rs2.RecordCount > 0 Then
            sngIncome = CSng(sngIncome) + CSng(rs2(0).Value)
            sngExpense = CSng(sngExpense) + CSng(rs2(1).Value) + CSng(rs2(2).Value) + CSng(rs2(3).Value) + CSng(rs2(4).Value) + CSng(rs2(5).Value)
        End If
        rs2.Close
        Set rs2 = Nothing

        sngProfit = CSng(sngIncome) - CSng(sngExpense)

        Select Case sWhichGraph
            Case "Income"
                EventVal = Round(sngIncome, 2)
            Case "Expenses"
                EventVal = Round(sngExpense, 2)
            Case "Profit"
                EventVal = Round(sngProfit, 2)
        End Select
    End If
End Function

Private Function GetMargin(sSport, iYear)
    Dim x
    Dim sngProfit, sngIncome, sngExpense
    Dim Events()

    GetMargin = "0"

    sngIncome = "0"
    sngExpense = "0"
    sngProfit = "0"

    'get events for this sport for this year
    x = 0
    ReDim Events(1, 0)
    Select Case sSport
        Case "Fitness Event"
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Fitness Event"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Case "Total"
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Fitness Event"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "') AND Sport = '"
            sql = sql & "Nordic Ski'"
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Nordic Ski"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing    

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "') AND Sport = '"
            sql = sql & "Cross-Country'"
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Cross-Country"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing    
        Case Else
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MeetsID, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "')"
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date And rs(2).Value = sSport Then
                    Events(0, x) = sSport
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing    
    End Select

    'get income and expense
    For x = 0 To UBound(Events, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Invoice, Staffing, MiscCost, PartCost, LaborCost, Mileage FROM FinanceEvents WHERE Sport = '" & Events(0, x) & "' "
        sql = sql & "AND EventID = " & Events(1, x)
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            sngIncome = CSng(sngIncome) + CSng(rs(0).Value)
            sngExpense = CSng(sngExpense) + CSng(rs(1).Value) + CSng(rs(2).Value) + CSng(rs(3).Value) + CSng(rs(4).Value) + CSng(rs(5).Value)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next

    'calculate profit and margin
    sngProfit = CSng(sngIncome) - CSng(sngExpense)
    If CSng(sngIncome) > 0 Then GetMargin = Round(CSng(sngProfit)/CSng(sngIncome), 4)*100
End Function

Private Function GetStaff(sSport, iYear)
    Dim x
    Dim sngStaff
    Dim Events()

    GetStaff = "0"
    sngStaff = "0"

    'get events for this sport for this year
    x = 0
    ReDim Events(1, 0)
    Select Case sSport
        Case "Fitness Event"
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & iYear 
            sql = sql & "' AND EventDate <= '12/31/" & iYear & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Fitness Event"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Case "Total"
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Fitness Event"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "') AND Sport = '"
            sql = sql & "Nordic Ski'"
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Nordic Ski"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing    

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MeetsID, MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "') AND Sport = '"
            sql = sql & "Cross-Country'"
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date Then
                    Events(0, x) = "Cross-Country"
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing    
        Case Else
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MeetsID, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "')"
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                If CDate(rs(1).Value) <= Date And rs(2).Value = sSport Then
                    Events(0, x) = sSport
                    Events(1, x) = rs(0).Value
                    x = x + 1
                    ReDim Preserve Events(1, x)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing    
    End Select

    'get staff cost
    For x = 0 To UBound(Events, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Staffing FROM FinanceEvents WHERE Sport = '" & Events(0, x) & "' AND EventID = " & Events(1, x)
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            sngStaff = CSng(sngStaff) + CSng(rs(0).Value)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        'subtract my payments
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT TransAmt FROM FinanceStaff WHERE StaffID = 1 AND Sport = '" & Events(0, x) & "' AND EventID = " & EVents(1, x)
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            sngStaff = CSng(sngStaff) - CSng(rs(0).Value)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next

    GetStaff = sngStaff
    If sShowAvg = "y" Then 
        If CSng(UBound(Events, 2) - 1) > 0 Then GetStaff = Round(CSng(sngStaff)/(UBound(Events, 2) - 1), 0)
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
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Finance Graphs: Yearly Graph</title>
<!--#include file = "../../../includes/js.asp" -->
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
Chart.YAxisText = "$"
Chart.LegendX = 650
Chart.LegendY = 50
Chart.ShowLegend = true
Chart.ShowLegendBox = true
Chart.LegendTextSize = 9
Chart.Title = sWhichGraph
Chart.TitleX = 75
Chart.TitleY = 50
Chart.TitleSize = 10
Chart.TitleBold = true
Chart.TitleColor = "000000"

'show ind graph
For i = 0 To UBound(Fitness, 2)
	Chart.AddPoint Fitness(0, i), Fitness(1, i), "ff0000", "Fitness Events"
    If sWhichGraph = "Margin" Then
        Chart.AddLineGraphText Fitness(1, i) & "%", Fitness(0, i), Fitness(1, i), 0
    ElseIf sWhichGraph = "Events" Then  'just take out the $ or %
        Chart.AddLineGraphText Fitness(1, i), Fitness(0, i), Fitness(1, i), 0
    Else
        Chart.AddLineGraphText "$" & Fitness(1, i), Fitness(0, i), Fitness(1, i), 0
    End If

    Chart.AddPoint Nordic(0, i), Nordic(1, i), "003399", "Nordic Ski"
    If sWhichGraph = "Margin" Then
        Chart.AddLineGraphText Nordic(1, i) & "%", Nordic(0, i), Nordic(1, i), 0
    ElseIf sWhichGraph = "Events" Then  'just take out the $ or %
        Chart.AddLineGraphText Nordic(1, i), Nordic(0, i), Nordic(1, i), 0
    Else
        Chart.AddLineGraphText "$" & Nordic(1, i), Nordic(0, i), Nordic(1, i), 0
    End If

    Chart.AddPoint CC(0, i), CC(1, i), "009933", "Cross-Country"
    If sWhichGraph = "Margin" Then
        Chart.AddLineGraphText CC(1, i) & "%", CC(0, i), CC(1, i), 0
    ElseIf sWhichGraph = "Events" Then  'just take out the $ or %
        Chart.AddLineGraphText CC(1, i), CC(0, i), CC(1, i), 0
    Else
        Chart.AddLineGraphText "$" & CC(1, i), CC(0, i), CC(1, i), 0
    End If

    If sWhichGraph = "Income" Then
        Chart.AddPoint Other(0, i), Other(1, i), "800000", "Other"
        Chart.AddLineGraphText "$" & Total(1, i), Total(0, i), Total(1, i), 0
    End If

    Chart.AddPoint Total(0, i), Total(1, i), "000000", "Total"
    If sWhichGraph = "Margin" Then
        Chart.AddLineGraphText Total(1, i) & "%", Total(0, i), Total(1, i), 0
    ElseIf sWhichGraph = "Events" Then  'just take out the $ or %
        Chart.AddLineGraphText Total(1, i), Total(0, i), Total(1, i), 0
    Else
        Chart.AddLineGraphText "$" & Total(1, i), Total(0, i), Total(1, i), 0
    End If
Next

Response.ContentType = "Image/Gif"
Response.BinaryWrite Chart.GIFLine

Response.End
%>
</body>
</html>