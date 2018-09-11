<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j
Dim iNumYears
Dim sWhichGraph, sLineColor
Dim sngMaxY, sngThisVal
Dim Years(), Events()
Dim Chart

sWhichGraph = Request.QueryString("which_graph")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sngMaxY = 0
iNumYears = Year(Date) - CInt(2013) + 1

ReDim Years(iNumYears)

For i = 0 To UBound(Years) - 1
    Years(i) = i + 2013 'which year
Next

Private Sub GetThisData(iThisMonth, iThisYear)
    Dim x
    Dim sngIncome, sngExpense

    sngThisVal = 0

    Select Case sWhichGraph
        Case "Events"
            'fitness event count
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT EventDate FROM Events WHERE (EventDate >= '1/1/" & iThisYear & "' AND EventDate <= '12/31/" & iThisYear & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If Month(rs(0).Value) = iThisMonth Then sngThisVal = CSng(sngThisVal) + 1
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            'cc/nordic event count
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MeetDate FROM Meets WHERE (MeetDate >= '1/1/" & iThisYear & "' AND MeetDate <= '12/31/" & iThisYear & "')"
            rs.Open sql, conn2, 1, 2
            Do While Not rs.EOF
                If Month(rs(0).Value) = iThisMonth Then sngThisVal = CSng(sngThisVal) + 1
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Case "Income"
            Call GetEvents(iThisMonth, iThisYear)

            'get event expense (income is accounted for in the financeexpense table)
            For x = 0 To UBound(Events, 2) - 1
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT Invoice FROM FinanceEvents WHERE EventID = " & Events(0, x) & " AND Sport = '" & Events(1, x) & "'"
                rs.Open sql, conn, 1, 2
                Do While Not rs.EOF
                    sngThisVal = CSng(sngThisVal) + CSng(rs(0).Value)
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
            Next

            'get regular income
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT AmtRcvd, WhenRcvd FROM FinanceIncome WHERE EventID = 0 AND (WhenRcvd >= '1/1/" & iThisYear & "' AND WhenRcvd <= '12/31/" & iThisYear & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If Month(CDate(rs(1).Value)) = Month(CDate(iThisMonth)) Then sngThisVal = CSng(sngThisVal) + CSng(rs(0).Value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Case "Profit"
            Call GetEvents(iThisMonth, iThisYear)

            'get event expense (income is accounted for in the financeexpense table)
            For x = 0 To UBound(Events, 2) - 1
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT Invoice, Staffing, MiscCost, PartCost, LaborCost, Mileage FROM FinanceEvents WHERE EventID = " & Events(0, x)
                sql = sql & " AND Sport = '" & Events(1, x) & "'"
                rs.Open sql, conn, 1, 2
                Do While Not rs.EOF
                    sngIncome = CSng(sngIncome) + CSng(rs(0).Value)
                    sngExpense = CSng(sngExpense) + CSng(rs(1).Value) + CSng(rs(2).Value) + CSng(rs(3).Value) + CSng(rs(4).Value) + CSng(rs(5).Value)
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
            Next

            'get regular income
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT AmtRcvd, WhenRcvd FROM FinanceIncome WHERE EventID = 0 AND (WhenRcvd >= '1/1/" & iThisYear & "' AND WhenRcvd <= '12/31/" & iThisYear & "')"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If Month(CDate(rs(1).Value)) = Month(CDate(iThisMonth)) Then sngIncome = CSng(sngIncome) + CSng(rs(0).Value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            'get regular expense
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT AmtPaid, WhenPaid FROM FinanceExpense WHERE WhenPaid >= '1/1/" & iThisYear & "' AND WhenPaid <= '12/31/" & iThisYear & "'"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If Month(CDate(rs(1).Value)) = Month(CDate(iThisMonth)) Then sngExpense = CSng(sngExpense) + CSng(rs(0).Value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            'calculate profit
            sngThisVal = CSng(sngIncome) - CSng(sngExpense)
        Case "Expenses"
            Call GetEvents(iThisMonth, iThisYear)

            'get event expense (income is accounted for in the financeexpense table)
            For x = 0 To UBound(Events, 2) - 1
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT Staffing, MiscCost, PartCost, LaborCost, Mileage FROM FinanceEvents WHERE EventID = " & Events(0, x)
                sql = sql & " AND Sport = '" & Events(1, x) & "'"
                rs.Open sql, conn, 1, 2
                Do While Not rs.EOF
                    sngThisVal = CSng(sngThisVal) + CSng(rs(0).Value) + CSng(rs(1).Value) + CSng(rs(2).Value) + CSng(rs(3).Value) + CSng(rs(4).Value)
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
            Next

            'get regular expense
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT AmtPaid, WhenPaid FROM FinanceExpense WHERE WhenPaid >= '1/1/" & iThisYear & "' AND WhenPaid <= '12/31/" & iThisYear & "'"
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If Month(CDate(rs(1).Value)) = Month(CDate(iThisMonth)) Then sngThisVal = CSng(sngThisVal) + CSng(rs(0).Value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Case "Margin"
    End Select

    If CSng(sngThisVal) > CSng(sngMaxY) Then sngMaxY = sngThisVal
End Sub

Private Sub GetEvents(iEventMonth, iEventYear)
    Dim y

    ReDim Events(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventDate FROM Events WHERE (EventDate >= '1/1/" & iEventYear & "' AND EventDate <= '12/31/" & iEventYear & "')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Month(rs(1).Value) = iEventMonth Then
            Events(0, y) = rs(0).Value
            Events(1, y) = "Fitness Event"
            y = y + 1
            ReDim Preserve Events(1, y)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetsID, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & iEventYear & "' AND MeetDate <= '12/31/" & iEventYear & "')"
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        If Month(rs(1).Value) = iEventMonth Then
            Events(0, y) = rs(0).Value
            Events(1, y) = rs(2).Value
            y = y + 1
            ReDim Preserve Events(1, y)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Finance Graphs: Monthly Graph</title>
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
Chart.Title = sWhichGraph
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
	    Chart.AddPoint j, sngThisVal, sLineColor, Years(i)
'        If sWhichGraph = "Events" Then  'just take out the $ or %
'            Chart.AddLineGraphText sngThisVal, j, sngThisVal, 0
'        Else
'            Chart.AddLineGraphText "$" & sngThisVal, j, sngThisVal, 0
'        End If
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