<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j
Dim iNumYears
Dim sWhichGraph, sLineColor, sThisDay, sThisMonth
Dim sngMaxY, sngMinY, sngThisVal, sngProfit, sngIncome, sngExpenses, sngMargin
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
sngMinY = 100000

iNumYears = Year(Date) - CInt(2013) + 1

ReDim Years(iNumYears)

For i = 0 To UBound(Years) - 1
    Years(i) = i + 2013 'which year
Next

sThisDay = Day(Date)
sThisMonth = Month(Date)

Private Sub GetThisData(iThisMonth, iThisYear)
    Select Case sWhichGraph
        Case "Events"
            Call GetEvents(iThisMonth, iThisYear)

            If UBound(Events, 2) > 0 Then sngThisVal = CSng(sngThisVal) + CSng(UBound(Events, 2) - 1)
        Case "Income"
            Call GetEvents(iThisMonth, iThisYear)

            sngIncome = 0

            Call GetIncome(iThisMonth, iThisYear)

            sngThisVal = CSng(sngThisVal) + CSng(sngIncome)
        Case "Profit"
            Call GetEvents(iThisMonth, iThisYear)

            sngIncome = 0
            sngExpenses = 0

            Call GetExpenses(iThisMonth, iThisYear)
            Call GetIncome(iThisMonth, iThisYear)

            sngProfit = CSng(sngIncome) - CSng(sngExpenses)
            sngThisVal = CSng(sngThisVal) + CSng(sngProfit)
        Case "Expenses"
            Call GetEvents(iThisMonth, iThisYear)

            sngExpenses = 0

            Call GetExpenses(iThisMonth, iThisYear)

            sngThisVal = CSng(sngThisVal) + CSng(sngExpenses)
        Case "Margin"
            Call GetEvents(iThisMonth, iThisYear)

            sngMargin = 0
            sngIncome = 0
            sngExpenses = 0

            Call GetMargin(iThisMonth, iThisYear)

            sngThisVal = CSng(sngMargin)
    End Select

    If CSng(sngThisVal) > CSng(sngMaxY) Then sngMaxY = sngThisVal
    If CSng(sngThisVal) < CSng(sngMinY) Then sngMinY = sngThisVal
End Sub

Private Sub GetEvents(sCurrMonth, sCurrYear)
    Dim y

    ReDim Events(2, 0)
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

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetsID, MeetDate, Sport FROM Meets WHERE (MeetDate >= '1/1/" & sCurrYear 
    sql = sql & "' AND MeetDate <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "') ORDER BY MeetDate"
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
End Sub

Private Sub GetExpenses(sCurrMonth, sCurrYear)
    Dim x

    sngExpenses = 0

    For x = 0 To UBound(Events, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Staffing, MiscCost, PartCost, LaborCost, Mileage FROM FinanceEvents WHERE EventID = " & Events(0, x)
        sql = sql & " AND Sport = '" & Events(1, x) & "'"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Month(CDate(Events(2, x))) = CInt(sCurrMonth) Then 
                sngExpenses = CSng(sngExpenses) + CSng(rs(0).Value) + CSng(rs(1).Value) + CSng(rs(2).Value) + CSng(rs(3).Value) + CSng(rs(4).Value)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtPaid, WhenPaid FROM FinanceExpense WHERE WhenPaid >= '1/1/" & sCurrYear 
    sql = sql & "' AND WhenPaid <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Month(CDate(rs(1).Value)) = CInt(sCurrMonth) Then sngExpenses = CSng(sngExpenses) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetIncome(sCurrMonth, sCurrYear)
    Dim x

    For x = 0 To UBound(Events, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Invoice FROM FinanceEvents WHERE EventID = " & Events(0, x) & " AND Sport = '" & Events(1, x) & "'"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Month(CDate(Events(2, x))) = CInt(sCurrMonth) Then sngIncome = CSng(sngIncome) + CSng(rs(0).Value)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next

    'get regular income
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtRcvd, WhenRcvd FROM FinanceIncome WHERE EventID = 0 AND (WhenRcvd >= '1/1/" & sCurrYear 
    sql = sql & "' AND WhenRcvd <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Month(CDate(rs(1).Value)) = Int(sCurrMonth) Then sngIncome = CSng(sngIncome) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetMargin(sCurrMonth, sCurrYear)
    Dim x

    For x = 0 To UBound(Events, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Staffing, MiscCost, PartCost, LaborCost, Mileage, Invoice FROM FinanceEvents WHERE EventID = " & Events(0, x)
        sql = sql & " AND Sport = '" & Events(1, x) & "'"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If Month(CDate(Events(2, x))) <= CInt(sCurrMonth) Then 
                sngExpenses = CSng(sngExpenses) + CSng(rs(0).Value) + CSng(rs(1).Value) + CSng(rs(2).Value) + CSng(rs(3).Value) + CSng(rs(4).Value)
                sngIncome = CSng(sngIncome) + CSng(rs(5).Value)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtPaid, WhenPaid FROM FinanceExpense WHERE WhenPaid >= '1/1/" & sCurrYear 
    sql = sql & "' AND WhenPaid <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Month(CDate(rs(1).Value)) <= CInt(sCurrMonth) Then sngExpenses = CSng(sngExpenses) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get regular income
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtRcvd, WhenRcvd FROM FinanceIncome WHERE EventID = 0 AND (WhenRcvd >= '1/1/" & sCurrYear 
    sql = sql & "' AND WhenRcvd <= '" & sThisMonth & "/" & sThisDay & "/" & sCurrYear & "')"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Month(CDate(rs(1).Value)) <= CInt(sCurrMonth) Then sngIncome = CSng(sngIncome) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

If sCurrYear = "2018" Then Response.Write sngIncome & " - " & sngExpenses & " = " & sngIncome - sngExpenses & "<br>"

    If CSng(sngIncome) > 0 Then sngMargin = Round((CSng(sngIncome) - CSng(sngExpenses))/CSng(sngIncome), 4)*100
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
        If j <= CInt(sThisMonth) Then
	        Chart.AddPoint j, sngThisVal, sLineColor, Years(i)
'            Chart.AddLineGraphText sngThisVal, j, sngThisVal, 0
        End If

        If j = 12 Then sngThisVal = 0
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