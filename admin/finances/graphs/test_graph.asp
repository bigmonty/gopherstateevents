<%@ Language=VBScript%>
<!DOCTYPE html>
<html lang="en">
<head>
</head>
<body>
<%
Response.Expires = 0
Response.Buffer = true
Response.Clear
Set Chart = Server.CreateObject("csDrawGraph64.Draw")
Chart.ShowGrid = true
Chart.YAxisText = "Y Axis"
Chart.XAxisText = "X Axis"
Chart.AxisTextBold = true
Chart.AddPoint 0, 0, "ff0000", "Red Line"
Chart.AddPoint 30, 30, "ff0000", ""
Chart.AddPoint 0, 0, "00ff00", "Green Line"
Chart.AddPoint 30, 20, "00ff00", ""
Response.ContentType = "image/gif"
Response.BinaryWrite Chart.GIFLine
%>
</body>
</html>