<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k, m
Dim SeriesMeets(), SeriesRaces(), SeriesRslts(), TempArr()
Dim lSeriesID
Dim sSeriesName, sGender, sThisGender
Dim iMaxScore, iDifference, iScaleFactor, iTeamPlace
Dim bExitFor

iScaleFactor = 1
iMaxScore = 25
iDifference = 5

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect("http://www.google.com")
If CLng(lSeriesID) < 0 Then Response.Redirect("http://www.google.com")

sGender = Request.QueryString("gender")

If sGender = "M" Then
	sThisGender = "Boys"
Else
	sThisGender = "Girls"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
Set rs = conn.Execute(sql)
sSeriesName = Replace(rs(0).Value, "''", "'")
Set rs = Nothing

'get meets in this series
i = 0
ReDim SeriesMeets(2, 0)
sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate FROM Meets m INNER JOIN SeriesMeets sm ON m.MeetsID = sm.MeetsID "
sql = sql & "WHERE sm.SeriesID = " & lSeriesID & " ORDER BY m.MeetDate"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	SeriesMeets(0, i) = rs(0).Value
	SeriesMeets(1, i) = Left(Replace(rs(1).Value, "''", "'"), 11)
	SeriesMeets(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve SeriesMeets(2, i)
	rs.MoveNext
Loop
Set rs = Nothing
 
'get all races in this series
i = 0
ReDim SeriesRaces(0)
sql = "SELECT RacesID FROM SeriesMeets WHERE SeriesID = " & lSeriesID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	SeriesRaces(i) = rs(0).Value
	i = i + 1
	ReDim Preserve SeriesRaces(i)
	rs.MoveNext
Loop
Set rs = Nothing

'get all teams in this series
i = 0
ReDim SeriesRslts(UBound(SeriesMeets, 2) + 4, 0)
sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt ON t.TeamsID = mt.TeamsID "
sql = sql & "WHERE t.Gender = '" & sGender & "' AND mt.MeetsID = " & SeriesMeets(0, 0)
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	SeriesRslts(0, i) = rs(0).Value
	SeriesRslts(1, i) = Replace(rs(1).Value, "''", "'")
	For j = 2 To UBound(SeriesMeets, 2) + 4
		SeriesRslts(j, i) = 0
	Next
	i = i + 1
	ReDim Preserve SeriesRslts(UBound(SeriesMeets, 2) + 4, i)
	rs.MoveNext
Loop
Set rs = Nothing

'now get the scores for these teams for this series to date
For j = 0 To UBound(SeriesRaces) - 1
	bExitFor = False

	If CLng(SeriesRaces(j)) = 281 Or CLng(SeriesRaces(j)) = 282 Then iScaleFactor = 2
	
	sql = "SELECT m.MeetDate FROM Meets m INNER JOIN Races r ON r.MeetsID = m.MeetsID "
	sql = sql & "WHERE r.RacesID = " & SeriesRaces(j)
	Set rs = conn.Execute(sql)
	If Date < CDate(rs(0).Value) Then bExitFor = True
	Set rs = Nothing

	If bExitFor = True Then Exit For

	'get the scores into the array
	For i = 0 To UBound(SeriesRslts, 2) - 1
		'for each team get their points for this meet
		sql = "SELECT Score FROM TmRslts WHERE RacesID = " & SeriesRaces(j) & " AND TeamsID = " & SeriesRslts(0, i)
		Set rs = conn.Execute(sql)
        If rs.BOF and rs.EOF Then
            '--
        Else
            SeriesRslts(UBound(SeriesMeets, 2) + 4, i) = rs(0).Value
        End If
		Set rs = Nothing
	Next
	
	'order by scores	
	ReDim TempArr(UBound(SeriesMeets, 2) + 4)
	For i = 0 To UBound(SeriesRslts, 2) - 2
		For m = i + 1 To UBound(SeriesRslts, 2) - 1
			If CInt(SeriesRslts(UBound(SeriesMeets, 2) + 4, i)) < CInt(SeriesRslts(UBound(SeriesMeets, 2) + 4, m)) Then
				For k = 0 To UBound(SeriesMeets, 2) + 4
					TempArr(k) = SeriesRslts(k, i)
					SeriesRslts(k, i) = SeriesRslts(k, m)
					SeriesRslts(k, m) = TempArr(k)
				Next
			End If
		Next
	Next

	iTeamPlace = 0
	For i = 0 To UBound(SeriesRslts, 2) - 1
		SeriesRslts(j + 2, i) = (CInt(iMaxScore) - CInt(iTeamPlace)*CInt(iDifference))*CInt(iScaleFactor)
		SeriesRslts(UBound(SeriesMeets, 2) + 3, i) = CInt(SeriesRslts(UBound(SeriesMeets, 2) + 3, i)) + CInt(SeriesRslts(j + 2, i))
		
		iTeamPlace = CInt(iTeamPlace) + 1
	Next
Next

'sort the results array
For i = 0 To UBound(SeriesRslts, 2) - 2
	For j = i + 1 To UBound(SeriesRslts, 2) - 1
		If CInt(SeriesRslts(UBound(SeriesMeets, 2) + 3, i)) < CInt(SeriesRslts(UBound(SeriesMeets, 2) + 3, j)) Then
			For k = 0 To UBound(SeriesMeets, 2) + 4
				TempArr(k) = SeriesRslts(k, i)
				SeriesRslts(k, i) = SeriesRslts(k, j)
				SeriesRslts(k, j) = TempArr(k)
			Next
		End If
	Next
Next
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE CCMeet/Nordic Ski Series</title>
<meta name="description" content="GSE cross-country running and nordic ski series results.">
<!--#include file = "../includes/js.asp" -->
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 style="margin-left:10px;">GSE Cross-Country Running/Nordic Skiing Series</h4>
			<h4 style="border:1px solid #ccc;background-color:#ececec;color:#036;"><%=sSeriesName%>&nbsp;(<%=sThisGender%>)</h4>
			
			<table style="font-size:0.85em;">
				<tr>
					<th valign="bottom">Pl</th>
					<th valign="bottom">Teams</th>
					<%For i = 0 To UBound(SeriesMeets, 2) - 1%>
						<th style="white-space:nowrap;">
							<%=SeriesMeets(1, i)%>
							<br>
							<span style="color:#555;"><%=SeriesMeets(2, i)%></span>
						</th>
					<%Next%>
					<th valign="bottom">Total</th>
				</tr>
				<%For i = 0 To UBound(SeriesRslts, 2) - 1%>
					<%If i mod 2 = 0 Then%>
						<tr>
							<td class="alt" style="text-align:right;color:#000;width:10px;"><%=i + 1%>)</td>
							<td class="alt" style="color:#000;width:100px;"><%=SeriesRslts(1, i)%></td>
							<%For j = 0 To UBound(SeriesMeets, 2) - 1%>
								<td class="alt" style="width:50px;color:#000;text-align:center;"><%=SeriesRslts(j + 2, i)%></td>
							<%Next%>
							<td class="alt" style="width:50px;text-align:center;"><%=SeriesRslts(j + 3, i)%></td>
						</tr>
					<%Else%>
						<tr>
							<td style="text-align:right;color:#000;width:10px;"><%=i + 1%>)</td>
							<td style="color:#000;width:100px;"><%=SeriesRslts(1, i)%></td>
							<%For j = 0 To UBound(SeriesMeets, 2) - 1%>
								<td style="width:50px;color:#000;text-align:center;"><%=SeriesRslts(j + 2, i)%></td>
							<%Next%>
							<td style="width:50px;text-align:center;"><%=SeriesRslts(j + 3, i)%></td>
						</tr>
					<%End If%>
				<%Next%>
			</table>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
