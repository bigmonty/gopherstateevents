<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet, lThisRace
Dim Races(), TmRslts()
Dim sMeetName, sRaceName, sScoreMethod, sRaceGender
Dim dMeetDate

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")
lThisRace = Request.QueryString("this_race")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisMeet
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
    Races(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not CLng(lThisRace) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceName, ScoreMethod, Gender FROM Races WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    sRaceName = Replace(rs(0).Value, "''", "'")
    sScoreMethod = rs(1).Value
    sRaceGender = rs(2).Value
    rs.Close
    Set rs = Nothing

    i = 0
    ReDim TmRslts(8, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
    sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lThisRace & " AND tr.Score <> ''"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    TmRslts(0, i) = rs(0).Value
	    TmRslts(1, i) = rs(1).Value
	    TmRslts(2, i) = Trim(rs(2).Value)
	    TmRslts(3, i) = Trim(rs(3).Value)
	    TmRslts(4, i) = Trim(rs(4).Value)
	    TmRslts(5, i) = Trim(rs(5).Value)
	    TmRslts(6, i) = Trim(rs(6).Value)
	    TmRslts(7, i) = Trim(rs(7).Value)
	    TmRslts(8, i) = Trim(rs(8).Value)
	    i = i + 1
	    ReDim Preserve TmRslts(8, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Sub GetTmRslts(lRaceID)
    Dim x

    x = 0
    ReDim TmRslts(8, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
    sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> ''"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    TmRslts(0, x) = rs(0).Value
	    TmRslts(1, x) = rs(1).Value
	    TmRslts(2, x) = Trim(rs(2).Value)
	    TmRslts(3, x) = Trim(rs(3).Value)
	    TmRslts(4, x) = Trim(rs(4).Value)
	    TmRslts(5, x) = Trim(rs(5).Value)
	    TmRslts(6, x) = Trim(rs(6).Value)
	    TmRslts(7, x) = Trim(rs(7).Value)
	    TmRslts(8, x) = Trim(rs(8).Value)
	    x = x + 1
	    ReDim Preserve TmRslts(8, x)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>Print GSE CC/Nordic Results Manager: Team Scores</title>
<!--#include file = "../../../includes/meta2.asp" -->



</head>
<body>
<div style="margin: 10px;padding: 10px;background-color: #fff;">
    <h4 style="margin-top: 10px;">Team Scores for <%=sMeetName%> on <%=dMeetDate%></h4>
            	
    <div style="margin: 0;padding: 0;font-size: 0.8em;text-align: left;">
        <a href="javascript:window.print();">Print These</a>
    </div>	
    <%If CLng(lThisRace) = 0 Then%>	
        <%For j = 0 To UBound(Races, 2) - 1%>
            <%Call GetTmRslts(Races(0, j))%>
 			<h4 style="margin-top:10px;text-align: left;"><%=Races(1, j)%></h4>
						
			<table>
				<tr>
					<th style="text-align:right">Pl</th>
					<th>Team</th>
					<th style="text-align:center">Score</th>
					<th style="text-align:center;color:#717889">R1</th>
					<th style="text-align:center;color:#717889">R2</th>
					<th style="text-align:center;color:#717889">R3</th>
					<th style="text-align:center;color:#717889">R4</th>
					<th style="text-align:center;color:#717889">R5</th>
					<th style="text-align:center;color:#717889">R6</th>
					<th style="text-align:center;color:#717889">R7</th>
				</tr>
				<%For i = 0 to UBound(TmRslts, 2) - 1%>
					<%If i mod 2 = 0 Then%>
						<tr>
							<td class="alt" style="width:15px;text-align:right"><%=i + 1%></td>
							<td class="alt" style="width:150px;text-align:left;white-space:nowrap;"><%=TmRslts(0, i)%></td>
							<td class="alt" style="text-align:center"><%=TmRslts(1, i)%></td>
							<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(2, i)%></td>
							<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(3, i)%></td>
							<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(4, i)%></td>
							<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(5, i)%></td>
							<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(6, i)%></td>
							<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(7, i)%></td>
							<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(8, i)%></td>
						</tr>
					<%Else%>
						<tr>
							<td style="width:15px;text-align:right"><%=i + 1%></td>
							<td style="width:150px;text-align:left;white-space:nowrap;"><%=TmRslts(0, i)%></td>
							<td style="text-align:center"><%=TmRslts(1, i)%></td>
							<td style="text-align:center;color:#717889"><%=TmRslts(2, i)%></td>
							<td style="text-align:center;color:#717889"><%=TmRslts(3, i)%></td>
							<td style="text-align:center;color:#717889"><%=TmRslts(4, i)%></td>
							<td style="text-align:center;color:#717889"><%=TmRslts(5, i)%></td>
							<td style="text-align:center;color:#717889"><%=TmRslts(6, i)%></td>
							<td style="text-align:center;color:#717889"><%=TmRslts(7, i)%></td>
							<td style="text-align:center;color:#717889"><%=TmRslts(8, i)%></td>
						</tr>
					<%End If%>
				<%Next%>
			</table>
        <%Next%>
    <%Else%>	
		<h4 style="margin-top:10px;text-align: left;"><%=sRaceName%></h4>
						
		<table>
			<tr>
				<th style="text-align:right">Pl</th>
				<th>Team</th>
				<th style="text-align:center">Score</th>
				<th style="text-align:center;color:#717889">R1</th>
				<th style="text-align:center;color:#717889">R2</th>
				<th style="text-align:center;color:#717889">R3</th>
				<th style="text-align:center;color:#717889">R4</th>
				<th style="text-align:center;color:#717889">R5</th>
				<th style="text-align:center;color:#717889">R6</th>
				<th style="text-align:center;color:#717889">R7</th>
			</tr>
			<%For i = 0 to UBound(TmRslts, 2) - 1%>
				<%If i mod 2 = 0 Then%>
					<tr>
						<td class="alt" style="width:15px;text-align:right"><%=i + 1%></td>
						<td class="alt" style="width:150px;text-align:left;white-space:nowrap;"><%=TmRslts(0, i)%></td>
						<td class="alt" style="text-align:center"><%=TmRslts(1, i)%></td>
						<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(2, i)%></td>
						<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(3, i)%></td>
						<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(4, i)%></td>
						<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(5, i)%></td>
						<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(6, i)%></td>
						<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(7, i)%></td>
						<td class="alt" style="text-align:center;color:#717889"><%=TmRslts(8, i)%></td>
					</tr>
				<%Else%>
					<tr>
						<td style="width:15px;text-align:right"><%=i + 1%></td>
						<td style="width:150px;text-align:left;white-space:nowrap;"><%=TmRslts(0, i)%></td>
						<td style="text-align:center"><%=TmRslts(1, i)%></td>
						<td style="text-align:center;color:#717889"><%=TmRslts(2, i)%></td>
						<td style="text-align:center;color:#717889"><%=TmRslts(3, i)%></td>
						<td style="text-align:center;color:#717889"><%=TmRslts(4, i)%></td>
						<td style="text-align:center;color:#717889"><%=TmRslts(5, i)%></td>
						<td style="text-align:center;color:#717889"><%=TmRslts(6, i)%></td>
						<td style="text-align:center;color:#717889"><%=TmRslts(7, i)%></td>
						<td style="text-align:center;color:#717889"><%=TmRslts(8, i)%></td>
					</tr>
				<%End If%>
			<%Next%>
		</table>
    <%End If%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
