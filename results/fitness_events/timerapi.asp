<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID
Dim Races(), IndRslts()
   
lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
	
Private Sub RaceResults(lRaceID, sGender)
     Dim x

	x = 0
	ReDim IndRslts(4, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT pr.Bib, p.FirstName, p.LastName, pr.Age, ir.ChipTime, ir.FnlTime, ir.ChipStart, p.City, p.St FROM Participant  p "
    sql = sql & "INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
	sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
    sql = sql & "INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
	sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
    rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		IndRslts(0, x) = rs(0).Value & "-" & Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
		IndRslts(1, x) = rs(3).Value
        If Not rs(4).Value & "" = "" Then 
            IndRslts(2, x) = Right(rs(4).Value, 9)
            IndRslts(2, x) = Left(rs(4).Value, 8)
        End If
        IndRslts(3, x) = rs(5).Value
        IndRslts(4, x) = rs(6).Value
		x = x + 1
		ReDim Preserve IndRslts(4, x)
		rs.MoveNext
	Loop
    rs.Close
	Set rs = Nothing
End Sub
%>
<html>
<body>

<%For j = 0 To UBound(Races, 2) - 1%>
    <div style="float: left;width: 640px;margin: 0;padding: 5px;font-size: 1.1em;">
        <%Call RaceResults(Races(0, j), "m")%>

        <h4 style="margin-top: 10px;">Results for <%=Races(1, j)%>-Male</h4>

		<table>
			<tr>
				<th style="width:10px">Pl</th>
				<th style="text-align:left;">Bib-Name</th>
  				<th>Age</th>
                <th>Chip Time</th>
				<th>Gun Time</th>
                <th>Chip Start</th>
			</tr>
			<%For i = 0 To UBound(IndRslts, 2) - 1%>
				<%If i mod 2 = 0 Then%>
					<tr>
						<td class="alt" style="width:10px;"><%=i + 1%></td>
						<td class="alt"><%=IndRslts(0, i)%></td>
						<td class="alt" style="text-align:center;"><%=IndRslts(1, i)%></td>
						<td class="alt"><%=IndRslts(2, i)%></td>
						<td class="alt"><%=IndRslts(3, i)%></td>
						<td class="alt"><%=IndRslts(4, i)%></td>
					</tr>
                <%Else%>
					<tr>
						<td style="width:10px;"><%=i + 1%></td>
						<td><%=IndRslts(0, i)%></td>
						<td style="text-align:center;"><%=IndRslts(1, i)%></td>
						<td><%=IndRslts(2, i)%></td>
						<td><%=IndRslts(3, i)%></td>
						<td><%=IndRslts(4, i)%></td>
					</tr>
                <%End If%>
			<%Next%>
		</table>
    </div>
    <div style="margin: 0 0 0 650px;padding: 1px 0 0 0;font-size: 1.1em;">
        <%Call RaceResults(Races(0, j), "f")%>
 
        <h4 style="margin-top: 10px;">Results for <%=Races(1, j)%>-Female</h4>

		<table>
			<tr>
				<th style="width:10px">Pl</th>
				<th style="text-align:left;">Bib-Name</th>
  				<th>Age</th>
				<th>Chip Time</th>
                <th>Gun Time</th>
                <th>Chip Start</th>
			</tr>
			<%For i = 0 To UBound(IndRslts, 2) - 1%>
				<%If i mod 2 = 0 Then%>
					<tr>
						<td class="alt" style="width:10px;"><%=i + 1%></td>
						<td class="alt"><%=IndRslts(0, i)%></td>
						<td class="alt" style="text-align:center;"><%=IndRslts(1, i)%></td>
						<td class="alt"><%=IndRslts(2, i)%></td>
						<td class="alt"><%=IndRslts(3, i)%></td>
						<td class="alt"><%=IndRslts(4, i)%></td>
					</tr>
                <%Else%>
					<tr>
						<td style="width:10px;"><%=i + 1%></td>
						<td><%=IndRslts(0, i)%></td>
						<td style="text-align:center;"><%=IndRslts(1, i)%></td>
						<td><%=IndRslts(2, i)%></td>
						<td><%=IndRslts(3, i)%></td>
						<td><%=IndRslts(4, i)%></td>
					</tr>
                <%End If%>
			<%Next%>
		</table>
    </div>
<%Next%>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>