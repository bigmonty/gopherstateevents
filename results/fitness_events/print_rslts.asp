<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, x, m, n
Dim lRaceID, lRsltsEventID
Dim iRaceType
Dim sEventName, sGender, sSortRsltsBy, sDist, sRaceName, sMF, sThisLeg, sThisSplit, sThisTrans, sAllowDuplAwds, sGallery, sTimingMethod, sChipStart
Dim sWeather, sShowAge, sEventRaces
Dim sngMyTime
Dim dEventDate
Dim Events(), IndRslts, TempArr(9)

lRsltsEventID = Request.QueryString("rslts_event_id")
If CStr(lRsltsEventID) = vbNullString Then lRsltsEventID = 0
If Not IsNumeric(lRsltsEventID) Then Response.Redirect("http://www.google.com")
If CLng(lRsltsEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
sGender = Request.QueryString("gender")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate, TimingMethod, Weather FROM Events WHERE EventID = " & lRsltsEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sTimingMethod = rs(2).Value
If Not rs(3).Value & "" = "" Then sWeather = Replace(rs(3).Value, "''", "'")
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lRsltsEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sEventRaces = sEventRaces & rs(0).Value & ", "
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

sql = "SELECT Dist, RaceName, Type, AllowDuplAwds, ChipStart, SortRsltsBy, ShowAge FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sDist = rs(0).Value
sRaceName = rs(1).Value
iRaceType = rs(2).Value
sAllowDuplAwds = rs(3).Value
sChipStart = rs(4).Value
sSortRsltsBy = rs(5).Value
sShowAge = rs(6).Value
Set rs = Nothing

If sGender = "B" Then
    If sSortRsltsBy = "place" Then
        Set rs = Server.CreateObject("ADODB.Recordset")  
        sql = "Exec ResultsByPlace @RaceID = " & lRaceID
        Set rs = conn.execute(sql) 
        If rs.BOF and rs.EOF Then
            ReDim IndRslts(10, 0)
        Else
            IndRslts = rs.GetRows()
        End If
        rs.Close
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")  
        sql = "Exec OverallResults @RaceID = " & lRaceID
        Set rs = conn.execute(sql) 
        If rs.BOF and rs.EOF Then
            ReDim IndRslts(10, 0)
        Else
            IndRslts = rs.GetRows()
        End If
        rs.Close
        Set rs = Nothing
    End If
Else
    If sSortRsltsBy = "place" Then
        Set rs = Server.CreateObject("ADODB.Recordset")  
        sql = "Exec GenderByPlace @RaceID = " & lRaceID & ", @Gender = '" & sGender & "'"
        Set rs = conn.execute(sql) 
        If rs.BOF and rs.EOF Then
            ReDim IndRslts(10, 0)
        Else
            IndRslts = rs.GetRows()
        End If
        rs.Close
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")  
        sql = "Exec GenderResults @RaceID = " & lRaceID & ", @Gender = '" & sGender & "'"
        Set rs = conn.execute(sql) 
        If rs.BOF and rs.EOF Then
            ReDim IndRslts(10, 0)
        Else
            IndRslts = rs.GetRows()
        End If
        rs.Close
        Set rs = Nothing
    End If
End If

For i = 0 To UBound(IndRslts, 2)
    If sShowAge = "n" Then
        IndRslts(4, i) = MyAgeGrp(IndRslts(0, i))
    Else
		If IndRslts(4, i) = "99" Then IndRslts(4, i) = "0"
    End If
Next

Private Function MyAgeGrp(iMyBib)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AgeGrp FROM PartRace WHERE Bib = " & iMyBib & " AND RaceID IN (" & sEventRaces & ")"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        If Left(rs(0).Value, 3) = "110" Then
            MyAgeGrp = "n/a"
        Else
            MyAgeGrp = rs(0).Value
        End If
    Else
        MyAgeGrp = "n/a"
    End If
    rs.Close
End Function

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/pace_per_mile.asp" -->
<!--#include file = "../../includes/pace_per_km.asp" -->
<%

Private Function GetThisLeg(lThisRace, lThisPart, iMmbrNum)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MmbrName FROM TeamMmbrs WHERE RaceID = " & lThisRace & " AND ParticipantID = " & lThisPart & " AND MmbrNum = "
    sql = sql & iMmbrNum & " ORDER BY MmbrNum"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetThisLeg = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

Private Function GetThisSplit(lThisRace, lThisPart, iSplitNum)
    Set rs = Server.CreateObject("ADODB.Recordset")
    Select Case iSplitNum
        Case 1
            sql = "SELECT rd.RaceDelay, pr.IndDelay, pr.Trans1Out FROM IndResults ir INNER JOIN RaceData rd ON rd.RaceID = ir.RaceID "
            sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID WHERE pr.RaceID = " & lThisRace
            sql = sql & " AND pr.ParticipantID = " & lThisPart
        Case 2
            sql = "SELECT pr.Trans1Out, pr.Trans2Out FROM IndResults ir INNER JOIN RaceData rd ON rd.RaceID = ir.RaceID "
            sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID WHERE pr.RaceID = " & lThisRace
            sql = sql & " AND pr.ParticipantID = " & lThisPart
        Case 3
            sql = "SELECT pr.Trans2Out, ir.ElpsdTime FROM IndResults ir INNER JOIN RaceData rd ON rd.RaceID = ir.RaceID "
            sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID WHERE pr.RaceID = " & lThisRace
            sql = sql & " AND pr.ParticipantID = " & lThisPart
    End Select
    rs.Open sql, conn, 1, 2
    Select Case iSplitNum
        Case 1
            If ConvertToSeconds(rs(2).Value) = 0 Then
                GetThisSplit = "unavail"
            Else
                GetThisSplit = ConvertToMinutes(ConvertToSeconds(rs(2).Value) - ConvertToSeconds(rs(1).Value) - ConvertToSeconds(rs(0).Value))
            End If
        Case Else
            If ConvertToSeconds(rs(0).Value) = 0 Or ConvertToSeconds(rs(1).Value) = 0 Then
                GetThisSplit = "unavail"
            Else
                GetThisSplit = ConvertToMinutes(ConvertToSeconds(rs(1).Value) - ConvertToSeconds(rs(0).Value))
            End If
    End Select
    rs.Close
    Set rs = Nothing
End Function

Private Function GetThisTrans(lThisRace, lThisPart, iTransNum)
    Set rs = Server.CreateObject("ADODB.Recordset")
    Select Case iTransNum
        Case 1
            sql = "SELECT pr.Trans1In, pr.Trans1Out FROM IndResults ir INNER JOIN RaceData rd "
            sql = sql & "ON rd.RaceID = ir.RaceID INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID "
            sql = sql & "WHERE pr.RaceID = " & lThisRace & " AND pr.ParticipantID = " & lThisPart
        Case 2
            sql = "SELECT pr.Trans2In, pr.Trans2Out FROM IndResults ir INNER JOIN RaceData rd "
            sql = sql & "ON rd.RaceID = ir.RaceID INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID "
            sql = sql & "WHERE pr.RaceID = " & lThisRace & " AND pr.ParticipantID = " & lThisPart
    End Select
    rs.Open sql, conn, 1, 2
    If ConvertToSeconds(rs(0).Value) = 0 Or ConvertToSeconds(rs(1).Value) = 0 Then
        GetThisTrans = "unavail"
    Else
        GetThisTrans = ConvertToMinutes(ConvertToSeconds(rs(1).Value) - ConvertToSeconds(rs(0).Value))
    End If
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Print GSE Results</title>
<!--#include file = "../../includes/js.asp" -->
<meta name="description" content="Print Gopher State Events (GSE) results.">
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
    <a href="javascript:window.print();" style="font-size: 0.8em;">Print</a>

	<h1 class="h1">Gopher State Events Results</h1>
    <h2 class="h2"><%=sEventName%> On <%=dEventDate%></h2>

    <%If Not sWeather = vbNullString Then%>
        <p><span style="font-weight:bold;">Weather:</span>&nbsp;<%=sWeather%></p>
    <%End If%>

	<%If Not CLng(lRaceID) = 0 Then%>
        <div class="bg-success">Note:  This race used a chip start which takes into consideration when you crossed the starting line as well as 
        when you crossed the finish line.  As a result, people that were close to you at the finish line may not appear that way in the results.</div>
            
		<table class="table table-striped">
			<tr>
                <%If sTimingMethod = "RFID" And sChipStart = "y" Then%>
				    <th>Pl</th>
				    <th>Bib-Name</th>
				    <th>M/F</th>
  				    <%If sShowAge = "y" Then%>
                        <th>Age</th>
                    <%Else%>
                        <th>Age Grp</th>
                    <%End If%>
				    <th>Chip Time</th>
				    <th>Gun Time</th>
				    <th>Start Time</th>
				    <th>From</th>
                <%Else%>
				    <th>Pl</th>
				    <th>Bib-Name</th>
				    <th>M/F</th>
  				    <th>Age</th>
				    <th>Time</th>
				    <th>Per Mi</th>
				    <th>Per Km</th>
				    <th>From</th>
                    <th>&nbsp;</th>
                <%End If%>
			</tr>

			<%For i = 0 To UBound(IndRslts, 2)%>
				<tr>
					<td><%=i + 1%></td>
					<td><%=IndRslts(0, i)%> - <%=IndRslts(2, i)%>&nbsp;<%=IndRslts(1, i)%></td>
					<td><%=IndRslts(3, i)%></td>
					<%If CLng(lRaceID) = 350 Then%>
			            <td>n/a</td>
		            <%Else%>
			            <td><%=IndRslts(4, i)%></td>
		            <%End If%>
					<td><%=IndRslts(5, i)%></td>
					<td><%=IndRslts(6, i)%></td>
					<td><%=IndRslts(7, i)%></td>
					<td><%=IndRslts(8, i)%>, <%=IndRslts(9, i)%></td>
				</tr>
				<%If CInt(iRaceType) = 10 Then%>
					<tr>
						<td colspan="7">
							<table class="table">
								<%For x = 0 To 2%>
							        <%sThisLeg = GetThisLeg(lRaceID, IndRslts(6, i), x + 1)%>
							        <%sThisSplit = GetThisSplit(lRaceID, IndRslts(6, i), x + 1)%>
							        <%If x = 2 Then%>
										<%sThisTrans = vbNullString%>
									<%Else%>
										<%sThisTrans = GetThisTrans(lRaceID, IndRslts(6, i), x + 1)%>
									<%End If%>
									<tr>
										<th><%=sThisLeg%></th>
										<td><%=sThisSplit%></td>
										<td>
											<%If sThisTrans = vbNullString Then%>
												&nbsp;
											<%Else%>
												(Trans:&nbsp;<%=sThisTrans%>)
											<%End If%>
										</td>
									</tr>
								<%Next%>
							</table>
						</td>
					</tr>
				<%End If%>
			<%Next%>
		</table>
	<%End If%>
	<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>