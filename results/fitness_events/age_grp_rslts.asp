<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim iRaceType
Dim lRaceID
Dim sRaceName, sGender, sMF, sThisLeg, sThisSplit, sThisTrans, sTimingMethod, sChipStart, sOrderBy, sSortRsltsBy, sShowAge, sDist
Dim i, x, j, n, m, k
Dim sngMyTime
Dim AgeGrps(), iBegAge, IndRslts(), TempArr(6), Races()
Dim lEventID, sEventName, dEventDate

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
sGender = Request.QueryString("gender")
If sGender = vbNullString Then sGender = "M"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get races
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
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

If Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
	sGender = Request.Form.Item("gender")
End If

If sGender = "M" Then
	sMF = "Male"
Else
	sMF = "Female"
End If

Dim sWhichTime
sWhichTime = Request.QueryString("which_time")
If sWhichTime = vbNullString Then sWhichTime = "chip"

sql = "SELECT EventName, EventDate, TimingMethod FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sTimingMethod = rs(2).Value
Set rs = Nothing

sql = "SELECT Dist, Type, ChipStart, SortRsltsBy, ShowAge, RaceName FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sDist = rs(0).Value
iRaceType = rs(1).Value
sChipStart = rs(2).Value
sSortRsltsBy = rs(3).Value
sShowAge = rs(4).Value
sRaceName = rs(5).Value
Set rs = Nothing
	
If sSortRsltsBy = "FnlTime" Then
    sOrderBy = "ir.FnlScnds"
Else
    sOrderBy = "ir.EventPl"
End If

i = 0
iBegAge = 0
ReDim AgeGrps(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EndAge, AgeGrpName FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = '" & LCase(sGender) & "' ORDER BY EndAge"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    AgeGrps(0, i) = iBegAge
    AgeGrps(1, i) = rs(0).Value
    AgeGrps(2, i) = rs(1).Value

    iBegAge = CInt(rs(0).Value) + 1
    i = i + 1
    ReDim Preserve AgeGrps(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetTheseRslts(iBegAge, iEndAge)
	i = 0
	ReDim IndRslts(6, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
    If sWhichTime = "chip" Then
	    sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlScnds, p.City, p.St, p.ParticipantID, pr.Bib, ir.ChipStart "
        sql = sql & "FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
	    sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
        sql = sql & "INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
	    sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND pr.Age >= " & iBegAge & " AND pr.Age <= " & iEndAge 
	    sql = sql & " AND pr.Age <> 99 AND ir.Eligible = 'y' AND ir.FnlScnds > 0 ORDER BY " & sOrderBy
    Else
	    sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlTime, p.City, p.St, p.ParticipantID, pr.Bib, ir.ChipStart "
        sql = sql & "FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
	    sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID "
        sql = sql & "INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
	    sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND pr.Age >= " & iBegAge & " AND pr.Age <= " & iEndAge 
	    sql = sql & " AND pr.Age <> 99 AND ir.Eligible = 'y' AND ir.FnlScnds > 0 ORDER BY " & sOrderBy
    End If
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
	    If rs(2).Value >= iBegAge Then
	        If rs(2).Value <= iEndAge Then
	            IndRslts(0, i) = rs(7).Value & "-" & Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
				
				If CLng(lRaceID) = 350 Then
					IndRslts(1, i) = "na"
				Else
					IndRslts(1, i) = rs(2).Value
				End If

	            If rs(3).Value & "" = "" Then
	                IndRslts(2, i) = "00:00"
	                IndRslts(3, i) = PacePerMile(ConvertToSeconds("00:00"), sDist)
	                IndRslts(4, i) = PacePerKM(ConvertToSeconds("00:00"), sDist)
	            Else
				    If sWhichTime = "chip" Then
                        IndRslts(2, i) = ConvertToMinutes(CSng(rs(3).Value))
                        IndRslts(2, i) = Replace(IndRslts(2, i), "-", "")
				        IndRslts(3, i) = PacePerMile(rs(3).Value, sDist)
				        IndRslts(4, i) = PacePerKM(rs(3).Value, sDist)
                    Else
	                    IndRslts(2, i) = rs(3).Value
	                    IndRslts(3, i) = PacePerMile(ConvertToSeconds(rs(3).Value), sDist)
	                    IndRslts(4, i) = PacePerKM(ConvertToSeconds(rs(3).Value), sDist)
                    End If
	            End If
	            If rs(4).Value & "" = "" Then
	                If rs(5).Value & "" = "" Then
	                    IndRslts(5, i) = "--"
	                Else
	                    IndRslts(5, i) = rs(5).Value
	                End If
	            Else
	                If rs(5).Value & "" = "" Then
	                    IndRslts(5, i) = Replace(rs(4).Value, "''", "'")
	                Else
	                    IndRslts(5, i) = Replace(rs(4).Value, "''", "'") & ", " & rs(5).Value
	               End If
	            End If
				
	            IndRslts(6, i) = rs(6).Value
		                             
	            i = i + 1
	            ReDim Preserve IndRslts(6, i)
	        End If
	    End If
		                     
	    If rs.RecordCount > 0 Then rs.MoveNext
	Loop

    For i = 0 To UBound(IndRslts, 2) - 1
        If sShowAge = "n" Then
            IndRslts(1, i) = ""
        Else
		    If IndRslts(1, i) = "99" Then IndRslts(1, i) = "0"
        End If
    Next
End Sub

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
    GetThisLeg = Replace(rs(0).Value, "''", "'")
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
<title>GSE Age Group Results</title>
<meta name="description" content="GSE age group results.">
<!--#include file = "../../includes/js.asp" --> 
</head>
<body>
<div class="container">
     <div class="row">
        <div class="col-sm-6">
            <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
        </div>
        <div class="col-sm-6">
            <h1 class="h1">GSE Age Group Results</h1>
        </div>
    </div>
 
    <div class="bg-danger">
	    <a href="javascript:window.print()" style="color:#fff;">Print This</a>
        &nbsp;|&nbsp;
	    <a style="color:#fff;" href="dwnld_age_grps.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=<%=sWhichTime%>">Download</a>
        &nbsp;|&nbsp;
        <%If sWhichTime = "gun" Then%>
            <a style="color:#fff;" href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=chip">View By Chip Time</a>
        <%Else%>
            <a style="color:#fff;" href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=gun">View By Gun Time</a>
        <%End If%>
    </div>

    <%If sWhichTime = "gun" Then%>
        <h1 class="h1">A<%=sEventName%></h1>
        <h2 class="h2"><%=sRaceName%>  (<%=sMF%>) <%=dEventDate%> - Gun Time</h2>
    <%Else%>
        <h1 class="h1"><%=sEventName%></h1>
        <h2 class="h2"><%=sRaceName%>  (<%=sMF%>) <%=dEventDate%> - Chip Time</h2>
    <%End If%>

    <form role="form" class="form-inline" name="get_races" method="post" action="age_grp_rslts.asp?event_id=<%=lEventID%>">
	<label for="races">Race:</label>
	<select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
		<%For i = 0 to UBound(Races, 2) - 1%>
			<%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
				<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
			<%Else%>
				<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
			<%End If%>
		<%Next%>
	</select>
	&nbsp;&nbsp;<label for="gender">Gender:</label>
	<select class="form-control" name="gender" id="gender" onchange="this.form.get_race.click()">
		<%Select Case sGender%>
			<%Case "M"%>
				<option value="M" selected>Male</option>
				<option value="F">Female</option>
			<%Case "F"%>
				<option value="M">Male</option>
				<option value="F" selected>Female</option>
		<%End Select%>
	</select>
	<input class="form-control" type="hidden" name="submit_race" id="submit_race" value="submit_race">
	<input class="form-control" type="submit" name="get_race" id="get_race" value="View This">
    </form>

	<br>

    <table class="table table-striped">
	    <%For j = 0 to UBound(AgeGrps, 2) - 1%>
		    <%Call GetTheseRslts(AgeGrps(0, j), AgeGrps(1, j))%>
		    <tr>
			    <th style="text-align:left;padding-top:10px;" colspan="9"><%=AgeGrps(2, j)%></th>
		    </tr>
		    <tr>
			    <th>Pl</th>
			    <th style="text-align:left;">Bib-Name</th>
			    <%If sShowAge = "y" Then%>
                    <th>Age</th>
                <%End If%>
			    <th>Time</th>
			    <th>Per Mi</th>
			    <th>Per Km</th>
			    <th>From</th>
		    </tr>
		    <%For i = 0 To UBound(IndRslts, 2) - 1%>
			    <tr>
				    <td style="width:10px;"><%=i + 1%></td>
				    <td style="text-align:left;"><%=IndRslts(0, i)%>
				    </td>
			        <%If sShowAge = "y" Then%>
					    <td><%=IndRslts(1, i)%></td>
                    <%End If%>
				    <td><%=IndRslts(2, i)%></td>
				    <td><%=IndRslts(3, i)%></td>
				    <td><%=IndRslts(4, i)%></td>
				    <td style="text-align:left;"><%=IndRslts(5, i)%></td>
			    </tr>
			    <%If CInt(iRaceType) = 10 Then%>
				    <tr>
					    <td style="text-align:right;padding-top:0;" colspan="7">
						    <table style="font-size:0.95em;margin:0 0 0 50px;">
							    <%For x = 0 To 2%>
				                    <%sThisLeg = GetThisLeg(lRaceID, IndRslts(6, i), x + 1)%>
				                    <%sThisSplit = GetThisSplit(lRaceID, IndRslts(6, i), x + 1)%>
				                    <%If x = 2 Then%>
									    <%sThisTrans = vbNullString%>
								    <%Else%>
									    <%sThisTrans = GetThisTrans(lRaceID, IndRslts(6, i), x + 1)%>
								    <%End If%>
								    <tr>
									    <th>
										    <%=sThisLeg%>
									    </th>
									    <td>
										    <%=sThisSplit%>
									    </td>
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
	    <%Next%>
    </table>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
