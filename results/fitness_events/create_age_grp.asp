<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql

Dim lEventID, lRaceID
Dim sRaceName, sGender, sMF, sChipStart, sOrderBy, sSortRsltsBy, sShowAge, sDist
Dim i, x, j, n, m, k
Dim sngMyTime
Dim Results(), TempArr(6), Races()
Dim iBegAge, iEnd Age, iRaceType
Dim sEventName, dEventDate

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
ElseIf Request.Form.Item("submit_ages") = "submit_ages" Then
    iBegAge = Request.Form.Item("beg_age")
    iEndAge =Request.Form.Item("end_age")
End If

If sGender = "M" Then
	sMF = "Male"
Else
	sMF = "Female"
End If

Dim sWhichTime
sWhichTime = Request.QueryString("which_time")
If sWhichTime = vbNullString Then sWhichTime = "chip"

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
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

Private Sub GetTheseRslts(iBegAge, iEndAge)
	i = 0
	ReDim Results(6, 0)
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
        If rs(2).Value <= iEndAge Then
            Results(0, i) = rs(7).Value & "-" & Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
            
            If CLng(lRaceID) = 350 Then
                Results(1, i) = "na"
            Else
                Results(1, i) = rs(2).Value
            End If

            If rs(3).Value & "" = "" Then
                Results(2, i) = "00:00"
                Results(3, i) = PacePerMile(ConvertToSeconds("00:00"), sDist)
                Results(4, i) = PacePerKM(ConvertToSeconds("00:00"), sDist)
            Else
                If sWhichTime = "chip" Then
                    Results(2, i) = ConvertToMinutes(CSng(rs(3).Value))
                    Results(2, i) = Replace(Results(2, i), "-", "")
                    Results(3, i) = PacePerMile(rs(3).Value, sDist)
                    Results(4, i) = PacePerKM(rs(3).Value, sDist)
                Else
                    Results(2, i) = rs(3).Value
                    Results(3, i) = PacePerMile(ConvertToSeconds(rs(3).Value), sDist)
                    Results(4, i) = PacePerKM(ConvertToSeconds(rs(3).Value), sDist)
                End If
            End If
            If rs(4).Value & "" = "" Then
                If rs(5).Value & "" = "" Then
                    Results(5, i) = "--"
                Else
                    Results(5, i) = rs(5).Value
                End If
            Else
                If rs(5).Value & "" = "" Then
                    Results(5, i) = Replace(rs(4).Value, "''", "'")
                Else
                    Results(5, i) = Replace(rs(4).Value, "''", "'") & ", " & rs(5).Value
                End If
            End If
            
            Results(6, i) = rs(6).Value
                                    
            i = i + 1
            ReDim Preserve Results(6, i)
        End If
		                     
	    If rs.RecordCount > 0 Then rs.MoveNext
	Loop

    For i = 0 To UBound(Results, 2) - 1
        If sShowAge = "n" Then
            Results(1, i) = ""
        Else
		    If Results(1, i) = "99" Then Results(1, i) = "0"
        End If
    Next
End Sub

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/pace_per_mile.asp" -->
<!--#include file = "../../includes/pace_per_km.asp" -->
<%
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
    <img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
    <div class="bg-info">
	    <a href="javascript:window.print()">Print This</a>
        &nbsp;|&nbsp;
	    <a href="dwnld_age_grps.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=<%=sWhichTime%>">Download</a>
        &nbsp;|&nbsp;
        <%If sWhichTime = "gun" Then%>
            <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=chip">View By Chip Time</a>
        <%Else%>
            <a href="age_grp_rslts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=gun">View By Gun Time</a>
        <%End If%>
    </div>

    <%If sWhichTime = "gun" Then%>
        <h1 class="h1">Age Group Results for <%=sEventName%></h1>
        <h2 class="h2"><%=sRaceName%>  (<%=sMF%>) <%=dEventDate%> - Gun Time</h2>
    <%Else%>
        <h1 class="h1">Age Group Results for <%=sEventName%></h1>
        <h2 class="h2"><%=sRaceName%>  (<%=sMF%>) <%=dEventDate%> - Chip Time</h2>
    <%End If%>

    <form role="form" class="form-inline" name="get_races" method="post" action="age_grp_rslts.asp?event_id=<%=lEventID%>">
    <div class="form_group">
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
	    <label for="gender">Gender:</label>
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
	    <input class="form-control" type="submit" name="get_race" id="get_race" value="View">
    </div>
    </form>

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
				    <td style="text-align:left;"><%=Results(0, i)%>
				    </td>
			        <%If sShowAge = "y" Then%>
					    <td><%=Results(1, i)%></td>
                    <%End If%>
				    <td><%=Results(2, i)%></td>
				    <td><%=Results(3, i)%></td>
				    <td><%=Results(4, i)%></td>
				    <td style="text-align:left;"><%=Results(5, i)%></td>
			    </tr>
			    <%If CInt(iRaceType) = 10 Then%>
				    <tr>
					    <td style="text-align:right;padding-top:0;" colspan="7">
						    <table style="font-size:0.95em;margin:0 0 0 50px;">
							    <%For x = 0 To 2%>
				                    <%sThisLeg = GetThisLeg(lRaceID, Results(6, i), x + 1)%>
				                    <%sThisSplit = GetThisSplit(lRaceID, Results(6, i), x + 1)%>
				                    <%If x = 2 Then%>
									    <%sThisTrans = vbNullString%>
								    <%Else%>
									    <%sThisTrans = GetThisTrans(lRaceID, Results(6, i), x + 1)%>
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
