<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, x, j, n, m, k
Dim lRaceID, lEventID
Dim sRaceName, sGender, sMF, sThisLeg, sThisSplit, sThisTrans, sOrderBy, sSortRsltsBy, sEventName, sShowAge
Dim iMAwds, iFAwds, iBegAge, iRaceType, iNumAwds, iDuplAwds
Dim AgeGrps(), IndRslts(), TempArr(6), OpenRslts(), OpenAwds(), Races()
Dim bAddThis
Dim dEventDate

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
sGender = Request.QueryString("gender")
If sGender = vbNullString Then sGender = "M"

Dim sWhichTime
sWhichTime = Request.QueryString("which_time")
If sWhichTime = vbNullString Then sWhichTime = "chip"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = rs(0).Value
dEventDate = rs(1).Value
Set rs = Nothing

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

sql = "SELECT Dist, Type, MAwds, FAwds, SortRsltsBy, AllowDuplAwds, ShowAge FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
iRaceType = rs(1).Value
iMAwds = rs(2).Value
iFAwds = rs(3).Value
sSortRsltsBy = rs(4).Value
iDuplAwds = rs(5).Value
sShowAge = rs(6).Value
Set rs = Nothing
		
If sSortRsltsBy = "FnlTime" Then
    sOrderBy = "ir.FnlScnds"
Else
    sOrderBy = "ir.EventPl"
End If

'get age groups for this race
i = 0
iBegAge = 0
ReDim AgeGrps(2, 0)
sql = "SELECT EndAge, NumAwds, AgeGrpName FROM AgeGroups WHERE Gender = '" & sGender & "' AND RaceID = " & lRaceID
sql = sql & " ORDER BY EndAge"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    If CInt(rs(1).Value) > 0 Then
        AgeGrps(0, i) = iBegAge
        AgeGrps(1, i) = rs(0).Value
        AgeGrps(2, i) = rs(2).Value
        iBegAge = rs(0).Value + 1
        i = i + 1
        ReDim Preserve AgeGrps(2, i)
    End If
    rs.MoveNext
Loop
Set rs = Nothing

'get open results
i = 0
ReDim OpenRslts(6, 0)
If sWhichTime = "gun" Then
    sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlTime, p.City, p.St, p.ParticipantID, pr.Bib FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
    sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
    sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND ir.FnlScnds > 0 AND ir.Eligible = 'y' ORDER BY " & sOrderBy
Else
    sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlScnds, p.City, p.St, p.ParticipantID, pr.Bib FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
    sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
    sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND ir.FnlScnds > 0 AND ir.Eligible = 'y' ORDER BY " & sOrderBy
End If
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    OpenRslts(0, i) = rs(7).Value & "-" & Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
				
	If CLng(lRaceID) = 350 Then
		OpenRslts(1, i) = "na"
	Else
		OpenRslts(1, i) = rs(2).Value
	End If
				
    If rs(3).Value & "" <> "" Then
		If sWhichTime = "chip" Then
            OpenRslts(2, i) = ConvertToMinutes(CSng(rs(3).Value))
			OpenRslts(3, i) = PacePerMile(rs(3).Value, sRaceName)
			OpenRslts(4, i) = PacePerKM(rs(3).Value, sRaceName)
        Else
	        OpenRslts(2, i) = rs(3).Value
	        OpenRslts(3, i) = PacePerMile(ConvertToSeconds(rs(3).Value), sRaceName)
	        OpenRslts(4, i) = PacePerKM(ConvertToSeconds(rs(3).Value), sRaceName)
        End If
	End If
	
    If rs(4).Value & "" = "" Then
        If rs(5).Value & "" = "" Then
            OpenRslts(5, i) = "--"
        Else
            OpenRslts(5, i) = rs(5).Value
        End If
    Else
        If rs(5).Value & "" = "" Then
            OpenRslts(5, i) = Replace(rs(4).Value, "''", "'")
        Else
            OpenRslts(5, i) = Replace(rs(4).Value, "''", "'") & ", " & rs(5).Value
       End If
    End If
	
    OpenRslts(6, i) = rs(6).Value
                
    i = i + 1
    ReDim Preserve OpenRslts(6, i)
	
    rs.MoveNext
Loop
Set rs = Nothing

'now just take the top few
If sGender = "M" Then
	ReDim OpenAwds(6, CInt(iMAwds))
Else
	ReDim OpenAwds(6, CInt(iFAwds))
End If

If UBound(OpenRslts, 2) > 0 Then
    For i = 0 To UBound(OpenAwds, 2)
	    For j = 0 To 6
		    OpenAwds(j, i) = OpenRslts(j, i)
	    Next
    Next
End If

'get remaining results
Private Sub GetTheseRslts(iBegAge, iEndAge)
    Dim x

    'get num awds for this age group
    iNumAwds = 0
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT NumAwds, EndAge FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = '" & sGender & "' ORDER BY EndAge"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If CInt(rs(1).Value) = "110" Then
            iNumAwds = Cint(rs(0).Value)
        Else
            If CInt(rs(1).Value) = CInt(iEndAge) Then 
                iNumAwds = CInt(rs(0).Value)
                Exit Do
            End If
        End If
                
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

	ReDim IndRslts(6, 0)
    If CInt(iNumAwds) > 0 Then
	    i = 0
	    Set rs = Server.CreateObject("ADODB.Recordset")
        If sWhichTime = "gun" Then
	        sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlTime, p.City, p.St, p.ParticipantID, pr.Bib FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
	        sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
	        sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND pr.Age >= " & iBegAge & " AND pr.Age <= " & iEndAge 
	        sql = sql & " AND ir.Eligible = 'y' AND pr.Age <> 99 AND ir.FnlScnds > 0 ORDER BY " & sOrderBy
        Else
	        sql = "SELECT p.FirstName, p.LastName, pr.Age, ir.FnlScnds, p.City, p.St, p.ParticipantID, pr.Bib FROM Participant  p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
	        sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID INNER JOIN RaceData rd ON rd.RaceID = pr.RaceID AND rd.RaceID = ir.RaceID "
	        sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND p.Gender = '" & sGender & "' AND pr.Age >= " & iBegAge & " AND pr.Age <= " & iEndAge 
	        sql = sql & " AND ir.Eligible = 'y' AND pr.Age <> 99 AND ir.FnlScnds > 0 ORDER BY " & sOrderBy
        End If
	    rs.Open sql, conn, 1, 2
	    Do While Not rs.EOF
            bAddThis = True

	        If CInt(rs(2).Value) >= CInt(iBegAge) Then
	            If CInt(rs(2).Value) <= CInt(iEndAge) Then
                    If iDuplAwds = "n" Then
                        For x = 0 To UBound(OpenAwds, 2) - 1
                            If CLng(rs(6).Value) = Clng(OpenAwds(6, x)) Then
                                bAddThis = False
                                Exit For
                            End If
                        Next
                    End If

                    If bAddThis = True Then
	                    IndRslts(0, i) = rs(7).Value & "-" & Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
	                    IndRslts(1, i) = rs(2).Value
	                    If rs(3).Value & "" = "" Then
	                        IndRslts(2, i) = "00:00"
	                        IndRslts(3, i) = PacePerMile(ConvertToSeconds("00:00"), sRaceName)
	                        IndRslts(4, i) = PacePerKM(ConvertToSeconds("00:00"), sRaceName)
	                    Else
				            If sWhichTime = "chip" Then
                                IndRslts(2, i) = ConvertToMinutes(CSng(rs(3).Value))
				                IndRslts(3, i) = PacePerMile(rs(3).Value, sRaceName)
				                IndRslts(4, i) = PacePerKM(rs(3).Value, sRaceName)
                            Else
	                            IndRslts(2, i) = rs(3).Value
	                            IndRslts(3, i) = PacePerMile(ConvertToSeconds(rs(3).Value), sRaceName)
   	                            IndRslts(4, i) = PacePerKM(ConvertToSeconds(rs(3).Value), sRaceName)
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

                        If i = CInt(iNumAwds) Then Exit Do
                    End If
	            End If
	        End If
		                     
	        rs.MoveNext
	    Loop
	    rs.Close
	    Set rs = Nothing
    End If
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
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events (GSE) Awards Results</title>
<!--#include file = "../../includes/js.asp" -->
<meta name="description" content="Gopher State Events (GSE) Awards Results.">
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
    <ul class="list-inline">
	    <li><a href="javascript:window.print()">Print This</a></li>
        <%If sWhichTime = "gun" Then%>
            <li><a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=chip">View By Chip Time</a></li>
        <%Else%>
            <li><a href="awards.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;gender=<%=sGender%>&amp;which_time=gun">View By Gun Time</a></li>
        <%End If%>
    </ul>

    <h4 class="h4 bg-danger" style="color:#fff;">Awards for <%=sEventName%> (<%=dEventDate%>)</h4>

    <form role="form" class="form-inline" name="get_races" method="post" action="awards.asp?event_id=<%=lEventID%>">
    <label for="races">Race:</label>
    <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
        <%For i = 0 to UBound(Races, 2) - 1%>
            <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
                <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
            <%Else%>
                <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
            <%End If%>
        <%Next%>
    </select>&nbsp;&nbsp;
    <label for="gender">Gender:</label>
    <select class="form-control" name="gender" id="gender" onchange="this.form.get_race.click()">
        <%Select Case sGender%>
            <%Case "M"%>
                <option value="M" selected>Male</option>
                <option value="F">Female</option>
                <option value="X">Combined</option>
            <%Case "F"%>
                <option value="M">Male</option>
                <option value="F" selected>Female</option>
                <option value="X">Combined</option>
            <%Case "X"%>
                <option value="M">Male</option>
                <option value="F">Female</option>
                <option value="X" selected>Combined</option>
        <%End Select%>
    </select>
    <input class="form-control" type="hidden" name="submit_race" id="submit_race" value="submit_race">
    <input class="form-control" type="submit" name="get_race" id="get_race" value="View">
    </form>

    <%If sWhichTime = "gun" Then%>
        <h4 class="h4">(Gun Time)</h4>
    <%Else%>
        <h4 class="h4">(Chip Time)</h4>
    <%End If%>

    <%If UBound(OpenAwds, 2) > 0 Then%>
        <div class="bg-success">
            <h5 class="h5">Open Awards</h5>
            <table class="table">
	            <tr>
		            <th>Pl</th>
		            <th>Bib-Name</th>
		            <%If sShowAge = "y" Then%>
                        <th>Age</th>
                    <%End if%>
		            <th>Time</th>
		            <th>Per Mi</th>
		            <th>Per Km</th>
		            <th>Location</th>
	            </tr>
	            <%For i = 0 To UBound(OpenAwds, 2) - 1%>
			        <tr>
				        <td><%=i + 1%></td>
				        <td><%=OpenAwds(0, i)%></td>
		                <%If sShowAge = "y" Then%>
				            <td><%=OpenAwds(1, i)%></td>
                        <%End if%>
				        <td><%=OpenAwds(2, i)%></td>
				        <td><%=OpenAwds(3, i)%></td>
				        <td><%=OpenAwds(4, i)%></td>
				        <td><%=OpenAwds(5, i)%></td>
			        </tr>
	            <%Next%>
            </table>
        </div>
    <%End If%>

    <%For j = 0 to UBound(AgeGrps, 2) - 1%>
        <%Call GetTheseRslts(AgeGrps(0, j), AgeGrps(1, j))%>
	    <h5 class="h5"><%=AgeGrps(2, j)%></h5>
        <table class="table table-striped">
		    <tr>
			    <th>Pl</th>
			    <th>Bib-Name</th>
		        <%If sShowAge = "y" Then%>
                    <th>Age</th>
                <%End if%>
			    <th>Time</th>
			    <th>Per Mi</th>
			    <th>Per Km</th>
			    <th>From</th>
		    </tr>
		    <%For i = 0 To UBound(IndRslts, 2) - 1%>
			    <tr>
				    <td><%=i + 1%></td>
				    <td><%=IndRslts(0, i)%></td>
		            <%If sShowAge = "y" Then%>
				        <td><%=IndRslts(1, i)%></td>
                    <%End if%>
				    <td><%=IndRslts(2, i)%></td>
				    <td><%=IndRslts(3, i)%></td>
				    <td><%=IndRslts(4, i)%></td>
				    <td style="text-align:left;"><%=IndRslts(5, i)%></td>
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
    <%Next%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
