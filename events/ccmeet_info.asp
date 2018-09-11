<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim sMeetName, dMeetDate, sMeetSite, sMeetHost, sWebSite, sComments, sSport, sStartList, sLogo, sMeetRaces
Dim lMeetID, sMeetDir, sMeetDirEmail
Dim MTeams(), FTeams(), Races(), SortArr(6), RaceSpecs(6)
Dim sMapLink
Dim i, j, k
Dim iNumMale, iNumFemale, iNumParts
Dim sCourseMap
Dim sMeetInfoSheet
Dim bFound

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect "htttp://www.google.com"
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

iNumParts = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT IndRsltsID FROM IndRslts WHERE MeetsID = " & lMeetID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iNumParts = rs.RecordCount
rs.Close
Set rs = Nothing

iNumMale = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ir.IndRsltsID FROM IndRslts ir INNER JOIN Roster r ON r.RosterID = ir.RosterID WHERE r.Gender = 'M' AND ir.MeetsID = " & lMeetID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iNumMale = rs.RecordCount
rs.Close
Set rs = Nothing

iNumFemale = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ir.IndRsltsID FROM IndRslts ir INNER JOIN Roster r ON r.RosterID = ir.RosterID WHERE r.Gender = 'F' AND ir.MeetsID = " & lMeetID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iNumFemale = rs.RecordCount
rs.Close
Set rs = Nothing

'get meet info
ReDim Races(6, 0)
ReDim MTeams(1, 0)
ReDim FTeams(1, 0)
If Not CLng(lMeetID) = 0 Then
	sql = "SELECT MeetName, MeetDate, MeetSite, MeetHost, WebSite, Comments, Sport, StartList, Logo FROM Meets WHERE MeetsID = " & lMeetID
	Set rs = conn.Execute(sql)
	If Not rs(0).Value & "" = "" Then sMeetName = Replace(rs(0).Value, "''", "'")
	dMeetDate = rs(1).Value
	If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
	If Not rs(2).Value & "" = "" Then sMeetHost = Replace(rs(3).Value, "''", "'")
	sWebSite = rs(4).Value
	If Not rs(5).Value & "" = "" Then sComments = Replace(rs(5).Value, "''", "'")
    sSport = rs(6).Value
    sStartList = rs(7).Value
    sLogo= rs(8).Value
	Set rs = Nothing

	'get meet dir info
	sql = "SELECT md.FirstName, md.LastName, md.Email FROM MeetDir md INNER JOIN Meets m "
	sql = sql & "ON md.MeetDirID = m.MeetDirID WHERE m.MeetsID = " & lMeetID
	Set rs = conn.Execute(sql)
	sMeetDir = rs(0).Value & " " & rs(1).Value
	sMeetDirEmail = rs(2).Value
	Set rs = Nothing
	
	'get participating teams info	
	i = 0
	sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
	sql = sql & "WHERE mt.MeetsID = " & lMeetID & " AND Gender = 'M' ORDER BY t.TeamName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		MTeams(0, i) = rs(0).Value
        MTeams(1, i) = Replace(rs(1).Value, "''", "'")
		i = i + 1
		ReDim Preserve MTeams(1, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
	sql = sql & "WHERE mt.MeetsID = " & lMeetID & " AND Gender = 'F' ORDER BY t.TeamName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		FTeams(0, i) = rs(0).Value
        FTeams(1, i) = Replace(rs(1).Value, "''", "'")
		i = i + 1
		ReDim Preserve FTeams(1, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	'get race information
	i = 0
    sMeetRaces = vbNullString
    sql = "SELECT RacesID, RaceDesc, StartType, RaceBreak, RaceTime, IndivRelay, RaceDist, RaceUnits FROM Races WHERE MeetsID = " & lMeetID
    sql = sql & " ORDER BY ViewOrder"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        sMeetRaces = sMeetRaces & rs(0).Value & ","
	    Races(0, i) = rs(0).Value
	    Races(1, i) = Replace(rs(1).Value, "''", "'")
        Races(2, i) = rs(2).Value
        Races(3, i) = ConvertToMinutes(rs(3).Value)
        Races(4, i) = rs(4).Value
        Races(5, i) = rs(5).Value
        Races(6, i) = rs(6).Value & " " & rs(7).Value
	    i = i + 1
	    ReDim Preserve Races(6, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing	

    If Len(sMeetRaces) > 0 Then sMeetRaces = Left(sMeetRaces, Len(sMeetRaces) - 1)

    For i = 0 To UBOund(Races, 2) - 2
        For j = i + 1 To UBound(RAces, 2) - 1
            If CDate(Races(4, i)) > CDate(Races(4, j)) Then
                For k = 0 To 6
                    SortArr(k) = Races(k, i)
                    Races(k, i) = Races(k, j)
                    Races(k, j) = SortArr(k)
                Next
            End If
        Next
    Next
	'get maplink
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lMeetID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMapLink = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT InfoSheet FROM MeetInfo WHERE MeetsID = " & lMeetID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMeetInfoSheet = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get course map
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Map FROM CourseMap WHERE MeetsID = " & lMeetID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sCourseMap = rs(0).Value
	rs.Close
	Set rs = Nothing
End If

Private Sub GetRaceSpecs(lRaceID, sRaceStart)
    bFound = False
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SortOrder, FirstTeam, IntDelay, WaveDelay, WaveSize, WaveAutoFill, Gates FROM RunOrder WHERE RacesID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        RaceSpecs(0) = rs(0).Value
        If rs(1).Value & "" = "" Then
            RaceSpecs(1) = 0
        Else
            RaceSpecs(1) = rs(1).Value
        End If

        If sRaceStart = "Wave" Then
            RaceSpecs(2) = rs(3).Value
        Else
            RaceSpecs(2) = rs(2).Value
        End If
            
        RaceSpecs(3) = rs(4).Value
        RaceSpecs(4) = rs(5).Value
        RaceSpecs(5) = rs(6).Value

        bFound = True
    End If
    rs.Close
    Set rs = Nothing
 
    If bFound = False Then
        RaceSpecs(0) = "Team"
        RaceSpecs(1) = "0"
        RaceSpecs(2) = 15
        RaceSpecs(3) = 5
        RaceSpecs(4) = "Y"
        RaceSpecs(5) = 1

        sql = "INSERT INTO RunOrder (RacesID, FirstTeam) VALUES (" & lRaceID & ", 0)"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
End Sub

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If sglScnds <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(sglScnds \ 60)
    sSecondPart = CStr(((sglScnds / 60) - (sglScnds \ 60)) * 60)
    
    'add leading zero to seconds if necessary
    If CSng(sSecondPart) < 10 Then
        sSecondPart = "0" & sSecondPart
    End If
    
    'make sure there are exactly two decimal places
    If Len(sSecondPart) < 5 Then
        If Len(sSecondPart) = 2 Then
            sSecondPart = sSecondPart & ".00"
        ElseIf Len(sSecondPart) = 4 Then
            sSecondPart = sSecondPart & "0"
        End If
    Else
        sSecondPart = Left(sSecondPart, 5)
    End If
    
    'do the conversion
    If CInt(sMinutePart) <= 60 Then
        ConvertToMinutes = sMinutePart & ":" & sSecondPart
    Else
        sHourPart = CStr(CSng(sMinutePart) \ 60)
        sMinutePart = CStr(CSng(sMinutePart) Mod 60)

        If Len(sMinutePart) = 1 Then
            sMinutePart = "0" & sMinutePart
        End If

        ConvertToMinutes = sHourPart & ":" & sMinutePart & ":" & sSecondPart

        ConvertToMinutes = Replace(ConvertToMinutes, "-", "")
    End If
End Function

Private Function GetRaceParts(lThisRace)
    GetRaceParts = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM IndRslts WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetRaceParts = rs.RecordCount
    rs.Close
    Set rs = Nothing
End Function

Private Function GetTeamParts(lThisTeam)
    GetTeamParts = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID IN (" 
    sql = sql & sMeetRaces & ") AND r.TeamsID = " & lThisTeam
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetTeamParts = rs.RecordCount
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE Cross-Country/Nordic Ski Meet Info</title>
<meta name="description" content="Gopher State Events cross-country running and Nordic ski event information page.">
</head>
<body>
<div class="container">
    <div style="margin: 0;padding: 0;">
        <img src="/graphics/html_header.png" class="img-responsive" alt="Individual Results">
	    <h4 class="h4"> <%=sMeetName%><br><small><%=dMeetDate%></small></h4>
    </div>
					
    <%If Date >= CDate(dMeetDate) Then%>
        <a href="/results/cc_rslts/cc_rslts.asp?sport=<%=sSport%>&meet_id=<%=lMeetID%>&rslts_page=overall_rslts.asp" 
            style="position: absolute;top: 20px;left: 760px;">
            <img src="http://www.gopherstateevents.com/graphics/race_results.jpg" alt="View" style="height: 60px;">
        </a>
	<%End If%>

    <div class="row">
	    <div class="col-xs-4">
            <%If Not sLogo & "" = "" Then%>
                <img src="/events/logos/<%=sLogo%>" class="img-responsive" alt="Logo">
            <%End If%>

            <h5 class="h5">Meet Director: <%=sMeetDir%></h5>
		    <table class="table table-condensed">
			    <tr>
                    <th>Meet Host:</th>
                    <td><%=sMeetHost%></td>
                </tr>
			    <tr>
				    <th>Meet Site:</th>
				    <td><%=sMeetSite%></td>
			    </tr>
			    <tr>
				    <td colspan="2">
                        <%If Not sWebSite = vbNullString Then%>
						    <a href="<%=sWebSite%>" target="_blank">
						        <img src="http://www.gopherstateevents.com/graphics/social_media/web_logo.jpg" alt="View" style="height: 75px;">
                            </a>
                        <%End If%>
					    <%If Not sMapLink = vbNullString Then%>
                            <a href="<%=sMapLink%>" target="_blank">
                                <img src="http://www.gopherstateevents.com/graphics/social_media/map_quest.jpg" alt="View" style="height: 75px;">
                            </a>
					    <%End If%>
					    <%If Not sCourseMap = vbNullString Then%>
						    <a href="../cc_meet/meet_dir/meets/meet_info/course_maps/<%=sCourseMap%>" target="_blank">
                                <img src="http://www.gopherstateevents.com/graphics/social_media/course_map.jpg" alt="View" style="height: 60px;">
                            </a>
					    <%End If%>
					    <%If Not sMeetInfoSheet = vbNullString Then%>
                            <a href="http://www.gopherstateevents.com/cc_meet/meet_dir/meets/meet_info/info_sheets/<%=sMeetInfoSheet%>" target="_blank">
                                <img src="http://www.gopherstateevents.com/graphics/social_media/info_logo.jpg" alt="View" style="height: 60px;">
                            </a>
					    <%End If%>
					    <%If sSport = "Nordic Ski" Then%>
						    <a href="http://www.gopherstateevents.com/misc/change_order.pdf" target="_blank">
                                <img src="http://www.gopherstateevents.com/graphics/social_media/form.jpg" alt="View" style="height: 60px;">
                            </a>
                        <%End If%>
					    <%If Not sStartList = vbNullString Then%>
						    <a href="/ccmeet_admin/manage_meet/run_order/<%=sStartList%>" target="_blank">
                                <img src="http://www.gopherstateevents.com/graphics/social_media/list.png" alt="View" style="height: 50px;">
                            </a>
					    <%End If%>
                    </td>
			    </tr>
			    <%If Not sComments = vbNullString Then%>
				    <tr>
					    <th>Comments:</th>
					    <td><%=sComments%></td>
				    </tr>
			    <%End If%>
		    </table>
	    </div>
			
        <div class="col-xs-8">
 		    <h4 class="h4">Race Schedule (<%=iNumParts%> Participants)</h4>
            <%If sSport = "Nordic Ski" Then%>
                <p class="bg-success">(Race start times MAY BE approximations if they are determined by the start time and the number of skiers in the race(s) 
                    prior to them, as well as the break between races as determined by the coaches.)</p>
            <%End If%>
		    <table class="table table-striped">
			    <tr>
				    <td>Race</td>
				    <td>Time</td>
				    <td>Dist</td>
                    <td>Participants</td>
			    </tr>
			    <%For i = 0 to UBound(Races, 2) - 1%>
				    <tr>
					    <td>
						    <a href="javascript:pop('race_details.asp?race_id=<%=Races(0, i)%>&amp;meet_id=<%=lMeetID%>',600,700)"><%=Races(1, i)%></a>
					    </td>
					    <td><%=Races(4, i)%></td>
                        <td><%=Races(6, i)%></td>
                        <td><%=GetRaceParts(Races(0, i))%></td>
				    </tr>
			    <%Next%>
		    </table>

            <h4 class="h4">Participating Teams</h4>
            <table class="table">
                <tr>
                    <td valign="top">
				        <h5 class="h5">Female (<%=UBound(FTeams, 2)%> Teams, <%=iNumFemale%> Participants)</h5>
				        <ul class="list-group">
					        <%For i = 0 to UBound(FTeams, 2) - 1%>
						        <li class="list-group-item"><%=FTeams(1, i)%> (<%=GetTeamParts(FTeams(0, i))%>)</li>
					        <%Next%>
				        </ul>
                    </td>
                    <td valign="top">
				        <h5 class="h5">Male (<%=UBound(MTeams, 2)%> Teams, <%=iNumMale%> Participants)</h5>
				        <ul  class="list-group">
					        <%For i = 0 to UBound(MTeams, 2) - 1%>
						        <li  class="list-group-item"><%=MTeams(1, i)%> (<%=GetTeamParts(MTeams(0, i))%>)</li>
					        <%Next%>
				        </ul>
			        </td>
		        </tr>
	        </table>
        </div>
    </div>

    <%If sSport = "Nordic Ski" Then%>
        <h4 class="h4">Race Specs</h4>
        <table class="table table-striped">
            <tr>
                <th>Race</th>
                <th>Race Break</th>
                <th>Indiv/Relay</th>
                <th>Start Type</th>
                <th>Wave Delay</th>
                <th>Wave Size</th>
                <th>Auto-Fill</th>
                <th>Num Gates</th>
            </tr>
            <%For i = 0 To UBound(Races, 2) - 1%>
                <%Call GetRaceSpecs(Races(0, i), Races(2, i))%>

                <tr>
                    <td><%=Races(1, i)%></td>
                    <td><%=Races(3, i)%></td>
                    <td><%=Races(5, i)%></td>
                    <td><%=Races(2, i)%></td>
                    <%If Races(2, i) = "Mass" Then%>
                        <td colspan="4">n/a (mass start)</td>
                    <%ElseIf Races(2, i) = "Pursuit" Then%>
                        <td colspan="3">n/a (pursuit)</td>
                        <td><%=RaceSpecs(5)%></td>
                    <%ElseIf Races(2, i) = "Interval" Then%>
                        <td><%=RaceSpecs(2)%></td>
                        <td colspan="2">n/a (interval)</td>
                        <td><%=RaceSpecs(5)%></td>
                    <%Else%>
                        <td><%=RaceSpecs(2)%></td>
                        <td><%=RaceSpecs(3)%></td>
                        <td><%=RaceSpecs(4)%></td>
                        <td><%=RaceSpecs(5)%></td>
                    <%End If%>
                </tr>
            <%Next%>
        </table>
    <%End If%>
<!--#include file = "../includes/footer.asp" --> 
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
