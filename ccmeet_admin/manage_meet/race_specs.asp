<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisMeet, lThisRace, lFirstTeam
Dim sMeetName, sGender, sSortOrder, sStartType, sAutoFill, sIndivRelay, iNumParts
Dim iGates, iRaceBreak, iWaveSize, iDelay
Dim MeetTeams(), Races(), StartSort(2), StartType(3), RaceSpecs(6)
Dim dMeetDate
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")
lThisRace = Request.QueryString("this_race")
lFirstTeam = Request.QueryString("this_team")

StartSort(0) = "Team"
StartSort(1) = "Bib"
StartSort(2) = "Random"

StartType(0) = "Mass"
StartType(1) = "Wave"
StartType(2) = "Interval"
StartType(3) = "Pursuit"

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

'get meet teams array
i = 0
ReDim MeetTeams(2, 0)
sql = "SELECT mt.TeamsID, t.TeamName, t.Gender FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	MeetTeams(2,  i) = rs(2).Value
	i = i + 1
	ReDim Preserve MeetTeams(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

'get races in this meet
i = 0
ReDim Races(5, 0)
sql = "SELECT RacesID, RaceDesc, Gender, StartType, RaceBreak, RaceTime, IndivRelay FROM Races WHERE MeetsID = " & lThisMeet
sql = sql & " ORDER BY ViewOrder"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = Replace(rs(1).Value, "''", "'")
    Races(2, i) = rs(3).Value
    Races(3, i) = rs(4).Value
    Races(4, i) = rs(5).Value
    Races(5, i) = rs(6).Value
	i = i + 1
	ReDim Preserve Races(5, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("submit_order") = "submit_order" Then
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
    lThisRace = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_specs") = "submit_specs" Then
    iRaceBreak = Request.Form.Item("race_break")
    If CStr(iRaceBreak) = vbNullString Then iRaceBreak = 0

    sStartType = Request.Form.Item("start_type")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceBreak, StartType FROM Races WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    rs(0).Value = iRaceBreak
    rs(1).Value = sStartType
    rs.Update
    rs.Close
    Set rs = Nothing

    sSortOrder = Request.Form.Item("sort_order")
    lFirstTeam = Request.Form.Item("first_team")
    iDelay = Request.Form.Item("delay")
    iWaveSize = Request.Form.Item("wave_size")
    sAutoFill = Request.Form.Item("auto_fill")
    iGates = Request.Form.Item("gates")

    If lFirstTeam & "" = "" Then lFirstTeam = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SortOrder, FirstTeam, IntDelay, WaveDelay, WaveSize, WaveAutoFill, Gates FROM RunOrder WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    rs(0).Value = sSortOrder
    rs(1).Value = lFirstTeam
    If sStartType = "Wave" Then
        rs(3).Value = iDelay
        rs(4).Value = iWaveSize
    Else
        rs(2).Value = iDelay
    End If
    rs(5).Value = sAutoFill
    rs(6).Value = iGates
    rs.Update
    rs.Close
    Set rs = Nothing
End If

If CStr(lThisRace) & "" = "" Then lThisRace = 0
If CStr(lFirstTeam) & "" = "" Then lFirstTeam = 0

If Not CLng(lThisRace) = 0 Then
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT Gender, StartType, RaceBreak, IndivRelay FROM Races WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    sGender = Left(rs(0).Value, 1)
    sStartType = rs(1).Value
    iRaceBreak = rs(2).Value
    sIndivRelay = rs(3).Value
    rs.Close
    Set rs = Nothing

    'get num legs
    If sIndivRelay = "Relay" Then
        bFound = False
        Set rs = SErver.CreateObject("ADODB.Recordset")
        sql = "SELECT NumParts FROM RElays WHERE RacesID = " & lThisRace
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0  Then
            iNumParts = rs(0).Value
            bFound = True
        End If
        rs.Close
        Set rs = Nothing
    End If

    If bFound = False Then
        iNumParts = 4

        sql = "INSERT INTO Relays(RacesID) VALUES (" & lThisRace & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If

    bFound = False
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SortOrder, FirstTeam, IntDelay, WaveDelay, WaveSize, WaveAutoFill, Gates FROM RunOrder WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sSortOrder = rs(0).Value

        If rs(1).Value & "" = "" Then
            lFirstTeam = 0
        Else
            lFirstTeam = rs(1).Value
        End If

        If sStartType = "Wave" Then
            iDelay = rs(3).Value
        Else
            iDelay = rs(2).Value
        End If
            
        iWaveSize = rs(4).Value
        sAutoFill = rs(5).Value
        iGates = rs(6).Value
    
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

        sql = "INSERT INTO RunOrder (RacesID) VALUES (" & lThisRace & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
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

Private Function GetTeamName(lTeamID)
    Dim x

    GetTeamName = vbNullString

    If Not lTeamID = "0" Then
        For x = 0 To UBound(MeetTeams, 2) - 1
            If CLng(lTeamID) = CLng(MeetTeams(0, x)) Then
                GetTeamName = MeetTeams(1, x)
                Exit For
            End If
        Next
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE  Admin Meet Specs</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->

			    <div style="text-align:right;background-color:#ececd8;font-size:0.85em;margin:0 0 10px 0;">
                    <a href="/ccmeet_admin/manage_meet/race_specs.asp?meet_id=<%=lThisMeet%>"">Race Specs</a>
                    &nbsp;|&nbsp;
				    <a href="races.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>">Race Data</a>
				    &nbsp;|&nbsp;
				    <a href="add_race.asp?meet_id=<%=lThisMeet%>">Add Race</a>
			    </div>
			<%End If%>
			
			<h4 class="h4">CCMeet Race Specs: <%=sMeetName%></h4>

            <form role="form" class="form-inline" name="sel_race" method="post" action="race_specs.asp?meet_id=<%=lThisMeet%>">
            <label for="races">Select Race:</label>
            <select class="form-control" name="races" id="races" onchange="this.form.submit1.click()">
                <option value="">&nbsp;</option>
                <%For i = 0 To UBound(Races, 2) - 1%>
                    <%If CLng(lThisRace) = CLng(Races(0, i)) Then%>
                        <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
                    <%Else%>
                        <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
                    <%End If%>
                <%Next%>
            </select>
            <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
            <input class="form-control" type="submit" name="submit1" id="submit1" value="Get Race">
            </form>

            <%If Not CLng(lThisRace) = 0 Then%>
                <div class="row">
                    <div class="col-sm-4">
                        <form name="set_specs" method="post" action="race_specs.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>">
                        <table>
                            <tr>
                                <th style="text-align: center;background-color: #ececd8;" colspan="2">EDIT SPECS</th>
                            </tr>
                            <tr>
                                <th>Start Type:</th>
                                <td>
                                    <select name="start_type" id="start_type">
                                        <%For i = 0 To UBound(StartType)%>
                                            <%If sStartType = StartType(i) Then%>
                                                <option value="<%=StartType(i)%>" selected><%=StartType(i)%></option>
                                            <%Else%>
                                                <option value="<%=StartType(i)%>"><%=StartType(i)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                            </td>
                            </tr>
                        <tr>
                                <th valign="top">First Team:</th>
                                <td>
                                    <select name="first_team" id="first_team">
                                        <option value="">&nbsp;</option>
                                        <%For i = 0 To UBound(MeetTeams, 2) - 1%>
                                            <%If CStr(MeetTeams(2, i)) = CStr(sGender) Then%>
                                                <%If CLng(lFirstTeam) = CLng(MeetTeams(0, i)) Then%>
                                                    <option value="<%=MeetTeams(0, i)%>" selected><%=MeetTeams(1, i)%></option>
                                                <%Else%>
                                                    <option value="<%=MeetTeams(0, i)%>"><%=MeetTeams(1, i)%></option>
                                                <%End If%>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <th>Race Break:</th>
                                <td><input type="text" name="race_break" id="race_break" size="2" value="<%=iRaceBreak%>"></td>
                            </tr>
                            <tr>
                                <th>Wave/Int Delay:</th>
                                <td><input type="text" name="delay" id="delay" size="2" value="<%=iDelay%>"></td>
                            </tr>
                            <tr>
                                <th>Gates:</th>
                                <td>
                                    <select name="gates" id="gates">
                                        <%For i = 1 To 5%>
                                            <%If CInt(iGates) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                            </td>
                            </tr>
                            <tr>
                                <th>Start Sort:</th>
                                <td>
                                    <select name="sort_order" id="sort_order">
                                        <%For i = 0 To UBound(StartSort)%>
                                            <%If sSortOrder = StartSort(i) Then%>
                                                <option value="<%=StartSort(i)%>" selected><%=StartSort(i)%></option>
                                            <%Else%>
                                                <option value="<%=StartSort(i)%>"><%=StartSort(i)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                            </td>
                            </tr>
                            <tr>
                                <th>Wave Size:</th>
                                <td>
                                    <select name="wave_size" id="wave_size">
                                        <%For i = 1 To 100%>
                                            <%If CInt(iWaveSize) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                            </td>
                            </tr>
                            <tr>
                                <th>Auto-Fill:</th>
                                <td>
                                    <select name="auto_fill" id="auto_fill">
                                        <%If sAutoFill = "Y" Then%>
                                            <option value="N">No</option>
                                            <option value="Y" selected>Yes</option>
                                        <%Else%>
                                            <option value="N">No</option>
                                            <option value="Y">Yes</option>
                                        <%End If%>
                                    </select>
                            </td>
                            </tr>
                            <%If sIndivRelay = "Relay" Then%>
                                <tr>
                                    <th>Num Legs:</th>
                                    <td>
                                        <select name="num_parts" id="num_parts">
                                            <%For i = 2 To 10%>
                                                <%If CInt(iNumParts) = CInt(i) Then%>
                                                    <option value="<%=i%>" selected><%=i%></option>
                                                <%Else%>
                                                    <option value="<%=i%>"><%=i%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                </td>
                                </tr>
                            <%End If%>
                        <tr>
                                <td style="text-align: center;" colspan="2">
                                    <input type="hidden" name="submit_specs" id="submit_specs" value="submit_specs">
                                    <input type="submit" name="submit2" id="submit2" value="Set Specs">
                                </td>
                            </tr>
                        </table>
                        </form>
                    </div>
                    <div class="col-sm-8">
                        <table class="table tarble-striped">
                            <tr>
                                <th style="text-align: center;background-color: #ececec;" colspan="11">RACE SPECS</th>
                            </tr>
                            <tr>
                                <th>Race</th>
                                <th>Time</th>
                                <th>Start</th>
                                <th>Break</th>
                                <th>Sort</th>
                                <th>First Team</th>
                                <th>Delay</th>
                                <th>Size</th>
                                <th>Auto-Fill</th>
                                <th>Gates</th>
                            </tr>
                            <%For i = 0 To UBound(Races, 2) - 1%>
                                <%Call GetRaceSpecs(Races(0, i), Races(2, i))%>

                                <tr>
                                    <td><%=Races(1, i)%></td>
                                    <td><%=Races(4, i)%></td>
                                    <td><%=Races(2, i)%></td>
                                    <td><%=Races(3, i)%></td>
                                    <td><%=RaceSpecs(0)%></td>
                                    <td><%=GetTeamName(RaceSpecs(1))%></td>
                                    <td style="text-align: center;"><%=RaceSpecs(2)%></td>
                                    <td style="text-align: center;"><%=RaceSpecs(3)%></td>
                                    <td style="text-align: center;"><%=RaceSpecs(4)%></td>
                                    <td style="text-align: center;"><%=RaceSpecs(5)%></td>
                                </tr>
                            <%Next%>
                        </table>

                        <form name="set_order" method="post" action="race_specs.asp?meet_id=<%=lThisMeet%>">
                        <input type="hidden" name="submit_order" id="submit_order" value="submit_order">
                        <input type="submit" name="submit3" id="submit3" value="Set Run Order" disabled>
                        </form>
                    </div>
                </div>
            <%End If%>
        </div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
