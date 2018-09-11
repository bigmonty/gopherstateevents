<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID
Dim sTeamScore
Dim EventArray(21), Races(), RaceInfo(16), MaleArray(), FemaleArray()
Dim sngDeposit

lEventID = Request.QueryString("event_id")

sngDeposit = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventName, e.EventDate, s.SiteName, e.OnlineReg, s.Address, e.WhenShutdown, e.NeedBibs, e.NeedPins, e.OptOut, e.PacketPickup, "
sql = sql & "e.Announcer, e.DigitalDisplay, e.Location, e.LocalPower, e.TearOffs, e.AntFieldSize, s.MapLink, e.DynamicBibAssign, e.NeedTruss, "
sql = sql & "e.PixSponsor, e.Comments FROM Events e INNER JOIN SiteInfo s ON e.EventID = s.EventID INNER JOIN Waiver w ON w.EventID = e.EventID "
sql = sql & "WHERE e.EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
	For i = 0 to 15
		If Not rs(i).Value & "" = "" Then  EventArray(i) = Replace(rs(i).Value, "''", "'")
	Next

	EventArray(17) = rs(16).Value
	EventArray(18) = rs(17).Value
    EventArray(19) = rs(18).Value
    EventArray(20) = rs(19).Value
    If Not rs(20).Value & "" = "" Then EventArray(21) = Replace(rs(20).Value, "''", "'")
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RsltsSort FROM RFIDSettings WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
EventArray(16) = rs(0).Value
rs.Close
Set rs = Nothing

If EventArray(6) & "" = "" Then EventArray(6) = "y"
If EventArray(7) & "" = "" Then EventArray(7) = "n"
If EventArray(9) & "" = "" Then EventArray(9) = "n"
If EventArray(15) & "" = "" Then EventArray(15) = "50"

i = 0
ReDim Races(1, 0)
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub GetRaceInfo(lThisRace)
    Dim x, y

    sql = "SELECT RaceName, Dist, StartTime, Certified, StartType, MAwds, FAwds, OnlineRegLink, AllowDuplAwds, "
    sql = sql & "EntryFeePre, EntryFee, ChipStart, NumSplits, StartToFinish, NumLaps, IndivRelay FROM RaceData WHERE RaceID = " & lThisRace
    Set rs = conn.Execute(sql)
    RaceInfo(0) = rs(0).Value
    RaceInfo(1) = rs(1).Value
	
    'split the time field
    RaceInfo(2) = Left(rs(2).Value, Len(rs(2).Value) - 2)
    RaceInfo(3) = Right(rs(2).Value, 2)
	
    RaceInfo(4) = rs(3).Value
    RaceInfo(5) = rs(4).Value
    RaceInfo(6) = rs(5).Value
    RaceInfo(7) = rs(6).Value
    RaceInfo(8) = rs(7).Value
    RaceInfo(9) = rs(8).Value
    RaceInfo(10) = rs(9).Value
    RaceInfo(11) = rs(10).Value
    RaceInfo(12) = rs(11).Value
    RaceInfo(13) = rs(12).Value
    RaceInfo(14) = rs(13).Value
    RaceInfo(15) = rs(14).Value
    RaceInfo(16) = rs(15).Value
    Set rs = Nothing

    If RaceInfo(13) & "" = "" Then RaceInfo(13) = 0
    If RaceInfo(14) & "" = "" Then RaceInfo(14) = 0

    sTeamScore = "n"
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM TeamScoring WHERE RaceID = " & lThisRace
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sTeamScore = "y"
    rs.Close
    Set rs = Nothing

    'get male age group array
    x = 0
    ReDim MaleArray(2, 0)
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT AgeGroupsID, EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lThisRace
    sql = sql & " AND Gender = 'm' ORDER BY EndAge"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    For y = 0 to 2
		    MaleArray(y, x) = rs(y).Value
	    Next
		
	    x = x + 1
	    ReDim Preserve MaleArray(2, x)
		
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get female age group array
    x = 0
    ReDim FemaleArray(2, 0)
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT AgeGroupsID, EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lThisRace
    sql = sql & " AND Gender = 'f' ORDER BY EndAge"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    For y = 0 to 2
		    FemaleArray(y, x) = rs(y).Value
	    Next
		
	    x = x + 1
	    ReDim Preserve FemaleArray(2, x)
		
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=Replace(EventArray(0), "''", "'")%> Event Settings</title>
</head>

<body>
<div class="container">
    <div class="bg-info">
        <a href="javascript:window.print();">Print</a>
    </div>

    <div class="row">
        <div class="col-sm-5">
            <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
        </div>
        <div class="col-sm-7">
	        <h3 class="h3">GSE Event Settings: </h3>
	        <h4 class="h4"><%=EventArray(0)%></h4>
        </div>
 
        <p class="bg-warning">
            IMPORTANT NOTE:  Below is the data that we will be using to manage your upcoming event.  PLEASE CHECK IT OVER
            CAREFULLY to ensure it is correct.  We go to great lengths to provide the necessary staff and equipment to ensure that all of our events are
            primed for success but it hinges on this information!  We can not be held accountable for incorrect information that we are not made aware of
            in a timely fashion.  (Please pay special attention to date, time, and location!) 
        </p>
   </div>

    <div class="row">
        <div class="col-sm-4 bg-success">
            <h4 class="h4">General:</h4>                                
            <ul class="list-group">
                <li class="list-group-item">Event Date: <%=EventArray(1)%></li>
                <li class="list-group-item">Participants Expected: <%=EventArray(15)%></li>
                <li class="list-group-item">Packet Pickup: <%=EventArray(9)%></li>
            </ul>

            <%If Not EventArray(21) & "" = "" Then%>
                <h4 class="h4">Message:</h4>                                
                <p><%=EventArray(21)%></p>
            <%End If%>

            <h4 class="h4">Venue:</h4>                                
            <ul class="list-group">
                <li class="list-group-item">City, St: <%=EventArray(12)%></li>
                <li class="list-group-item">Event Site: <%=EventArray(2)%></li>
                <li class="list-group-item">Address: <a href="javascript:pop('<%=EventArray(17)%>',800,750)"><%=EventArray(4)%></a></li>
                <li class="list-group-item">Local Power: <%=EventArray(13)%></li>
            </ul>

            <h4 class="h4">Preferences:</h4>                                
            <ul class="list-group">
                <li class="list-group-item">Sort Results By: <%=EventArray(16)%></li>
                <li class="list-group-item">Email Opt-Out: <%=EventArray(8)%></li>
                <li class="list-group-item">Need Bibs: <%=EventArray(6)%></li>
                <li class="list-group-item">Need Pins: <%=EventArray(7)%></li>
                <li class="list-group-item">Need Tear Off Tags: <%=EventArray(14)%></li>
            </ul>

            <h4 class="h4">Registration:</h4>                                
            <ul class="list-group">
                <li class="list-group-item">Online Part Reg: <%=EventArray(3)%></li>
                <li class="list-group-item">End Pre-Reg: <%=EventArray(5)%></li>
                <li class="list-group-item">Dynamic Bib Assign: <%=EventArray(18)%></li>
            </ul>

            <h4 class="h4">Extra Features:</h4>                                
            <ul class="list-group">
                <li class="list-group-item">Announcer Portal: <%=EventArray(10)%></li>
                <li class="list-group-item">Digital Display: <%=EventArray(11)%></li>
                <li class="list-group-item">Need Truss: <%=EventArray(19)%></li>
                <li class="list-group-item">Pix Sponsor: <%=EventArray(20)%></li>
            </ul>
        </div>
        <div class="col-sm-8">
            <h4 class="h4">Race Settings:</h4>                                

            <%For i = 0 To UBound(Races, 2) - 1%>
                <div class="row">
                <%Call GetRaceInfo(Races(0, i))%>

                    <div class="col-sm-4">
                        <h5 class="h5"><%=Races(1, i)%></h5>
                        <ul class="list-group">
                            <li class="list-group-item">Distance: <%=RaceInfo(1)%></li>
                            <li class="list-group-item">Certified: <%=RaceInfo(4)%></li>
                            <li class="list-group-item">Start Type: <%=RaceInfo(5)%></li>
                            <li class="list-group-item">Start Time: <%=RaceInfo(2)%></li>
                            <li class="list-group-item">Chip Start: <%=RaceInfo(12)%></li>
                            <li class="list-group-item">Num Splits: <%=RaceInfo(13)%></li>
                            <li class="list-group-item">Num Laps: <%=RaceInfo(15)%></li>
                            <li class="list-group-item">Indiv/Relay: <%=RaceInfo(16)%></li>
                            <li class="list-group-item">Has Teams: <%=sTeamScore%></li>
                            <li class="list-group-item">Start-to-Finish (yds): <%=RaceInfo(14)%></li>
                            <li class="list-group-item">Allow Duplicate Awds: <%=RaceInfo(9)%></li>
                            <li class="list-group-item">Male Open Awds: <%=RaceInfo(6)%></li>
                            <li class="list-group-item">Female Open Awds: <%=RaceInfo(7)%></li>
                        </ul>
                    </div>
                    <div class="col-sm-4">
                        <h5 class="h5">Male Age Groups</h5>
                        
                        <%If UBound(MaleArray, 2) > 1 Then%>
                            <ul class="list-group">
                                <li class="list-group-item">Age Group (Awds)</li>
                                <%For j = 0 to UBound(MaleArray, 2) - 1%>
                                        <li class="list-group-item">
                                            <%If j = "0" Then%>
                                                <%=MaleArray(1, j)%>  And Under
                                            <%ElseIf MaleArray(1, j) = "110" Then%>
                                                <%=CInt(MaleArray(1, j - 1)) + 1%>  And Over
                                            <%Else%>
                                                <%=CInt(MaleArray(1, j - 1)) + 1%> - <%=MaleArray(1, j)%>
                                            <%End If%>
                                        (<%=MaleArray(2, j)%>)</li>
                                <%Next%>
                            </ul>
                        <%Else%>
                            <p>No age groups entered.</p>
                        <%End If%>
                    </div>
                    <div class="col-sm-4">
                        <h5 class="h5">Female Age Groups</h5>
                        
                        <%If UBound(FemaleArray, 2) > 1 Then%>
                            <ul class="list-group">
                                <li class="list-group-item">Age Group (Awds)</li>
                                <%For j = 0 to UBound(FemaleArray, 2) - 1%>
                                        <li class="list-group-item">
                                            <%If j = "0" Then%>
                                                <%=FemaleArray(1, j)%>  And Under
                                            <%ElseIf FemaleArray(1, j) = "110" Then%>
                                                <%=CInt(FemaleArray(1, j - 1)) + 1%>  And Over
                                            <%Else%>
                                                <%=CInt(FemaleArray(1, j - 1)) + 1%> - <%=FemaleArray(1, j)%>
                                            <%End If%>
                                        (<%=FemaleArray(2, j)%>)</li>
                                <%Next%>
                            </ul>
                        <%Else%>
                            <p>No age groups entered.</p>
                        <%End If%>
                    </div>
                    <hr>
                </div>
            <%Next%>
        </div>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>