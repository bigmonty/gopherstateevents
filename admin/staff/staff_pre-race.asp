<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lEventID, lEventDirID
Dim sTeamScore, sEventDir, sEventDirEmail, sEventDirPhone
Dim EventArray(17), Races(), RaceInfo(16), MaleArray(), FemaleArray(), Staff()
Dim sngDeposit

lEventID = Request.QueryString("event_id")

sngDeposit = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventName, e.EventDate, s.SiteName, s.Address, e.NeedBibs, e.NeedPins, e.PacketPickup, e.Announcer, e.DigitalDisplay, e.Location, "
sql = sql & "e.LocalPower, e.TearOffs, e.AntFieldSize, s.MapLink, e.DynamicBibAssign, e.NeedTruss, e.StaffNotes, e.EventDirID "
sql = sql & "FROM Events e INNER JOIN SiteInfo s ON e.EventID = s.EventID INNER JOIN Waiver w ON w.EventID = e.EventID WHERE e.EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
	For i = 0 to 16
		If Not rs(i).Value & "" = "" Then  EventArray(i) = Replace(rs(i).Value, "''", "'")
	Next

    lEventDirID = rs(17).Value
End If
rs.Close
Set rs = Nothing

If EventArray(4) & "" = "" Then EventArray(4) = "y"
If EventArray(5) & "" = "" Then EventArray(5) = "n"
If EventArray(7) & "" = "" Then EventArray(7) = "n"
If EventArray(10) & "" = "" Then EventArray(10) = "n"
If EventArray(11) & "" = "" Then EventArray(11) = "n"
If EventArray(12) & "" = "" Then EventArray(14) = "50"

EventArray(16) = EventArray(16) & "<p class='text-danger'>PLEASE REMEMBER TO GET EXTRA BIBS-INCLUDING UN-USED PRE-REGISTERED ONES (NO SHOWS)!</p>"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RsltsSort FROM RFIDSettings WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
EventArray(17) = rs(0).Value
rs.Close
Set rs = Nothing

sql = "SELECT FirstName, LastName, Phone, Email FROM EventDir WHERE EventDirID = " & lEventDirID
Set rs = conn.Execute(sql)
sEventDir = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
sEventDirPhone = rs(2).Value
sEventDirEmail = rs(3).Value
Set rs = Nothing

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

i = 0
ReDim Staff(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT StaffID FROM StaffAsgmt WHERE EventID = " & lEventID & " AND EventType = 'Fitness Event'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Staff(i) = GetStaffName(rs(0).Value)
    i = i + 1
    ReDim Preserve Staff(i)
    rs.MoveNext
Loop
rs.Close
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
End Sub

Private Function GetStaffName(lThisStaff)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName, Phone FROM Staff WHERE StaffID = " & lThisStaff
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then 
        GetStaffName = Replace(rs2(0).Value, "''", "'") & " " & Replace(rs2(1).Value, "''", "'")
        GetStaffName = GEtSTaffName & " (" & rs2(2).Value &")"
    End If
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=Replace(EventArray(0), "''", "'")%> Event Settings</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    .list-group-item{
      margin:2px;padding:2px;  
    } 
</style>
</head>

<body>
<div class="container">
    <div class="bg-info">
        <a href="javascript:window.print();">Print</a>
    </div>

    <div class="row">
        <div class="col-md-4">
            <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
        </div>
        <div class="col-md-8">
	        <h3 class="h3">Staff Notes: <%=EventArray(0)%></h3>
        </div>
    </div>

    <div class="row">
        <div class="col-md-4 bg-danger">
            <h4 class="h4 text-danger">Staff Notes:</h4>
            <%=EventArray(16)%>

            <h5 class="h5">Event Director:</h5>                                
            <ul class="list-group">
			    <li class="list-group-item">Name: <%=sEventDir%></li>
			    <li class="list-group-item">Phone: <%=sEventDirPhone%></li>
                <li class="list-group-item">Email: <a href="mailto:<%=sEventDirEmail%>"><%=sEventDirEmail%></a></li>
            </ul>
        </div>
        <div class="col-md-4 bg-success">
            <h4 class="h4">General:</h4>                                
            <ul class="list-group">
			    <li class="list-group-item">Event Date: <%=EventArray(1)%></li>
			    <li class="list-group-item">Participants Expected: <%=EventArray(12)%></li>
                <li class="list-group-item">Packet Pickup: <%=EventArray(6)%></li>
            </ul>

            <h4 class="h4">Staff:</h4>                                
            <ul class="list-group">
			    <%For i = 0 To UBound(Staff) - 1%>
                    <li class="list-group-item"><%=Staff(i)%></li>
                <%Next%>
            </ul>

            <h4 class="h4">Venue:</h4>                                
            <ul class="list-group">
			    <li class="list-group-item">City, St: <%=EventArray(9)%></li>
			    <li class="list-group-item">Event Site: <%=EventArray(2)%></li>
			    <li class="list-group-item">Address: <a href="<%=EventArray(13)%>"><%=EventArray(3)%></a></li>
                <li class="list-group-item">Local Power: <%=EventArray(10)%></li>
            </ul>

            <h4 class="h4">Preferences:</h4>                                
            <ul class="list-group">
			    <li class="list-group-item">Sort Results By: <%=EventArray(17)%></li>
			    <li class="list-group-item">Need Bibs: <%=EventArray(4)%></li>
                <li class="list-group-item">Need Pins: <%=EventArray(5)%></li>
                <li class="list-group-item">Need Tear Off Tags: <%=EventArray(11)%></li>
            </ul>

            <h4 class="h4">Extra Features:</h4>                                
            <ul class="list-group">
                <li class="list-group-item">Dynamic Bib Assign: <%=EventArray(14)%></li>
			    <li class="list-group-item">Announcer Portal: <%=EventArray(7)%></li>
			    <li class="list-group-item">Digital Display: <%=EventArray(8)%></li>
                <li class="list-group-item">Need Truss: <%=EventArray(15)%></li>
            </ul>
        </div>
        <div class="col-md-4 bg-warning">
            <h4 class="h4">Race Settings:</h4>                                

            <%For i = 0 To UBound(Races, 2) - 1%>
                <%Call GetRaceInfo(Races(0, i))%>

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