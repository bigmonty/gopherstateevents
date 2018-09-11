<%@ Language=VBScript %>

<%
Option Explicit

Dim rs, sql, conn, rs2, sql2
Dim i, j, k
Dim lThisSite
Dim sSport, sSortBy, sBlendBy, sRaceName, sRaceDist, sMeetIDs, sMeetName, sMeetSite, sVenueIDs
Dim RsltsArr, Meets, SiteRslts()
Dim TempArr(4), BegDates(), EndDates(), MeetSites(), SortArr(3), BlendArr(2)
Dim dMeetDate, dBegDate, dEndDate
Dim bGetRslts, bNoChange

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

dBegDate = Request.QueryString("beg_date")
dEndDate = Request.QueryString("end_date")
bNoChange = False

SortArr(0) = "date"
SortArr(1) = "meet"
SortArr(2) = "site"
SortArr(3) = "perf"

BlendArr(0) = "part"
BlendArr(1) = "meet"
BlendArr(2) = "site"

Server.ScriptTimeout = 600

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'over which dates do we want to look at performances for
j = 0
ReDim BegDates(1, 0)
For i = 2005 To Year(Date)
    BegDates(0, j) = "8/1/" & i
    BegDates(1, j) = "Fall " & i
    j = j + 1
    ReDim Preserve BegDates(1, j)
Next

j = 0
ReDim EndDates(1, 0)
For i = 2006 To Year(Date) + 1
    EndDates(0, j) = "3/1/" & i
    EndDates(1, j) = "Spring " & i
    j = j + 1
    ReDim Preserve EndDates(1, j)
Next

ReDim MeetSites(2, 0)
ReDim SiteRslts(6, 0)

If Request.Form.Item("submit_filters") = "submit_filters" Then  
    'what meets and venues do we want to filter out of the view
    dBegDate = Request.Form.Item("beg_date")
    dEndDate = Request.Form.Item("end_date")
    sSortBy = Request.Form.Item("sort")
    sBlendBy = Request.Form.Item("blend")   
 
    Call GetMeets

    If IsArray(Meets) Then 
        sMeetIDs = vbNullString
        For i = 0 To UBound(Meets, 2)
            If Request.Form.Item("meet_" & Meets(0, i)) = "on" Then 
                Meets(4, i) = "n"                               'make sure this meet is set to NOT be included
            Else
                sMeetIDs = sMeetIDs & Meets(0, i) & ","
            End If
        Next

        'remove trailing commas
        If Len(sMeetIDs) > 0 Then 
            If Right(sMeetIDs, 1) = "," Then sMeetIDs = Left(sMeetIDs, Len(sMeetIDs) - 1)
        End If

        Call GetMeetSites

        sVenueIDs = vbNullString
        For i = 0 To UBound(MeetSites, 2) - 1
            If Request.Form.Item("site_" & MeetSites(0, i)) = "on" Then 
                MeetSites(2, i) = "n"
            Else
                sVenueIDs = sVenueIDs & MeetSites(0, i) & ","
            End If
        Next

        'remove trailing commas
        If Len(sVenueIDs) > 0 Then 
            If Right(sVenueIDs, 1) = "," Then sVenueIDs = Left(sVenueIDs, Len(sVenueIDs) - 1)
        End If

        bNoChange = True'make sure that we do not overwrite whe we have done here
    End If
End If

If sSortBy = vbNullString Then sSortBy = "perf"
If sBlendBy = vbNullString Then sBlendBy = "part"

If Month(Date) < 9 Then
    If CStr(dBegDate) = vbNullString Then dBegDate = "8/1/" & Year(Date) - 1
    If CStr(dEndDate) = vbNullString Then dEndDate = "3/1/" & Year(Date)
Else
    If CStr(dBegDate) = vbNullString Then dBegDate = "8/1/" & Year(Date)
    If CStr(dEndDate) = vbNullString Then dEndDate = "3/1/" & Year(Date) + 1
End If

If Not sBlendBy = "part" Then 
    'sort according to display
    If sSortBy = "date" Then
        For j = 0 To UBound(Meets, 2) - 2
            For k = j + 1 To UBound(Meets, 2) - 1
                If CDate(Meets(2, j)) < CDate(Meets(2, k)) Then
                    For i = 0 To 3
                        TempArr(i) = Meets(i, j)
                        Meets(i, j) = Meets(i, k)
                        Meets(i, k) = TempArr(i)
                    Next
                End If
            Next
        Next
    ElseIf sSortBy = "meet" Then
        For j = 0 To UBound(Meets, 2) - 2
            For k = j + 1 To UBound(Meets, 2) - 1
                If CStr(Meets(1, j)) > CStr(Meets(1, k)) Then
                    For i = 0 To 3
                        TempArr(i) = Meets(i, j)
                        Meets(i, j) = Meets(i, k)
                        Meets(i, k) = TempArr(i)
                    Next
                End If
            Next
        Next
    ElseIf sSortBy = "site" Then
        For j = 0 To UBound(Meets, 2) - 2
            For k = j + 1 To UBound(Meets, 2) - 1
                If CStr(Meets(3, j)) > CStr(Meets(3, k)) Then
                    For i = 0 To 3
                        TempArr(i) = Meets(i, j)
                        Meets(i, j) = Meets(i, k)
                        Meets(i, k) = TempArr(i)
                    Next
                End If
            Next
        Next
    End If
End If

Private Sub GetMeets()
    Dim x
    Dim sOrderBy

    Select Case sSortBy 
        Case "date"
            sOrderBy = "m.MeetDate"
        Case "meet"
            sOrderBy = "m.MeetName"
        Case Else
            sOrderBy = "m.VenuesID"
    End Select

    'first get all meets during this time to populate the list box...need a dummy field at end
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT DISTINCT m.MeetsID, m.MeetName, m.MeetDate, m.VenuesID, m.MeetHost FROM Meets m "
    sql = sql & "INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID INNER JOIN Roster r ON "
    sql = sql & "mt.TeamsID = r.TeamsID INNER JOIN IndRslts ir ON ir.MeetsID = mt.MeetsID WHERE ir.RosterID = " 
    sql = sql & Session("roster_id") & " AND ir.Place > 0 AND (m.MeetDate >= '" & dBegDate & "' AND m.MeetDate <= '" 
    sql = sql & dEndDate & "') AND Sport = '" & sSport & "' ORDER BY " & sOrderBy
    rs.Open sql, conn, 1, 2
    If Not rs.EOF Then Meets = rs.GetRows()
    rs.Close
    Set rs = Nothing

    If IsArray(Meets) Then
        For x = 0 To UBound(Meets, 2)
            Meets(4, x) = "y"  'set filter to include all meets
        Next
    End If
End Sub

Private Sub GetMeetSites()
    Dim x, y

    If IsArray(Meets) Then
        x = 0
        For y = 0 To UBound(Meets, 2)
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT VenuesID, Venue FROM Venues WHERE VenuesID = " & Meets(3, y)
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then
                If y = 0 Then
                    lThisSite = Meets(3, y) 
                    MeetSites(0, 0) = rs(0).Value
                    MeetSites(1, 0) = Replace(rs(1).Value, "''", "'")
                    MeetSites(2, 0) = "y"       'set it to include all sites initially    
                    x = x + 1
                    ReDim Preserve MeetSites(2, x)
                Else
                    If Not CLng(lThisSite) = CLng(Meets(3, y)) Then
                        lThisSite = Meets(3, y) 
                        MeetSites(0, x) = rs(0).Value
                        MeetSites(1, x) = Replace(rs(1).Value, "''", "'")
                        MeetSites(2, x) = "y"      'set it to include all sites initially    
                        x = x + 1
                        ReDim Preserve MeetSites(2, x)
                    End If
                End If
            End If
            rs.Close
            Set rs = Nothing
        Next
    End If
End Sub

Private Function GetMeetName(lMeetID)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lMeetID
    rs.Open sql, conn, 1, 2
    GetMeetName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

'get results
If Len(sMeetIDs) > 1 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.RosterID, ir.MeetsID, ir.RacesID, ir.Gate, ir.RaceTime FROM IndRslts ir INNER JOIN Meets m "
    sql = sql & "ON ir.MeetsID = m.MeetsID WHERE Place > 0 AND ir.MeetsID IN (" & sMeetIDs & ") AND m.VenuesID IN ("
    sql = sql & sVenueIDs & ") AND ir.RosterID = " & Session("roster_id")
    rs.Open sql, conn, 1, 2
    If Not rs.RecordCount = 0 Then RsltsArr = rs.GetRows()
    rs.Close
    Set rs = Nothing

    If IsArray(RsltsArr) Then
        For i = 0 To UBound(RsltsArr, 2)
            RsltsArr(3, i) = GetPlace(RsltsArr(2, i), RsltsArr(0, i))   'get place (racesid,lthismmbr)
            RsltsArr(4, i) = ConvertToSeconds(RsltsArr(4, i))
        Next

        'sort by time
        If sSortBy = "perf" Then
            For i = 0 To UBound(RsltsArr, 2) - 1
                For j = i + 1 To UBound(RsltsArr, 2)
                    If CSng(RsltsArr(4, i)) > CSng(RsltsArr(4, j)) Then
                        For k = 0 To 4
                            TempArr(k) = RsltsArr(k, i)
                            RsltsArr(k, i) = RsltsArr(k, j)
                            RsltsArr(k, j) = TempArr(k)
                        Next
                    End If
                Next
            Next
        End If

        'convert to minutes
        For i = 0 To UBound(RsltsArr, 2)
            RsltsArr(4, i) = ConvertToMinutes(Round(RsltsArr(4, i), 1))
        Next
    End If
End If

Function GetPlace(lThisRaceID)
	GetPlace = 0
	sql = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRaceID & " AND Place > 0 ORDER BY RaceTime"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		GetPlace = GetPlace + 1
		If CLng(rs(0).Value) = CLng(lThisRstrID) Then Exit Do
		rs.MoveNext
	Loop
	Set rs = Nothing
End Function

Private Sub MeetData(lMeetID)
    sMeetName = "unknown" 
    dMeetDate = "1/1/1900"
    sMeetSite = vbNullString

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT m.MeetName, m.MeetDate, v.Venue FROM Meets m INNER JOIN Venues v ON m.VenuesID = v.VenuesID "
    sql = sql & "WHERE m.MeetsID = " & lMeetID 
   	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sMeetName = Replace(rs(0).Value, "''", "'")
        dMeetDate = rs(1).Value
        sMeetSite = rs(2).Value
    End If
	rs.Close
	Set rs = Nothing
End Sub

Private Sub RaceData(lRaceID)
    sRaceName = "unknown" 
    sRaceDist = vbNullString

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RaceDesc, RaceDist, RaceUnits FROM Races WHERE RacesID = " & lRaceID 
   	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sRaceName = Replace(rs(0).Value, "''", "'")
        sRaceDist = rs(1).Value & " " & rs(2).Value
    End If
	rs.Close
	Set rs = Nothing
End Sub

Private Sub GetSiteRslts(lThisVenue)
    Dim x, y, z
    Dim SortArr(6)

    z = 0
    For x = 0 To UBound(Meets, 2)
        If CLng(Meets(3, x)) = CLng(lThisVenue) Then    'get all meets for this venue
            For y = 0 To UBound(RsltsArr, 2)
                If CLng(RsltsArr(1, y)) = CLng(Meets(0, x)) Then    'get all participants in the meet
                    Call PartData(RsltsArr(0, y))
                    Call RaceData(RsltsArr(2, y))

                    SiteRslts(0, z) = Session("my_name")
                    SiteRslts(1, z) = Session("team_name")
                    SiteRslts(2, z) = Meets(1, x)
                    SiteRslts(3, z) = sRaceName
                    SiteRslts(4, z) = sRaceDist
                    SiteRslts(5, z) = RsltsArr(3, y)
                    SiteRslts(6, z) = RsltsArr(4, y)

                    z = z + 1
                    ReDim Preserve SiteRslts(6, z)
                End If
            Next
        End If
    Next

    For x = 0 To UBound(SiteRslts, 2) - 2
        For y = x + 1 To UBound(SiteRslts, 2) - 1
            If SiteRslts(6, y) < SiteRslts(6, x) Then
                For z = 0 To 6
                    SortArr(z) = SiteRslts(z, x)
                    SiteRslts(z, x) = SiteRslts(z, y)
                    SiteRslts(z, y) = SortArr(z)
                Next
            End If
        Next
    Next
End Sub
%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Performance Tracker Performance Lists</title>
</head>

<body>
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->
    <!--#include file = "perf_trkr_nav.asp" -->

    <h4 class="h4">My Performance Tracker History</h4>

    <h5 class="h5 bg-danger" style="padding:2px;">Filters</h5>

    <form class="form" name="list_filters" method="post" action="my_hist.asp?beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>">
    <input type="hidden" name="submit_filters" id="submit_filters" value="submit_filters">
    <input class="form-control" type="submit" name="submit3" id="submit3" value="Submit Filters">
    <div class="row">
        <div class="col-sm-3">
            <div class="form-group">
                <label for="beg_date">From:</label>
                <select class="form-control" name="beg_date" id="beg_date">
                    <%For i = 0 To UBound(BegDates, 2) - 1%>
                        <%If CDate(dBegDate) = CDate(BegDates(0, i)) Then%>
                            <option value="<%=BegDates(0, i)%>" selected><%=BegDates(1, i)%></option>
                        <%Else%>
                            <option value="<%=BegDates(0, i)%>"><%=BegDates(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </div>
            <div class="form-group">
                <label for="end_date">To:</label>
                <select class="form-control" name="end_date" id="end_date">
                    <%For i = 0 To UBound(EndDates, 2) - 1%>
                        <%If CDate(dEndDate) = CDate(EndDates(0, i)) Then%>
                            <option value="<%=EndDates(0, i)%>" selected><%=EndDates(1, i)%></option>
                        <%Else%>
                            <option value="<%=EndDates(0, i)%>"><%=EndDates(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </div>
        </div>
        <div class="col-sm-3">
            <label for="sort">Sort By:</label>
            <ul class="list-group">
                <%For i = 0 To UBound(SortArr)%>
                    <%If CStr(sSortBy) = CStr(SortArr(i)) Then%>
                        <li class="list-group-item">
                            <input type="radio" name="sort" id="sort" value="<%=SortArr(i)%>" checked>
                            <%=SortArr(i)%>
                        </li> 
                    <%Else%>
                        <li class="list-group-item">
                            <input type="radio" name="sort" id="sort" value="<%=SortArr(i)%>">
                            <%=SortArr(i)%>
                        </li> 
                    <%End If%>
                <%Next%>
            </ul>

            <label for="blend">Blend By:</label>
            <ul class="list-group">
                <%For i = 0 To UBound(BlendArr)%>
                    <%If sBlendBy = BlendArr(i) Then%>
                        <li class="list-group-item">
                            <input type="radio" name="blend" id="blend" value="<%=BlendArr(i)%>" checked>
                            <%=BlendArr(i)%>
                        </li> 
                    <%Else%>
                        <li class="list-group-item">
                            <input type="radio" name="blend" id="blend" value="<%=BlendArr(i)%>">
                            <%=BlendArr(i)%>
                        </li> 
                    <%End If%>
                <%Next%>
            </ul>
        </div>
        <div class="col-sm-3">
            <label>Sites (select to exclude)</label>
            <ul class="list-group">
                <%For i = 0 To UBound(MeetSites, 2) - 1%>
                    <%If MeetSites(2, i) = "n" Then%>
                        <li class="list-group-item">
                            <input type="checkbox" name="site_<%=MeetSites(0, i)%>" id="site_<%=MeetSites(0, i)%>" checked>
                            <%=MeetSites(1, i)%>
                        </li>
                    <%Else%>
                        <li class="list-group-item">
                            <input type="checkbox" name="site_<%=MeetSites(0, i)%>" id="site_<%=MeetSites(0, i)%>">
                            <%=MeetSites(1, i)%>
                        </li>
                    <%End If%>
                <%Next%>
            </ul>
        </div>
        <div class="col-sm-3">
            <label>Meets (select to exclude)</label>

            <%If IsArray(Meets) Then%>
                <ul class="list-group">
                    <%For i = 0 To UBound(Meets, 2) - 1%>
                        <%If Meets(4, i) = "n" Then%>
                            <li class="list-group-item">
                                <input type="checkbox" name="meet_<%=Meets(0, i)%>" id="meet_<%=Meets(0, i)%>" checked>
                                <%=Meets(1, i)%>
                            </li>
                        <%Else%>
                            <li class="list-group-item">
                                <input type="checkbox" name="meet_<%=Meets(0, i)%>" id="meet_<%=Meets(0, i)%>">
                                <%=Meets(1, i)%>
                            </li>
                        <%End If%>
                    <%Next%>
                </ul>
            <%End If%>
        </div>
        </form>
    </div>

    <div class="row">
        <%Select Case sBlendBy%>
            <%Case "part"%>
                <ul class="list-group">
                    <li class="list-group-item">
                        <%=Session("my_name")%> (<%=Session("team_name")%>)

                        <%If IsArray(RsltsArr) Then%>
                            <table class="table table-striped">
                                <tr>
                                    <th>Meet</th>
                                    <th>Date</th>
                                    <th>Site</th>
                                    <th>Race</th>
                                    <th>Dist</th>
                                    <th>Pl</th>
                                    <th>Time</th>
                                </tr>
                                <%For j = 0 To UBound(RsltsArr, 2)%>
                                    <%Call MeetData(RsltsArr(1, j))%>
                                    <%Call RaceData(RsltsArr(2, j))%>
                                    <tr>
                                        <td><%=sMeetName%></td>
                                        <td><%=dMeetDate%></td>
                                        <td><%=sMeetSite%></td>
                                        <td><%=sRaceName%></td>
                                        <td><%=sRaceDist%></td>
                                        <td><%=RsltsArr(3, j)%></td>
                                        <td><%=RsltsArr(4, j)%></td>
                                    </tr>
                                <%Next%>
                            </table>
                        <%End If%>
                    </li>
                </ul>
            <%Case "meet"%>
                <%If IsArray(Meets) Then%>
                    <ul class="list-group">
                        <%For i = 0 To UBound(Meets, 2)%>
                            <%If Meets(4, i) = "y" Then%>
                                <%Call MeetData(Meets(0, i))%>
                                <li class="list-group-item">
                                    <%=sMeetName%>&nbsp;on&nbsp;<%=dMeetDate%> (<%=sMeetSite%>)
            
                                    <%If IsArray(RsltsArr) Then%>
                                        <table class="table table-striped">
                                            <tr>
                                                <th>Name</th>
                                                <th>Team</th>
                                                <th>Race</th>
                                                <th>Dist</th>
                                                <th>Pl</th>
                                                <th>Time</th>
                                            </tr>
                                            <%For j = 0 To UBound(RsltsArr, 2)%>
                                                <%Call RaceData(RsltsArr(2, j))%>

                                                <%If CLng(Meets(0, i)) = CLng(RsltsArr(1, j)) Then%>
                                                    <tr>
                                                        <td><%=Session("my_name")%></td>
                                                        <td><%=Session("team_name")%></td>
                                                        <td><%=sRaceName%></td>
                                                        <td><%=sRaceDist%></td>
                                                        <td><%=RsltsArr(3, j)%></td>
                                                        <td><%=RsltsArr(4, j)%></td>
                                                    </tr>
                                                <%End If%>
                                            <%Next%>
                                        </table>
                                    <%End If%>
                                </li>
                            <%End If%>
                        <%Next%>
                    </ul>
                <%End If%>
            <%Case "site"%>
                <%If UBound(MeetSites, 2) > 0 Then%>
                    <ul class="list-group">
                        <%For i = 0 To UBound(MeetSites, 2) - 1%>
                            <%If MeetSites(2, i) = "y" Then%>
                                <%Call GetSiteRslts(MeetSites(0, i))%>
                                <li class="list-group-item">
                                    <%=MeetSites(1, i)%>
        
                                    <table class="table table-striped">
                                        <tr>
                                            <th>Name</th>
                                            <th>Team</th>
                                            <th>Meet</th>
                                            <th>Race</th>
                                            <th>Dist</th>
                                            <th>Pl</th>
                                            <th>Time</th>
                                        </tr>
                                        <%If UBound(SiteRslts, 2) > 0 Then%>
                                            <%For j = 0 To UBound(SiteRslts, 2) - 1%>
                                                <tr>
                                                    <td><%=SiteRslts(0, j)%></td>
                                                    <td><%=SiteRslts(1, j)%></td>
                                                    <td><%=SiteRslts(2, j)%></td>
                                                    <td><%=SiteRslts(3, j)%></td>
                                                    <td><%=SiteRslts(4, j)%></td>
                                                    <td><%=SiteRslts(5, j)%></td>
                                                    <td><%=SiteRslts(6, j)%></td>
                                                </tr>
                                            <%Next%>
                                        <%End If%>
                                    </table>
                                </li>
                            <%End If%>
                        <%Next%>
                    </ul>
                <%End If%>
        <%End Select%>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.close
Set conn = Nothing
%>
</html>
