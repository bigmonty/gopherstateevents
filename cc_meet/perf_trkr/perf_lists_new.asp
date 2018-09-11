<%@ Language=VBScript %>

<%
Option Explicit

Dim rs, sql, conn2, rs2, sql2
Dim i, j, k
Dim lPacksID, lRosterID
Dim sPackName, sSport, sGender, sMeetName, sSite, sSortBy, sBlendBy, sFilter, sSelectQS, sCompOnly, sThisSite
Dim sShowFilters, sExclSites, sExclMeets, sRosterIDs, sMeetIDs, sAllMeetIDs
Dim MyPacks, PackMmbrs, ViewPerf, RsltsArr, Meets
Dim TempArr(7), SelUsers(), MeetRslts(), SiteRslts(), BegDates(), EndDates(), MeetSites()
Dim dWhenCreated, dMeetDate, dBegDate, dEndDate
Dim bGetRslts

Dim dLoadStart, dLoadEnd, dLoadTime

dLoadStart = Now()

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lPacksID = Request.QueryString("packs_id")
lRosterID = Request.QueryString("roster_id")
sSortBy = Request.QueryString("sort_by")
sBlendBy = Request.QueryString("blend_by")
sSelectQS = Request.QueryString("select_qs")
sExclSites = Request.QueryString("exclude_sites")
sExclMeets = Request.QueryString("exclude_meets")
sCompOnly = Request.QueryString("comp_only")
dBegDate = Request.QueryString("beg_date")
dEndDate = Request.QueryString("end_date")

sShowFilters = Request.QueryString("show_filters")
If sShowFilters = vbNullString Then sShowFilters = "y"

Server.ScriptTimeout = 600

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get list of selected pack members for a query
If Not sSelectQS = vbNullString Then 
    sSelectQS = Replace(sSelectQS, ";", ",")

    If Right(sSelectQS, 1) = "," Then sSelectQS = Left(sSelectQS, Len(sSelectQS) - 1)

    Call GetPerf(sSelectQS)
End If

'over which dates do we want to look at performances for
j = 0
ReDim BegDates(1, 0)
For i = 2005 To Year(Date)
    BegDates(0, j) = "8/1/" & i
    BegDates(1, j) = "Fall of " & i
    j = j + 1
    ReDim Preserve BegDates(1, j)
Next

j = 0
ReDim EndDates(1, 0)
For i = 2006 To Year(Date) + 1
    EndDates(0, j) = "3/1/" & i
    EndDates(1, j) = "Spring of " & i
    j = j + 1
    ReDim Preserve EndDates(1, j)
Next

'list of my performance packs for possible viewing...not much happening here but how do I escape the "'" with getrows?
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PerfTrkrPacksID, PackName FROM PerfTrkrPacks WHERE PerfTrkrID = " & Session("perf_trkr_id")
rs.Open sql, conn2, 1, 2
If Not rs.EOF Then MyPacks = rs.GetRows()
rs.Close
Set rs = Nothing

If IsArray(MyPacks) Then
    If CStr(lPacksID) = vbNullString Then lPacksID = 0
    If lPacksID = "0" Then lPacksID = MyPacks(0, 0)
    If CLng(lPacksID) > 0 Then Call GetPackData(lPacksID)
End If

ReDim MeetSites(1, 0)

If Request.Form.Item("submit_filters") = "submit_filters" Then  
    'what meets and venues do we want to filter out of the view
    dBegDate = Request.Form.Item("beg_date")
    dEndDate = Request.Form.Item("end_date")

    Call GetMeets
    
    If IsArray(Meets) Then 
        Call GetMeetSites

        sExclSites = vbNullString

        j = 0
        For i = 0 To UBound(MeetSites, 2) - 1
            If Request.Form.Item("site_" & MeetSites(0, i)) = "on" Then 
                MeetSites(1, i) & "y"
                sExclSites = sExclSites & MeetSites(1, 0) & "," 'to use as a filter in the query
            End If
        Next

        If Len(sExclSites) > 0 Then
            If Right(sExclSites, 1) = "," Then sExclSites = Left(sExclSites, Len(sExclSites) - 1)
        End If

        j = 0
        For i = 0 To UBound(Meets, 2) - 1
            If Request.Form.Item("meet_" & Meets(0, i)) = "on" Then 
                sExclMeets = sExclMeets & Meets(0, i) & ","
                Meets(4, i) = "y"
            End If
        Next

        If Len(sExclMeets) > 0 Then
            If Right(sExclMeets, 1) = "," Then sExclMeets = Left(sExclMeets, Len(sExclMeets) - 1)
        End If
    End If
ElseIf Request.Form.Item("submit_select") = "submit_select" Then
    'specifically which pack members will we view performances for...not sure how to use get rows here when we are possibly
    'only selecting some pack members
    sSelectQS = vbNullString

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PTPackMmbrsID FROM PTPackMmbrs WHERE PerfTrkrPacksID = " & lPacksID
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("select_all") = "on" Then
            sSelectQS = sSelectQS & rs(0).Value & ","   'for the print page
        Else
            If Request.Form.Item("select_" & rs(0).Value) = "on" Then
                sSelectQS = sSelectQS & rs(0).Value & ","   'for the print page
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Right(sSelectQS, 1) = "," Then sSelectQS = Left(sSelectQS, Len(sSelectQS) - 1)

    Call GetPerf(sSelectQS)
ElseIf Request.Form.Item("submit_pack") = "submit_pack" Then
    'which pack do you want to view members of for selection
    lPacksID = Request.Form.Item("my_packs")
    If CStr(lPacksID) = vbNullString Then lPacksID = Packs(0, 0)
    If CLng(lPacksID) > 0 Then Call GetPackData(lPacksID)
End If

If CStr(lPacksID) = vbNullString Then lPacksID = 0
If CStr(lRosterID) = vbNullString Then lRosterID = 0
If sSortBy = vbNullString Then sSortBy = "date"
If sBlendBy = vbNullString Then sBlendBy = "part"
If sCompOnly = vbNullString Then sCompOnly = "n"

If Month(Date) < 9 Then
    If CStr(dBegDate) = vbNullString Then dBegDate = "8/1/" & Year(Date) - 1
    If CStr(dEndDate) = vbNullString Then dEndDate = "3/1/" & Year(Date)
Else
    If CStr(dBegDate) = vbNullString Then dBegDate = "8/1/" & Year(Date)
    If CStr(dEndDate) = vbNullString Then dEndDate = "3/1/" & Year(Date) + 1
End If

If Not sSelectQS = vbNullString Then
    Call GetMeets
    If IsArray(Meets) Then Call GetMeetSites
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

Private Sub GetPackData(lThisPack)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PackName, Sport, Gender, WhenCreated FROM PerfTrkrPacks WHERE PerfTrkrPacksID = " & lPacksID
    rs.Open sql, conn2, 1, 2
    sPackName = Replace(rs(0).Value, "''", "'")
    sSport = rs(1).Value
    sGender = rs(2).Value
    dWhenCreated = rs(3).Value
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT pt.PTPackMmbrsID, pt.RosterID, r.FirstName, r.LastName, r.TeamsID FROM PTPackMmbrs pt INNER JOIN Roster r ON pt.RosterID = r.RosterID "
    sql = sql & "WHERE pt.PerfTrkrPacksID = " & lPacksID & " ORDER BY r.LastName, r.FirstName"
    rs.Open sql, conn2, 1, 2
    If Not rs.EOF Then PackMmbrs = rs.GetRows()
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetPerf(sTheseParts)
    Dim x

    'this determines who to view performances for
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT pt.RosterID, r.FirstName, r.LastName, r.TeamsID FROM PTPackMmbrs pt INNER JOIN Roster r ON pt.RosterID = r.RosterID "
    sql = sql & "WHERE pt.PerfTrkrPacksID = " & lPacksID & " AND pt.PTPackMmbrsID IN (" & sSelectQS & ") ORDER BY r.LastName, r.FirstName"
    rs.Open sql, conn2, 1, 2
    If Not rs.EOF Then ViewPerf = rs.GetRows()
    rs.Close
    Set rs = Nothing

    'get roster ids for reqults queries
    sRosterIDs = vbNullString
    For x = 0 To UBound(ViewPerf, 2)
        sRosterIDs = sRosterIDs & ViewPerf(0, x) & ","
    Next

    If Len(sRosterIDs) > 0 Then sRosterIDs = Left(sRosterIDs, Len(sRosterIDs) - 1)
End Sub

Private Sub GetMeets()
    Dim x

    'first get all meets during this time to populate the list box...need a dummy field at end
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT DISTINCT m.MeetsID, m.MeetName, m.MeetDate, m.MeetSite, m.MeetHost FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID INNER JOIN Roster r ON "
    sql = sql & "mt.TeamsID = r.TeamsID INNER JOIN IndRslts ir ON ir.MeetsID = mt.MeetsID WHERE ir.RosterID IN (" & sRosterIDs & ") AND ir.Place > 0 "
    sql = sql & "AND (m.MeetDate >= '" & dBegDate & "' AND m.MeetDate <= '" & dEndDate & "') AND Sport = '" & sSport & "' ORDER BY MeetSite DESC"
    rs.Open sql, conn2, 1, 2
    If Not rs.EOF Then Meets = rs.GetRows()
    rs.Close
    Set rs = Nothing

    If IsArray(Meets) Then
        For x = 0 To UBound(Meets, 2)
            Meets(4, x) = "n"  'make this an exclude field by setting it to null right now
            sMeetIDs = Meets(0, x) & ","
        Next
    End If

    If Len(sMeetIDs) > 0 Then sMeetIDs = Left(sMeetIDs, Len(sMeetIDs) - 1)
End Sub

Private Sub GetMeetSites()
    Dim x, y

    If IsArray(Meets) Then
        x = 0
        For y = 0 To UBound(Meets, 2)
            If y = 0 Then
                sThisSite = Meets(3, y) 
                MeetSites(0, 0) = sThisSite
                MeetSites(1, 0) = "n"
                x = x + 1
                ReDim Preserve MeetSites(1, x)
            Else
                If Not CStr(sThisSite) = CStr(Meets(3, y)) Then
                    sThisSite = Meets(3, y) 
                    MeetSites(0, x) = sThisSite
                    MeetSites(1, x) = "n"
                    x = x + 1
                    ReDim Preserve MeetSites(1, x)
                End If
            End If
        Next
    End If
End Sub

Private Function GetTeamName(lTeamID)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
    rs2.Open sql2, conn2, 1, 2
    GetTeamName = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

If Len(sMeetIDs) > 1 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID, MeetsID, RacesID, Gate, RaceTime FROM IndRslts WHERE Place > 0 AND MeetsID IN (" 
    sql = sql & sMeetIDs & ") "
    rs.Open sql, conn2, 1, 2
    If Not rs.RecordCount = 0 Then RsltsArr = rs.GetRows()
    rs.Close
    Set rs = Nothing

    If IsArray(RsltsArr) Then
        For x = 0 To UBound(RsltsArr, 2)
            RsltsArr(3, x) = GetPlace(RsltsArr(2, x), RsltsArr(0, x))   'get place (racesid,lthismmbr)
            RsltsArr(4, x) = ConvertToSeconds(RsltsArr(4, x))
        Next

        'sort by time
        For x = 0 To UBound(RsltsArr, 2) - 1
            For y = x + 1 To UBound(RsltsArr, 2)
                If CSng(RsltsArr(4, x)) > CSng(RsltsArr(4, y)) Then
                    For z = 0 To 4
                        TempArr(z) = RsltsArr(z, x)
                        RsltsArr(z, x) = RsltsArr(z, y)
                        RsltsArr(z, y) = TempArr(z)
                    Next
                End If
            Next
        Next

        'convert to minutes
        For x = 0 To UBound(RsltsArr, 2)
            RsltsArr(4, x) = ConvertToMinutes(Round(RsltsArr(4, x), 1))
        Next
    End If
End If

Function GetPlace(lThisRaceID, lThisRstrID)
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

    <h4 class="h4">My Performance Tracker Lists</h4>

    <form class="form-inline" name="get_pack" method="post" action="perf_lists.asp">
    <div class="form-group">
        <label for="my_packs">Select a Pack:</label>
        <select class="form-control" name="my_packs" id="my_packs" onchange="this.form.submit1.click();">
            <option value="">&nbsp;</option>
            <%For i = 0 To UBound(MyPacks, 2)%>
                <%If CLng(lPacksID) = CLng(MyPacks(0, i)) Then%>
                    <option value="<%=MyPacks(0, i)%>" selected><%=MyPacks(1, i)%></option>
                <%Else%>
                    <option value="<%=MyPacks(0, i)%>"><%=MyPacks(1, i)%></option>
                <%End If%>
            <%Next%>
        </select>
    </div>
    <div class="form-group">
        <input type="hidden" name="submit_pack" id="submit_pack" value="submit_pack">
        <input class="form-control" type="submit" name="submit1" id="submit1" value="Get Pack">
    </div>
    </form>

    <%If Not CLng(lPacksID) = 0 Then%>
        <br>
        <h5 class="h5"><%=sPackName%>&nbsp;&nbsp;<%=sSport%>&nbsp;<%=sGender%> <span style="font-weight: normal;">(created <%=dWhenCreated%>)</span></h5>

    <div class="row">
        <div class="col-sm-3">   
            <h5 class="h5 bg-warning" style="padding:2px;">Pack Members</h5>

            <form class="form" name="select_mmbrs" method="post" action="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>">
            <ul class="list-group">
                <li class="list-group-item" style="font-weight: bold;"><input type="checkbox" name="select_all" id="select_all">&nbsp;Select All</li>
                <%For i = 0 To UBound(PackMmbrs, 2)%>
                    <%If IsArray(ViewPerf) Then%>
                        <%For j = 0 To UBound(ViewPerf, 2)%>
                            <%If CLng(PackMmbrs(0, i)) = CLng(ViewPerf(0, j)) Then%>
                                <li class="list-group-item">
                                    <input type="checkbox" name="select_<%=PackMmbrs(0, i)%>" id="select_<%=PackMmbrs(0, i)%>" checked>
                                    <%=PackMmbrs(3, i)%>, <%=PackMmbrs(2, i)%>&nbsp;(<%=GetTeamName(PackMmbrs(4, i))%>)
                                </li>

                                <%Exit For%>
                            <%Else%>
                                <%If j = UBound(ViewPerf, 2) Then%>
                                    <li class="list-group-item">
                                        <input type="checkbox" name="select_<%=PackMmbrs(0, i)%>" id="select_<%=PackMmbrs(0, i)%>">
                                        <%=PackMmbrs(3, i)%>, <%=PackMmbrs(2, i)%>&nbsp;(<%=GetTeamName(PackMmbrs(4, i))%>)
                                    </li>
                                <%End If%>
                            <%End If%>
                        <%Next%>
                    <%Else%>
                        <li class="list-group-item">
                            <input type="checkbox" name="select_<%=PackMmbrs(0, i)%>" id="select_<%=PackMmbrs(0, i)%>">
                            <%=PackMmbrs(3, i)%>, <%=PackMmbrs(2, i)%>(<%=GetTeamName(PackMmbrs(4, i))%>)
                        </li>
                    <%End If%>
                <%Next%>
            </ul>
            <div class="form-group">
                <input type="hidden" name="submit_select" id="submit_select" value="submit_select">
                <input class="form-control" type="submit" name="submit2" id="submit2" value="Select Participant(s)">
            </div>
            </form>
        </div>            
        <div class="col-sm-9">
            <%If IsArray(ViewPerf) Then%>
                <h5 class="h5 bg-danger" style="padding:2px;">Filters</h5>

                <div>
                    <%If sShowFilters = "n" Then%>
                        <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=y&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;sort_by=<%=sSortBy%>&amp;select_qs=<%=sSelectQS%>">Show Filters</a>
                    <%Else%>
                        <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=n&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;sort_by=<%=sSortBy%>&amp;select_qs=<%=sSelectQS%>">Hide Filters</a>
                    <%End If%>

                    <%If sShowFilters = "y" Then%>
                        <form class="form" name="list_filters" method="post" 
                            action="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=y&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;sort_by=<%=sSortBy%>&amp;select_qs=<%=sSelectQS%>">
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
                            <div class="col-sm-4">
                                <label>Sites (select to exclude)</label>
                                <ul class="list-group">
                                    <%For i = 0 To UBound(MeetSites) - 1%>
                                        <%If UBound(ExclSites) > 0 Then%>
                                            <%For j = 0 To UBound(ExclSites) - 1%>
                                                <%If CStr(MeetSites(i)) = Replace(CStr(ExclSites(j)), "'", "") Then%>
                                                    <li class="list-group-item">
                                                        <input type="checkbox" name="site_<%=MeetSites(i)%>" id="site_<%=MeetSites(i)%>" 
                                                        checked> <%=MeetSites(i)%>
                                                    </li>

                                                    <%Exit For%>
                                                <%Else%>
                                                    <%If j = UBound(ExclSites) - 1 Then%>
                                                        <li class="list-group-item">
                                                            <input type="checkbox" name="site_<%=MeetSites(i)%>" id="site_<%=MeetSites(i)%>">
                                                            <%=MeetSites(i)%>
                                                        </li>
                                                    <%End If%>
                                                <%End If%>
                                            <%Next%>
                                        <%Else%>
                                            <li class="list-group-item">
                                                <input type="checkbox" name="site_<%=MeetSites(i)%>" id="site_<%=MeetSites(i)%>">
                                                <%=MeetSites(i)%>
                                            </li>
                                        <%End If%>
                                    <%Next%>
                                </ul>
                            </div>
                            <div class="col-sm-5">
                                <label>Meets (select to exclude)</label>

                                <%If IsArray(Meets) Then%>
                                    <ul class="list-group">
                                        <%For i = 0 To UBound(Meets, 2) - 1%>
                                            <%If UBound(ExclMeets) > 0 Then%>
                                                <%For j = 0 To UBound(ExclMeets) - 1%>
                                                    <%If CLng(Meets(0, i)) = CLng(ExclMeets(j)) Then%>
                                                        <li class="list-group-item">
                                                            <input type="checkbox" name="meet_<%=Meets(0, i)%>" id="meet_<%=Meets(0, i)%>" 
                                                            checked> <%=Meets(1, i)%>
                                                        </li>

                                                        <%Exit For%>
                                                    <%Else%>
                                                        <%If j = UBound(ExclMeets) - 1 Then%>
                                                            <li class="list-group-item">
                                                                <input type="checkbox" name="meet_<%=Meets(0, i)%>" id="meet_<%=Meets(0, i)%>">
                                                                <%=Meets(1, i)%>
                                                            </li>
                                                        <%End If%>
                                                    <%End If%>
                                                <%Next%>
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
                        </div>
                        <input type="hidden" name="submit_filters" id="submit_filters" value="submit_filters">
                        <input class="form-control" type="submit" name="submit3" id="submit3" value="Submit Filters">
                        </form>
                    <%End If%>
                </div>

                <br>

                <div class="row">
                    <div class="col-sm-4" style="font-size: 0.8em;">
                        <span style="font-weight:bold;">Sort By:</span>&nbsp;&nbsp;
                        <%If sSortBy = "date" Then%>
                            <a style="font-weight:bold;" href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Date</a>
                            &nbsp;|&nbsp;
                        <%Else%>
                            <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Date</a>
                            &nbsp;|&nbsp;
                        <%End If%>
                        <%If sSortBy = "meet" Then%>
                            <a style="font-weight:bold;" href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;sort_by=meet&amp;select_qs=<%=sSelectQS%>">Meet</a>
                        <%Else%>
                            <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;sort_by=meet&amp;select_qs=<%=sSelectQS%>">Meet</a>
                        <%End If%>

                        <%If Not sBlendBy = "site" Then%>
                            <%If sSortBy = "site" Then%>
                                &nbsp;|&nbsp;
                                <a style="font-weight:bold;" href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;sort_by=site&amp;select_qs=<%=sSelectQS%>">Site</a>                    
                            <%Else%>
                                &nbsp;|&nbsp;
                                <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;sort_by=site&amp;select_qs=<%=sSelectQS%>">Site</a>
                            <%End If%>
                        <%End If%>
                        <%If Not sBlendBy = "meet" Then%>
                            <%If sSortBy = "perf" Then%>
                                &nbsp;|&nbsp;
                                <a style="font-weight:bold;" href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;sort_by=perf&amp;select_qs=<%=sSelectQS%>">Perf</a>
                            <%Else%>
                                &nbsp;|&nbsp;
                                <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;sort_by=perf&amp;select_qs=<%=sSelectQS%>">Perf</a>
                            <%End If%>
                        <%End If%>
                    </div>
                    <div class="col-sm-4" style="font-size: 0.8em;">
                        <span style="font-weight:bold;">Blend By:</span>&nbsp;&nbsp;
                        <%If sBlendBy = "part" Then%>
                            <a style="font-weight:bold;" href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Particip</a>
                            &nbsp;|&nbsp;
                        <%Else%>
                            <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Particip</a>
                            &nbsp;|&nbsp;
                        <%End If%>
                        <%If sBlendBy = "meet" Then%>
                            <a style="font-weight:bold;" href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;blend_by=meet&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Meet</a>
                            &nbsp;|&nbsp;
                        <%Else%>
                            <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;blend_by=meet&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Meet</a>
                            &nbsp;|&nbsp;
                        <%End If%>
                        <%If sBlendBy = "site" Then%>
                            <a style="font-weight:bold;" href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;blend_by=site&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Site</a>
                        <%Else%>    
                            <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;blend_by=site&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Site</a>
                        <%End If%>
                    </div>
                    <div class="col-sm-4" style="font-size: 0.8em;">
                        <%If Not sBlendBy = "part" Then%>
                            <%If sCompOnly="n" Then%>
                                <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=y&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Comparisons Only</a>
                                &nbsp;|&nbsp;
                            <%Else%>
                                <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=n&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Show All</a>
                                &nbsp;|&nbsp;
                            <%End If%>
                        <%End If%>
                        <a href="perf_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;show_filters=<%=sShowFilters%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>">Refresh</a>
                        &nbsp;|&nbsp;
                        <a href="javascript:pop('print_lists.asp?excl_sites=<%=sExclSites%>&amp;excl_meets=<%=sExclMeets%>&amp;beg_date=<%=dBegDate%>&amp;end_date=<%=dEndDate%>&amp;comp_only=<%=sCompOnly%>&amp;packs_id=<%=lPacksID%>&amp;sort_by=<%=sSortBy%>&amp;blend_by=<%=sBlendBy%>&amp;filter=<%=sFilter%>&amp;select_qs=<%=sSelectQS%>',800,750)">Print</a>
                    </div>
                </div>

                <%Select Case sBlendBy%>
                    <%Case "part"%>
                        <ol class="list-group">
                            <%For i = 0 To UBound(ViewPerf, 2)%>
                                <%Call MyResults(ViewPerf(0, i))%>
                                <li class="list-group-item">
                                    <%=ViewPerf(2, i)%>,&nbsp;<%=ViewPerf(1, i)%> (<%=GetTeamName(ViewPerf(3, i))%>)
            
                                    <%If IsArray(RsltsArr) Then%>
                                        <table class="table table-striped">
                                            <tr>
                                                <th>No.</th>
                                                <th>Meet</th>
                                                <th>Date</th>
                                                <th>Site</th>
                                                <th>Race</th>
                                                <th>Dist</th>
                                                <th>Pl</th>
                                                <th>Time</th>
                                            </tr>
                                            <%For j = 0 To UBound(RsltsArr, 2)%>
                                                <tr>
                                                    <td style="text-align:right;"><%=j + 1%>)</td>
                                                    <td><%=RsltsArr(0, j)%></td>
                                                    <td><%=RsltsArr(1, j)%></td>
                                                    <td><%=RsltsArr(2, j)%></td>
                                                    <td><%=RsltsArr(3, j)%></td>
                                                    <td><%=RsltsArr(4, j)%>&nbsp;<%=RsltsArr(5, j)%></td>
                                                    <td><%=RsltsArr(6, j)%></td>
                                                    <td><%=RsltsArr(7, j)%></td>
                                                </tr>
                                            <%Next%>
                                        </table>
                                    <%End If%>
                                </li>
                            <%Next%>
                        </ol>
                    <%Case "meet"%>
                        <ol class="list-group">
                            <%For i = 0 To UBound(Meets, 2) - 1%>
                                <%Call GetMeetRslts(Meets(0, i))%>
                                <%If sCompOnly = "y" Then%>
                                    <%If UBound(MeetRslts, 2) > 1 Then%>
                                        <li class="list-group-item">
                                            <%=Meets(1, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>&nbsp; - &nbsp;<%=Meets(3, i)%>)</span>
                
                                            <table class="table table-striped">
                                                <tr>
                                                    <th>No.</th>
                                                    <th style="text-align:left;">Name</th>
                                                    <th style="text-align:left;">Team</th>
                                                    <th style="text-align:left;">Race</th>
                                                    <th style="text-align:left;">Dist</th>
                                                    <th>Pl</th>
                                                    <th>Time</th>
                                                </tr>
                                                <%For j = 0 To UBound(MeetRslts, 2) - 1%>
                                                    <tr>
                                                        <td style="text-align:right;"><%=j + 1%>)</td>
                                                        <%For k = 0 To 5%>
                                                            <td><%=MeetRslts(k, j)%></td>
                                                        <%Next%>
                                                    </tr>
                                                <%Next%>
                                            </table>
                                        </li>
                                    <%End If%>
                                <%Else%>
                                    <li class="list-group-item">
                                        <%=Meets(1, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>&nbsp; - &nbsp;<%=Meets(3, i)%></span>)
                
                                        <table class="table table-striped">
                                            <tr>
                                                <th>No.</th>
                                                <th style="text-align:left;">Name</th>
                                                <th style="text-align:left;">Team</th>
                                                <th style="text-align:left;">Race</th>
                                                <th style="text-align:left;">Dist</th>
                                                <th>Pl</th>
                                                <th>Time</th>
                                            </tr>
                                            <%For j = 0 To UBound(MeetRslts, 2) - 1%>
                                                <tr>
                                                    <td style="text-align:right;"><%=j + 1%>)</td>
                                                    <%For k = 0 To 5%>
                                                        <td><%=MeetRslts(k, j)%></td>
                                                    <%Next%>
                                                </tr>
                                            <%Next%>
                                        </table>
                                    </li>
                                <%End If%>
                            <%Next%>
                        </ol>
                    <%Case "site"%>
                            <ol class="list-group">
                            <%For i = 0 To UBound(Meets, 2) - 1%>
                                <%bGetRslts = False%>
                                <%If i = 0 Then%>
                                    <%sThisSite = Meets(3, i)%>
                                    <%bGetRslts = True%>
                                <%ElseIf Not CStr(sThisSite) = CStr(Meets(3, i)) Then%>
                                    <%sThisSite = Meets(3, i)%>
                                    <%bGetRslts = True%>
                                <%End If%> 
                    
                                <%If bGetRslts = True Then%>   
                                    <%Call GetSiteRslts(sThisSite)%>
                                    <%If sCompOnly = "y" Then%>
                                        <%If UBound(SiteRslts, 2) > 1 Then%>
                                            <li class="list-group-item">
                                                <%=Meets(3, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>)</span>
                
                                                <table class="table table-striped">
                                                    <tr>
                                                        <th>No.</th>
                                                        <th>Name</th>
                                                        <th>Team</th>
                                                        <th>Meet</th>
                                                        <th>Race</th>
                                                        <th>Dist</th>
                                                        <th>Pl</th>
                                                        <th>Time</th>
                                                    </tr>
                                                    <%For j = 0 To UBound(SiteRslts, 2) - 1%>
                                                        <tr>
                                                            <td style="text-align:right;"><%=j + 1%>)</td>
                                                            <%For k = 0 To 6%>
                                                                <td><%=SiteRslts(k, j)%></td>
                                                            <%Next%>
                                                        </tr>
                                                    <%Next%>
                                                </table>
                                            </li>
                                        <%End If%>
                                    <%Else%>
                                        <li class="list-group-item">
                                            <%=Meets(3, i)%> &nbsp; <span style="font-weight: normal;">(<%=Meets(2, i)%>)</span>
                
                                            <table class="table table-striped">
                                                <tr>
                                                    <th>No.</th>
                                                    <th>Name</th>
                                                    <th>Team</th>
                                                    <th>Meet</th>
                                                    <th>Race</th>
                                                    <th>Dist</th>
                                                    <th>Pl</th>
                                                    <th>Time</th>
                                                </tr>
                                                <%For j = 0 To UBound(SiteRslts, 2) - 1%>
                                                    <tr>
                                                        <td style="text-align:right;"><%=j + 1%>)</td>
                                                        <%For k = 0 To 6%>
                                                            <td><%=SiteRslts(k, j)%></td>
                                                        <%Next%>
                                                    </tr>
                                                <%Next%>
                                            </table>
                                        </li>
                                    <%End If%>
                                <%End If%>
                            <%Next%>
                        </ol>
                <%End Select%>
            <%End If%>
        </div>
    <%End If%>
</div>
<%
dLoadEnd = Now()
dLoadTime = DateDiff("s", dLoadStart, dLoadEnd)
Response.Write dLoadTime
%>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn2.close
Set conn2 = Nothing
%>
</html>
