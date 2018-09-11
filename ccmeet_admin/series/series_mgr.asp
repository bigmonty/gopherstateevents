<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lWhichSeries
Dim sSeriesName, sSport, sComments, sRankBy
Dim iYear, iMaxPts
Dim Series(), Meets(), SeriesMeets(), Add(), Remove(), MeetRaces()
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lWhichSeries = Request.QueryString("which_series")
If CStr(lWhichSeries) = vbNullString Then lWhichSeries = 0

Response.Buffer = False		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.item("submit_races") = "submit_races" Then
    Dim CurrRaces()

    'get meets
    Call GetSeriesMeets()

    'get races in each meet of this series
    For i = 0 To UBound(SeriesMeets, 2) - 1
        Call GetMeetRaces(SeriesMeets(0, i), SeriesMeets(3, i))

        For j = 0 To UBound(MeetRaces, 2) - 1
            If MeetRaces(2, j) = "x" Then   'if its there but not checked delete it
                If Not Request.Form.Item("race_" & MeetRaces(0, j)) = "on" Then
                    sql = "DELETE FROM CCSeriesRaces WHERE RacesID = " & MeetRaces(0, j) & " AND CCSeriesMeetsID = " & SeriesMeets(3, i)
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing

                    MeetRaces(2, j) = vbNullString
                End If
            Else    'if it's not there but checked add it
                If Request.Form.Item("race_" & MeetRaces(0, j)) = "on" Then
                    sql = "INSERT INTO CCSeriesRaces (CCSeriesMeetsID, RacesID, RaceName) VALUES (" & SeriesMeets(3, i) & ", " & MeetRaces(0, j)
                    sql = sql & ", '" & RaceDesc(MeetRaces(0, j)) & "')"
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing

                    MeetRaces(2, j) = "x"
               End If
            End If
        Next
    Next
ElseIf Request.Form.item("submit_remove") = "submit_remove" Then
    i = 0
    ReDim Remove(0)
    sql = "SELECT MeetsID FROM CCSeriesMeets WHERE CCSeriesID = " & lWhichSeries
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        If Request.Form.Item("series_meet_" & rs(0).Value) = "on" Then
            Remove(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Remove(i)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Remove) - 1
        sql = "DELETE FROM CCSeriesMeets WHERE MeetsID = " & Remove(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
ElseIf Request.Form.item("submit_meet") = "submit_meet" Then
    Call GetSeriesInfo()

    i = 0
    ReDim Add(3, 0)
    sql = "SELECT MeetsID, MeetName, MeetDate, Location FROM Meets WHERE MeetDate >= '7/1/" & iYear & "' AND MeetDate <= '6/30/" & iYear + 1 
    sql = sql & "' AND Sport = '" & sSport & "'"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        If Request.Form.Item("meet_" & rs(0).Value) = "on" Then
            Add(0, i) = rs(0).Value
            Add(1, i) = Replace(rs(1).Value, "''", "'")
            Add(2, i) = rs(2).Value
            If Not rs(3).Value & "" = "" Then Add(3, i) = Replace(rs(3).Value, "''", "'")
            i = i + 1
            ReDim Preserve Add(3, i)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Add, 2) - 1
        sql = "INSERT INTO CCSeriesMeets (CCSeriesID, MeetsID, MeetName, MeetDate, Location) VALUES (" & lWhichSeries & ", " & Add(0, i) & ", '"
        sql = sql & Add(1, i) & "', '" & Add(2, i) & "', '" & Add(3, i) & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
ElseIf Request.Form.item("submit_series") = "submit_series" Then
    lWhichSeries = Request.Form.Item("series")
    If CLng(lWhichSeries) = vbNullString Then lWhichSeries = 0
ElseIf Request.Form.item("submit_changes") = "submit_changes" Then
	sSeriesName = Replace(Request.Form.Item("series_name"), "''", "'")
	sSport =  Request.Form.Item("sport")
	iMaxPts =  Request.Form.Item("max_pts")
    iYear =  Request.Form.Item("year")
	sComments =  Replace(Request.Form.Item("comments"), "''", "'")
    sRankBy = Request.Form.Item("rank_by")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName, Sport, MaxPts, Comments, RankBy FROM CCSeries WHERE CCSeriesID = " & lWhichSeries
    rs.Open sql, conn, 1, 2
    If sSeriesName & "" = "" Then
        rs(0).Value = rs(0).OriginalValue
    Else
        rs(0).Value = sSeriesName
    End If
    rs(1).Value = sSport
    If iMaxPts & "" = "" Then
        rs(2).Value = rs(2).OriginalValue
    Else
        rs(2).Value = iMaxPts
    End If
    rs(3).Value = sComments
    rs(4).Value = sRankBy
    rs.Update
    rs.Close
    Set rs = Nothing
End If

i = 0
ReDim Series(1, 0)
sql = "SELECT CCSeriesID, SeriesName, Sport, SeriesYear FROM CCSeries ORDER BY SeriesYear DESC, Sport, CCSeriesID DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Series(0, i) = rs(0).Value
	Series(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & " " & rs(3).Value & ")"
	i = i + 1
	ReDim Preserve Series(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not CLng(lWhichSeries) = 0 Then
    Call GetSeriesInfo()

    Call GetSeriesMeets()

    i = 0
    ReDim Meets(2, 0)
    sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate >= '7/1/" & iYear & "' AND MeetDate <= '6/30/" & iYear + 1 
    sql = sql & "' AND Sport = '" & sSport & "' ORDER BY MeetDate DESC"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        bFound = False

        If UBound(SeriesMeets, 2) > 0 Then
            For j = 0 To UBound(SeriesMeets, 2) - 1
                If CLng(rs(0).Value) = CLng(SeriesMeets(0, j)) Then
                    bFound = True
                    Exit For
                End If
            Next
        Else
            bFound = True

	        Meets(0, i) = rs(0).Value
	        Meets(1, i) = rs(1).Value
	        Meets(2, i) = rs(2).Value
	        i = i + 1
	        ReDim Preserve Meets(2, i)
        End If

        If bFound = False Then
	        Meets(0, i) = rs(0).Value
	        Meets(1, i) = rs(1).Value
	        Meets(2, i) = rs(2).Value
	        i = i + 1
	        ReDim Preserve Meets(2, i)
        End If

	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function RaceDesc(lThisRace)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceDesc FROM Races WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    RaceDesc = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function

Private Sub GetSeriesInfo()
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName, Sport, MaxPts, Comments, SeriesYear, RankBy FROM CCSeries WHERE CCSeriesID = " & lWhichSeries
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    sSport = rs(1).Value
    iMaxPts = rs(2).Value
    If Not rs(3).Value & "" = "" Then sComments = Replace(rs(3).Value, "''", "'")
    iYear = rs(4).Value
    sRankBy = rs(5).Value
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetSeriesMeets()
    Dim x

    x = 0
    ReDim SeriesMeets(3, 0)
    sql = "SELECT sm.MeetsID, m.MeetName, m.MeetDate, sm.CCSeriesMeetsID FROM CCSeriesMeets sm INNER JOIN Meets m ON sm.MeetsID = m.MeetsID "
    sql = sql & "WHERE sm.CCSeriesID = " & lWhichSeries & " ORDER BY m.MeetDate DESC"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    SeriesMeets(0, x) = rs(0).Value
	    SeriesMeets(1, x) = rs(1).Value
	    SeriesMeets(2, x) = rs(2).Value
        SeriesMeets(3, x) = rs(3).Value
	    x = x + 1
	    ReDim Preserve SeriesMeets(3, x)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetMeetRaces(lThisMeet, lSeriesMeetsID)
    Dim x

    x = 0
    ReDim MeetRaces(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MeetRaces(0, x) = rs(0).Value
        MeetRaces(1, x) = rs(1).Value
        x = x + 1
        ReDim Preserve MeetRaces(2, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'check which ones are in the series
    For x = 0 To UBound(MeetRaces, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RacesID FROM CCSeriesRaces WHERE CCSeriesMeetsID = " & lSeriesMeetsID & " AND RacesID = " & MeetRaces(0, x)
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then MeetRaces(2, x) = "x"
        rs.Close
        Set rs = Nothing
    Next
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CC/Nordic Series Manager</title>

</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h2 style="margin-left:10px;">CC/Nordic Series Manager</h2>

            <!--#include file = "cc_series_nav.asp" -->

            <h4 class="h4">Select Series To Manage</h4>

            <form role="form" class="form-inline" name="get_series" method="post" action="series_mgr.asp">
            <select class="form-control" name="series" id="series" onchange="this.form.submit1.click();">
                <option value="0">&nbsp;</option>
                <%For i = 0 To UBound(Series, 2) - 1%>
                    <%If CLng(lWhichSeries) = CLng(Series(0, i)) Then%>
                        <option value="<%=Series(0, i)%>" selected><%=Series(1, i)%></option>
                    <%Else%>
                        <option value="<%=Series(0, i)%>"><%=Series(1, i)%></option>
                    <%End If%>
                <%Next%>
            </select>
            <input type="hidden" name="submit_series" id="submit_series" value="submit_series">
            <input class="form-control" type="submit" name="submit1" id="submit1" value="Manage This">
            </form>

            <%If Not CLng(lWhichSeries) = 0 Then%>
                <hr>
                <h4 class="h4">Manage Series</h4>

                <form role="form" class="form-horizontal" name="edit_series" method="post" action="series_mgr.asp?which_series=<%=lWhichSeries%>" onsubmit="return chkFlds();">
                <div class="form-group row">
                    <label class="col-sm-2 col-form-label" for="series_name">Series Name:</label>
                    <div class="col-sm-4">
                        <input class="form-control" type="text" name="series_name" id="series_name" value="<%=sSeriesName%>">
                    </div>
                    <label class="col-sm-2 col-form-label" for="sport">Sport:</label>
                    <div class="col-sm-4">
                        <select class="form-control" name="sport" id="sport">
                            <%If sSport = "Nordic Ski" Then%>
                                <option value="Nordic Ski" selected>Nordic Ski</option>
                                <option value="Cross-Country">Cross-Country</option>
                            <%Else%>
                                <option value="Nordic Ski">Nordic Ski</option>
                                <option value="Cross-Country" selected>Cross-Country</option>
                            <%End If%>
                        </select>
                    </div>
                </div>
                <div class="form-group row">
                    <label class="col-sm-2 col-form-label" for="year">Year Beginning:</label>
                    <div class="col-sm-4">
                        <select class="form-control" name="year" id="year">
                            <%For i = 2010 To Year(Date)%>
                                <%If CInt(iYear) = CInt(i) Then%>
                                    <option value="<%=i%>" selected><%=i%></option>
                                <%Else%>
                                    <option value="<%=i%>"><%=i%></option>
                                <%End If%>
                            <%Next%>
                        </select>
                    </div>
                    <label class="col-sm-2 col-form-label" for="max_pts">Max Pts:</label>
                    <div class="col-sm-4">
                        <input class="form-control" type="text" name="max_pts" id="max_pts" size="3" maxlength="4" value="<%=iMaxPts%>">
                    </div>
                 </div>
                <div class="form-group row">
                   <label class="col-sm-2 col-form-label" for="rank_by">Rank By:</label>
                    <div class="col-sm-4">
                        <select class="form-control" name="rank_by" id="rank_by">
                            <%If sRankBy = "points" Then%>
                                <option value="points" selected>points</option>
                                <option value="pctle">pctle</option>
                            <%Else%>
                                <option value="points">points</option>
                                <option value="pctle" selected>pctle</option>
                            <%End If%>
                        </select>
                    </div>
                    <label class="col-sm-2 col-form-label" for="comments">Comments:</label>
                    <div class="col-sm-4">
                        <textarea class="form-control" name="comments" id="comments" rows="3"><%=sComments%></textarea>
                    </div>
                </div>
                <div class="form-group row">
                    <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                    <input class="form-control" type="submit" name="submit2" id="submit2" value="Save Changes">
                </div>
                </form>

                <hr>

                <div class="row">
                    <div class="col-sm-6">
                        <h4 class="h4">Series Meets:</h4>

                        <form role="form" class="form" name="remove_meet" method="post" action="series_mgr.asp?which_series=<%=lWhichSeries%>">
                        <table class="table table-striped">
                            <tr>
                                <th>No.</th>
                                <th>Meet</th>
                                <th>Date</th>
                                <th>Remove</th>
                            </tr>
                            <%For i = 0 To UBound(SeriesMeets, 2) - 1%>
                                <tr>
                                    <td><%=i + 1%></td>
                                    <td><%=SeriesMeets(1, i)%></td>
                                    <td><%=SeriesMeets(2, i)%></td>
                                    <td style="text-align: center;">
                                        <input type="checkbox" name="series_meet_<%=SeriesMeets(0, i)%>" id="series_meet_<%=SeriesMeets(0, i)%>">
                                    </td>
                                </tr>
                            <%Next%>
                            <tr>
                                <td style="text-align: center;" colspan="4">
                                    <input type="hidden" name="submit_remove" id="submit_remove" value="submit_remove">
                                    <input class="form-control" type="submit" name="submit3" id="submit3" value="Remove Selected">
                                </td>
                            </tr>
                        </table>
                        </form>

                        <h4 class="h4">Series Races</h4>

                        <form role="form" class="form" name="manage_races" method="post" action="series_mgr.asp?which_series=<%=lWhichSeries%>">
                        <table class="table table-striped">
                            <%For i = 0 To UBound(SeriesMeets, 2) - 1%>
                                <tr>
                                    <th style="padding-top: 10px;" valign="top"><%=SeriesMeets(1, i)%></th>
                                    <td style="padding-top: 10px;">
                                        <%Call GetMeetRaces(SeriesMeets(0, i), SeriesMeets(3, i))%>
                                        <ul>
                                            <%For j = 0 To UBound(MeetRaces, 2) - 1%>
                                                <%If MeetRaces(2, j) = "x" Then%>
                                                    <li><input type="checkbox" name="race_<%=MeetRaces(0, j)%>" id="race_<%=MeetRaces(0, j)%>"
                                                            checked>&nbsp;<%=MeetRaces(1, j)%></li>
                                                <%Else%>
                                                    <li><input type="checkbox" name="race_<%=MeetRaces(0, j)%>" id="race_<%=MeetRaces(0, j)%>">
                                                    <%=MeetRaces(1, j)%></li>
                                                <%End If%>
                                            <%Next%>
                                        </ul>
                                    </td>
                                </tr>
                            <%Next%>
                            <tr>
                                <td style="text-align: center;" colspan="4">
                                    <input type="hidden" name="submit_races" id="submit_races" value="submit_races">
                                    <input class="form-control" type="submit" name="submit5" id="submit5" value="Save Changes">
                                </td>
                            </tr>
                        </table>
                        </form>
                    </div>
                    <div class="col-sm-6">
                        <h4 class="h4">Remaining Meets:</h4>

                        <form role="form" class="form" name="add_meet" method="post" action="series_mgr.asp?which_series=<%=lWhichSeries%>">
                        <table class="table table-striped">
                            <tr>
                                <th>No.</th>
                                <th>Meet</th>
                                <th>Date</th>
                                <th>Add</th>
                            </tr>
                            <%For i = 0 To UBound(Meets, 2) - 1%>
                                <tr>
                                    <td><%=i + 1%></td>
                                    <td><%=Meets(1, i)%></td>
                                    <td><%=Meets(2, i)%></td>
                                    <td style="text-align: center;"><input type="checkbox" name="meet_<%=Meets(0, i)%>" id="meet_<%=Meets(0, i)%>"></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <td style="text-align: center;" colspan="4">
                                    <input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
                                    <input class="form-control" type="submit" name="submit4" id="submit4" value="Add Selected">
                                </td>
                            </tr>
                        </table>
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
