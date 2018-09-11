<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k
Dim lSeriesID
Dim sSeriesName, sUpdateThese, sRankBy
Dim iMaxPts
Dim Series(), SeriesRaces(), SeriesParts, SeriesStdgs()
Dim dLastRsltsUpdate

Server.ScriptTimeout = 1200

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"

sUpdateThese = Request.QueryString("update_these")
If sUpdateThese = vbNullString Then sUpdateThese = "n"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.form.Item("submit_series") = "submit_series" Then
    lSeriesID = Request.Form.Item("series")
End If

If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CCSeriesID, SeriesName, SeriesYear FROM CCSeries ORDER BY SeriesYear DESC, SeriesName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Series(0, i) = rs(0).Value
	Series(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve Series(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not CLng(lSeriesID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SeriesName, LastRsltsUpdate, RankBy, MaxPts FROM CCSeries WHERE CCSeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    dLastRsltsUpdate = rs(1).Value
    sRankBy = rs(2).Value
    iMaxPts = rs(3).Value
    rs.Close
    Set rs = Nothing

    'get series races
    j = 0
    ReDim SeriesRaces(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sr.RacesID, sr.RaceName, sm.MeetName FROM CCSeriesRaces sr INNER JOIN CCSeriesMeets sm ON sr.CCSeriesMeetsID = sm.CCSeriesMeetsID "
    sql = sql & "WHERE sm.CCSeriesID = " & lSeriesID & " ORDER BY sm.MeetDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SeriesRaces(0, j) = rs(0).Value
        SeriesRaces(1, j) = Replace(rs(1).Value, "''", "'") & " - " & Replace(rs(2).Value, "''", "'")
        j = j + 1
        ReDim Preserve SeriesRaces(1, j)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID, CCSeriesPartsID FROM CCSeriesParts WHERE CCSeriesID = " & lSeriesID & " ORDER BY PartName"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        SeriesParts = rs.GetRows()
    Else
        ReDim SeriesParts(1, 0)
    End If
    rs.Close
    Set rs = Nothing
  
    If sUpdateThese = "y" Then
        'delete all series results
        sql = "DELETE FROM CCSeriesResults WHERE CCSeriesID = " & lSeriesID
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        For i = 0 To UBound(SeriesRaces, 2) - 1
            For j = 0 To UBound(SeriesParts, 2)
                Call MyPts(SeriesRaces(0, i), SeriesParts(0, j), SeriesParts(1, j))
            Next
        Next

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT LastRsltsUpdate FROM CCSeries WHERE CCSeriesID = " & lSeriesID
        rs.Open sql, conn, 1, 2
        rs(0).Value =Now()
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
End If

Private Sub MyPts(lThisRaceID, lMyID, lSeriesPartsID)
    Dim iMyEvntPl, iNumFin, iMyRank
    Dim sStartType
    Dim sngMyPts
    Dim bInsrtRcd
    Dim x

    iMyEvntPl = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT StartType FROM Races WHERE RacesID = " & lThisRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sStartType = rs(0).Value
    rs.Close
    Set rs = Nothing

    'get my event place
    x = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    IF sStartType = "Pursuit" Then
        sql = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRaceID & " AND FnlScnds > 0 AND Excludes = 'n' ORDER BY Place"
    Else
        sql = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRaceID & " AND FnlScnds > 0 AND Excludes = 'n' ORDER BY FnlScnds"
    End If
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        x = x + 1
        If CLng(lMyID) = CLng(rs(0).Value) Then
            iMyEvntPl = x
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'score based on points or pctle
    If sRankBy = "points" Then
        If CLng(iMyEvntPl) = 0 Then
            sngMyPts = 0
        Else
            If iMyEvntPl > CInt(iMaxPts) Then
                sngMyPts = 0
            Else
                sngMyPts = CInt(iMaxPts) - CInt(iMyEvntPl) + 1
            End If
        End If
    Else
        'get num open finishers
        If Not iMyEvntPl = 0 Then
            iNumFin = "0"
            iMyRank = "0"

            'get num gender finishers
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT ir.RosterID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.RacesID = " & lThisRaceID 
            sql = sql & " ORDER BY ir.Place"
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then iNumFin = rs.RecordCount
            Do While Not rs.EOF
                iMyRank = CINt(iMyRank) + 1
                If CLng(rs(0).Value) = CLng(lMyID) Then Exit Do
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            'get my points for this race
            sngMyPts = Round(((CInt(iNumFin) - CInt(iMyRank) + 1)/CInt(iNumFin))*100, 2)
        End If
    End If

    'add my points to my total
    bInsrtRcd = True
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Points FROM CCSeriesStdgs WHERE CCSeriesPartsID = " & lSeriesPartsID & " AND RacesID = " & lThisRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs(0).Value = CSng(sngMyPts)
        rs.Update
        bInsrtRcd = False
    End If
    rs.Close
    Set rs = Nothing

    If bInsrtRcd = True Then
        sql = "INSERT INTO CCSeriesStdgs (CCSeriesPartsID, RacesID, Points) VALUES (" & lSeriesPartsID & ", " & lThisRaceID & ", " & sngMyPts & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If

    'total their points for this series
    bInsrtRcd = False
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Score FROM CCSeriesResults WHERE CCSeriesID = " & lSeriesID & " AND CCSeriesPartsID = " & lSeriesPartsID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs(0).Value = CSng(rs(0).Value) + CSng(sngMyPts)
        rs.Update
    Else
        bInsrtRcd = True
    End If
    rs.Close
    Set rs = Nothing

    If bInsrtRcd = True Then
        sql = "INSERT INTO CCSeriesResults (CCSeriesPartsID, CCSeriesID, Score) VALUES (" & lSeriesPartsID & ", " & lSeriesID & ", " & sngMyPts & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CC/Nordic Standings Manager</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h2 class="h2">CC/Nordic Series Standings Manager</h2>

            <!--#include file = "cc_series_nav.asp" -->

            <%If UBound(Series, 2) > 0 Then%>
   			    <form role="form" class="form-inline" name="select_series" method="Post" action="stdgs_mgr.asp">
                <label for="series">Select Series:</label>
                <select class="form-control" name="series" id="series" onchange="this.form.submit1.click();">
                    <option value="">&nbsp;</option>
                    <%For i = 0 To UBound(Series, 2) - 1%>
                        <%If CLng(lSeriesID) = CLng(Series(0, i)) Then%>
                            <option value="<%=Series(0, i)%>" selected><%=Series(1, i)%></option>
                        <%Else%>
                            <option value="<%=Series(0, i)%>"><%=Series(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
			    <input type="hidden" name="submit_series" id="submit_series" value="submit_series">
			    <input class="form-control" type="submit" name="submit1" id="submit1" value="Select Series To View">
			    </form>
            <%End If%>

            <%If Not CLng(lSeriesID) = 0 Then%>
                <div class="bg-info">
                    <a href="stdgs_mgr.asp?update_these=y&amp;series_id=<%=lSeriesID%>" style="color:#fff;">Update Results</a>
                </div>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>