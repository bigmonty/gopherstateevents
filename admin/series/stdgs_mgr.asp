<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k
Dim lSeriesID
Dim sSeriesName, sUpdateThese
Dim Series(), SeriesRaces(), SeriesParts, SeriesStdgs()
Dim dLastRsltsUpdate

Server.ScriptTimeout = 1200

lSeriesID = Request.QueryString("series_id")

sUpdateThese = Request.QueryString("update_these")
If sUpdateThese = vbNullString Then sUpdateThese = "n"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.form.Item("submit_series") = "submit_series" Then
    lSeriesID = Request.Form.Item("series")
End If

If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID, SeriesName, SeriesYear FROM Series ORDER BY SeriesYear DESC, SeriesName"
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
    sql = "SELECT SeriesName, LastRsltsUpdate FROM Series WHERE SeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    dLastRsltsUpdate = rs(1).Value
    rs.Close
    Set rs = Nothing

    'get series races
    j = 0
    ReDim SeriesRaces(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT sr.RaceID, sr.RaceName, se.EventName FROM SeriesRaces sr INNER JOIN SeriesEvents se ON sr.SeriesEventsID = se.SeriesEventsID "
    sql = sql & "WHERE se.SeriesID = " & lSeriesID & " ORDER BY se.EventDate"
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
    sql = "SELECT ParticipantID, Age, Gender, SeriesPartsID FROM SeriesParts WHERE SeriesID = " & lSeriesID & " ORDER BY PartName"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        SeriesParts = rs.GetRows()
    Else
        ReDim SeriesParts(3, 0)
    End If
    rs.Close
    Set rs = Nothing
  
    If sUpdateThese = "y" Then
        'delete all series results
        sql = "DELETE FROM SeriesResults WHERE SeriesID = " & lSeriesID
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        For i = 0 To UBound(SeriesRaces, 2) - 1
            For j = 0 To UBound(SeriesParts, 2)
                Call MyPts(SeriesRaces(0, i), SeriesParts(0, j), SeriesParts(1, j), SeriesParts(2, j), SeriesParts(3, j))
            Next
        Next

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT LastRsltsUpdate FROM Series WHERE SeriesID = " & lSeriesID
        rs.Open sql, conn, 1, 2
        rs(0).Value =Now()
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
End If

Private Function GetRaceDist(lEventID)
    GetRaceDist = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Dist FROM SeriesRaces WHERE SeriesEventsID = " & lEventID
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        GetRaceDist = GetRaceDist & Replace(rs2(0).Value, "_", " ") & ", "
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If Not GetRaceDist = vbNullString Then
        GetRaceDist = Trim(GetRaceDist)
        GetRaceDist = Left(GetRaceDist, Len(GetRaceDist) - 1)
    End If
End Function

Private Sub MyPts(lThisRaceID, lMyID, iAge, sMF, lSeriesPartsID)
    Dim iMyEvntPl, iNumFin, iMyRank, iAgeTo, iAgeFrom
    Dim sngMyPts, sngMyTime
    Dim bInsrtRcd
    Dim x
    Dim bFound

    iMyEvntPl = 0
    iNumFin = 0
    sngMyPts = 0

    'get num gender finishers
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.ParticipantID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " 
    sql = sql & lThisRaceID & " AND p.Gender = '" & sMF & "' AND ir.FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iNumFin = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get my event place
    x = 0
    bFound = False
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, FnlScnds FROM IndResults WHERE RaceID = " & lThisRaceID & " AND FnlScnds > 0 ORDER BY FnlScnds"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        x = x + 1
        If CLng(rs(0).Value) = Clng(lMyID) Then 
            sngMyTime = rs(1).Value
            bFound = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If bFound = True Then
        iMyEvntPl = x
    Else
        Exit Sub
    End If

    'get num open finishers for this gender
    If CInt(iNumFin) > 0 Then
        iMyRank = "0"

        'get num gender finishers
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ir.ParticipantID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " & lThisRaceID 
        sql = sql & " AND p.Gender = '" & sMF & "' AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            iMyRank = CInt(iMyRank) + 1
            If CLng(rs(0).Value) = CLng(lMyID) Then Exit Do
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        'get my gender points for this race
        If CInt(iNumFin) > 0 Then 
            If CInt(iMyRank) > 0 Then sngMyPts = Round(((CInt(iNumFin) - CInt(iMyRank) + 1)/CInt(iNumFin))*100, 2)
        End If

        'add my points to my total
        bInsrtRcd = True
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT GndrPts, FnlScnds FROM SeriesStdgs WHERE SeriesPartsID = " & lSeriesPartsID & " AND RaceID = " & lThisRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            rs(0).Value = CSng(sngMyPts)
            rs(1).Value = sngMyTime
            rs.Update
            bInsrtRcd = False
        End If
        rs.Close
        Set rs = Nothing

        If bInsrtRcd = True Then
            sql = "INSERT INTO SeriesStdgs (SeriesPartsID, RaceID, GndrPts, FnlScnds) VALUES (" & lSeriesPartsID & ", " & lThisRaceID & ", " & sngMyPts 
            sql = sql & ", " & sngMyTime & ")"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If

        'total their points for this series by gender
        bInsrtRcd = False
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT GndrScore FROM SeriesResults WHERE SeriesID = " & lSeriesID & " AND SeriesPartsID = " & lSeriesPartsID
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
            sql = "INSERT INTO SeriesResults (SeriesPartsID, SeriesID, GndrScore) VALUES (" & lSeriesPartsID & ", " & lSeriesID & ", " & sngMyPts & ")"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If

        'get age from and age to
        If CInt(iAge) <= 14 Then
            iAgeFrom = 0
            iAgeTo = 14
        ElseIf CInt(iAge) >= 70 Then
            iAgeFrom = iAge
            iAgeTo = 98
        Else
            For x = 15 To 65 Step 5
                If CInt(iAge) >= CInt(x) Then
                    iAgeFrom = x
                    iAgeTo = CInt(iAgeFrom) + 4
                End If
            Next
        End If

        If CInt(iAge) = 99 Then
            sngMyPts = "0"
        Else
            'get num age group finishers for this gender
'            iMyRank = 0
 '           Set rs = Server.CreateObject("ADODB.Recordset")
  '          sql = "SELECT ir.ParticipantID FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID "
   '         sql = sql & "INNER JOIN PartRace pr ON pr.ParticipantID = p.ParticipantID WHERE pr.RaceID = " & lThisRaceID & " AND ir.RaceID = " 
    '        sql = sql & lThisRaceID & " AND p.Gender = '" & sMF & "' AND pr.Age >= " & iAgeFrom & " AND pr.Age <= " & iAgeTo 
     '       sql = sql & " AND FnlScnds > 0 ORDER BY ir.FnlScnds"
      '      rs.Open sql, conn, 1, 2
       '     If rs.RecordCount > 0 Then iNumFin = rs.RecordCount
        '    Do While Not rs.EOF
         '       iMyRank = CINt(iMyRank) + 1
          '      If CLng(rs(0).Value) = CLng(lMyID) Then Exit Do
           '     rs.MoveNext
'            Loop
 '           rs.Close
  '          Set rs = Nothing

            iMyRank = 0
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT sp.ParticipantID FROM SeriesParts sp INNER JOIN SeriesStdgs st ON sp.SeriesPartsID = st.SeriesPartsID "
            sql = sql & "WHERE st.RaceID = " & lThisRaceID & " AND sp.Gender = '" & sMF & "' AND sp.Age >= " & iAgeFrom & " AND sp.Age <= " & iAgeTo 
            sql = sql & " AND st.FnlScnds > 0 ORDER BY st.FnlScnds"
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then iNumFin = rs.RecordCount
            Do While Not rs.EOF
                iMyRank = CInt(iMyRank) + 1
                If CLng(rs(0).Value) = CLng(lMyID) Then Exit Do
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            'assign my points
            sngMyPts = 0
            If CInt(iNumFin) > 0 Then sngMyPts = Round(((CInt(iNumFin) - CInt(iMyRank) + 1)/CInt(iNumFin))*100, 2)
        End If

        'add my points to my total
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AgePts FROM SeriesStdgs WHERE SeriesPartsID = " & lSeriesPartsID & " AND RaceID = " & lThisRaceID
        rs.Open sql, conn, 1, 2
        rs(0).Value = CSng(sngMyPts)
        rs.Update
        rs.Close
        Set rs = Nothing

        'total their points for this series by age
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT AgeScore FROM SeriesResults WHERE SeriesID = " & lSeriesID & " AND SeriesPartsID = " & lSeriesPartsID
        rs.Open sql, conn, 1, 2
        rs(0).Value = CSng(rs(0).Value) + CSng(sngMyPts)
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE (Gopher State Events) Series Results</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <!--#include file = "series_nav.asp" -->

			<h4 class="h4">GSE Series Results</h4>

            <%If UBound(Series, 2) > 0 Then%>
   			    <form name="select_series" method="Post" action="stdgs_mgr.asp">
                <span style="font-weight: bold;">Select Series:</span>
                <select name="series" id="series" onchange="this.form.submit1.click();">
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
			    <input type="submit" name="submit1" id="submit1" value="Select Series To View">
			    </form>
            <%End If%>

            <%If Not CLng(lSeriesID) = 0 Then%>
                <div style="text-align: right;margin: 0;padding: 0;">
                    <a href="stdgs_mgr.asp?update_these=y&amp;series_id=<%=lSeriesID%>" style="font-size: 0.85em;">Update Results</a>
                </div>
            <%End If%>
		</div>
	</div>
<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>