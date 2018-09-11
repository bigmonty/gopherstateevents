<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k
Dim lSeriesID
Dim sSeriesName, sUpdateParts
Dim Series(), SeriesParts, SeriesRaces, Races(), RaceParts()
Dim dFirstDate

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"

sUpdateParts = Request.QueryString("update_parts")

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
    sql = "SELECT SeriesName FROM CCSeries WHERE CCSeriesID = " & lSeriesID
    rs.Open sql, conn, 1, 2
    sSeriesName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing

    If sUpdateParts = "y" Then
        j = 0
        ReDim Races(0)
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT sr.RacesID FROM CCSeriesRaces sr INNER JOIN CCSeriesMeets se ON sr.CCSeriesMeetsID = se.CCSeriesMeetsID " & "WHERE se.CCSeriesID = " 
        sql = sql & lSeriesID
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            Races(j) = rs(0).Value
            j = j + 1
            ReDim Preserve Races(j)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        For j = 0 To UBound(Races) - 1
            k = 0
            ReDim RaceParts(3, 0)
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT r.RosterID, r.LastName, r.FirstName, r.Gender, r.TeamsID FROM Roster r INNER JOIN IndRslts ir "
            sql = sql & "ON r.RosterID = ir.RosterID WHERE ir.RacesID = " & Races(j)
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF
                If InSeries(rs(0).Value, lSeriesID) = False Then
                    If FinishedRace(rs(0).Value, Races(j)) = "y" Then
                        RaceParts(0, k) = rs(0).Value
                        RaceParts(1, k) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
                        RaceParts(2, k) = rs(3).Value
                        RaceParts(3, k) = GetSchool(rs(4).Value)
                        k = k + 1
                        ReDim Preserve RaceParts(3, k)
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing

            For k = 0 To UBound(RaceParts, 2) - 1
                sql = "INSERT INTO CCSeriesParts(CCSeriesID, RosterID, PartName, Gender, School) VALUES (" & lSeriesID & ", " & RaceParts(0, k)
                sql = sql & ", '" & Replace(RaceParts(1, k), "'", "''") & "', '" & RaceParts(2, k) & "', '" & RaceParts(3, k) & "')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            Next
        Next
    End If
 
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT CCSeriesPartsID, RosterID, PartName, Gender, School FROM CCSeriesParts WHERE CCSeriesID = " & lSeriesID & " ORDER BY PartName"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        SeriesParts = rs.GetRows()
    Else
        ReDim SeriesParts(4, 0)
    End If
    rs.Close
    Set rs = Nothing
End If

Private Function InSeries(lThisPart, lThisSeries)
    InSeries = False

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT RosterID FROM CCSeriesParts WHERE CCSeriesID = " & lThisSeries & " AND RosterID = " & lThisPart
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then InSeries = True
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function FinishedRace(lThisPart, lThisRace)
    FinishedRace = "n"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Place FROM IndRslts WHERE RacesID = " & lThisRace & " AND RosterID = " & lThisPart & " AND Place > 0  AND Excludes = 'n'"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then FinishedRace = "y"
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetSchool(lTeamID)
    GetSchool = "unknown"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetSchool = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CC/Nordic Participants Manager</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h2 class="h2">CC/Nordic Series Participant Manager</h2>

            <!--#include file = "cc_series_nav.asp" -->

            <%If UBound(Series, 2) > 0 Then%>
   			    <form role="form" class="form-inline" name="select_series" method="Post" action="part_mgr.asp">
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
                    <a href="part_mgr.asp?update_parts=y&amp;series_id=<%=lSeriesID%>" style="color:#fff;">Update Series Parts</a>
                </div>

                <h4 class="h4">Num Parts: <%=UBound(SeriesParts, 2)%></h4>

                <div class="row">
                    <div class="col-sm-6">
                        <h5 class="h5">Male Participants</h5>

                        <table class="table table-striped">
                            <tr>
                                <th>No.</th>
                                <th>Part ID</th>
                                <th>Name</th>
                                <th>M/F</th>
                                <th>School</th>
                            </tr>
                            <%j = 1%>
                            <%For i = 0 To UBound(SeriesParts, 2)%>
                                <%If SeriesParts(3, i) = "M" Then%>
                                    <tr>
                                        <td><%=j%></td>
                                        <td><%=SeriesParts(1, i)%></td>
                                        <td><%=SeriesParts(2, i)%></td>
                                        <td><%=SeriesParts(3, i)%></td>
                                        <td><%=SeriesParts(4, i)%></td>
                                    </tr>
                                    <%j = j + 1%>
                                <%End If%>
                            <%Next%>
                        </table>
                    </div>
                    <div class="col-sm-6">
                        <h5 class="h5">Female Participants</h5>

                        <table class="table table-striped">
                            <tr>
                                <th>No.</th>
                                <th>Part ID</th>
                                <th>Name</th>
                                <th>M/F</th>
                                <th>School</th>
                            </tr>
                            <%j = 1%>
                            <%For i = 0 To UBound(SeriesParts, 2)%>
                                <%If SeriesParts(3, i) = "F" Then%>
                                    <tr>
                                        <td><%=j%></td>
                                        <td><%=SeriesParts(1, i)%></td>
                                        <td><%=SeriesParts(2, i)%></td>
                                        <td><%=SeriesParts(3, i)%></td>
                                        <td><%=SeriesParts(4, i)%></td>
                                    </tr>
                                    <%j = j + 1%>
                                <%End If%>
                            <%Next%>
                        </table>
                    </div>
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