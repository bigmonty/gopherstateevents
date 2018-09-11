<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID, lOCMen, lOCWomen
Dim sEventName, sGender, sDist, sRaceName, sChipStart, sErrMsg
Dim lngAgeGrDistID
Dim dEventDate
Dim Events(), Races(), IndRslts, SortArr(13)

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

sql = "SELECT Dist, RaceName, ChipStart FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sDist = rs(0).Value
sRaceName = rs(1).Value
sChipStart = rs(2).Value
Set rs = Nothing

lngAgeGrDistID = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AgeGrDistID FROM AgeGrDist WHERE Distance = '" & sDist & "'"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then lngAgeGrDistID = rs(0).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT OCMen, OCWomen FROM AgeGrOCTime WHERE AgeGrDistID = " & lngAgeGrDistID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    lOCMen = rs(0).Value
    lOCWomen = rs(1).Value
Else
    sErrMsg = "I'm sorry.  Age grading is not available for this distance."
End If
rs.Close
Set rs = Nothing

'AG Pl, Actual Pl, Bib-Name, MF, Age, AG Time, Actual Time, City, St
'(note: pr.RaceID, pr.Wave and pr. in the query is just to hold a place for the age graded time)
If sErrMsg = vbNullString Then
    sql = "SELECT p.ParticipantID, ir.EventPl, pr.Bib, p.LastName, p.FirstName, p.Gender, pr.Age, ir.FnlScnds, ir.FnlTime, p.City, p.St, pr.RaceID, pr.Wave, " 
    sql = sql & "pr.Wave FROM Participant p JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
    sql = sql & "JOIN PartRace pr ON pr.RaceID = ir.RaceID AND pr.ParticipantID = p.ParticipantID "
    sql = sql & "WHERE ir.RaceID = " & lRaceID & " AND ir.FnlTime IS NOT NULL AND ir.FnlTime > '00:00:00.000' AND pr.Age < 99 ORDER BY ir.FnlScnds"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        IndRslts = rs.GetRows()
    Else
        ReDim IndRslts(13, 0)
    End If
    rs.Close
    Set rs = Nothing

    If UBound(IndRslts, 2) > 0 Then
        If CLng(lngAgeGrDistID) > 0 Then
            'get age graded data
            For i = 0 To UBound(IndRslts, 2)
                IndRslts(0, i) = 0
                IndRslts(1, i) = i + 1                          'get actual place
                IndRslts(12, i) = "0"
                IndRslts(13, i) = "0"

                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT Factor FROM AgeGrFactors WHERE MF = '" & LCase(IndRslts(5, i)) & "' AND Age = " & IndRslts(6, i) 
                sql = sql & " AND AgeGrDistID = " & lngAgeGrDistID
                rs.Open sql, conn, 1, 2
                If rs.RecordCount > 0 Then
                    IndRslts(8, i) = CSng(IndRslts(7, i))*CSng(rs(0).Value)
                    IndRslts(11, i) = IndRslts(8, i)
                    IndRslts(12, i) = IndRslts(8, i)
                    IndRslts(8, i) = rs(0).Value
                End If
                rs.Close
                Set rs = Nothing

                If UCASE(IndRslts(5, i)) = "M" Then
                    IndRslts(13, i) = Round(CLng(lOCMen)/CSng(IndRslts(11, i)), 4)*100
                ElseIf UCASE(IndRslts(5, i)) = "F" Then
                    IndRslts(13, i) = Round(CLng(lOCWomen)/CSng(IndRslts(11, i)), 4)*100
                End If
                IndRslts(11, i) = ConvertToMinutes(IndRslts(11, i))
            Next

            'sort by age graded time
            For i = 0 To UBound(IndRslts,2)
                For j = i + 1 To UBound(IndRslts, 2) - 1
                    If CSng(IndRslts(13, i)) < CSng(IndRslts(13, j)) Then
                        For k = 0 To 13
                            SortArr(k) = IndRslts(k, i)
                            IndRslts(k, i) = IndRslts(k, j)
                            IndRslts(k, j) = SortArr(k)
                        Next
                    End If
                Next
            Next

            'get place
            For i = 0 To UBound(IndRslts, 2)
                IndRslts(0, i) = i + 1
            Next
        End If
    End If
End If

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->

<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Age-Graded Results</title>
<meta name="description" content="Gopher State Events (GSE) Age-Graded Results.">
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-sm-6">
            <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
        </div>
        <div class="col-sm-6">
            <h1 class="h1">GSE Age-Graded Results</h1>
        </div>
    </div>

    <h2 class="h2"><%=sEventName%>-<%=sRaceName%>&nbsp;On&nbsp;<%=dEventDate%></h2>

    <a href="javascript:window.print();">Print</a>

    <%If sErrMsg = vbNullString Then%>
	    <table class="table table-striped">
		    <tr>
			    <th style="border-bottom: 1px solid #ccc;" colspan="2">Place</th>
			    <th rowspan="2">Bib-Name</th>
			    <th rowspan="2">M/F</th>
  			    <th rowspan="2">Age</th>
			    <th style="border-bottom: 1px solid #ccc;" colspan="2">Time</th>
                <th rowspan="2">AG %</th>
                <th rowspan="2">AG Factor</th>
			    <th rowspan="2">From</th>
		    </tr>
		    <tr>
			    <th>Age Grd</th>
                <th>Actual</th>
			    <th>Age Grd</th>
			    <th>Actual</th>
		    </tr>

		    <%For i = 0 To UBound(IndRslts, 2)%>
				<tr>
					<td><%=IndRslts(0, i)%></td>
                    <td><%=IndRslts(1, i)%></td>
					<td><%=IndRslts(2, i)%> - <%=IndRslts(4, i)%>&nbsp;<%=IndRslts(3, i)%></td>
					<td><%=IndRslts(5, i)%></td>
					<td><%=IndRslts(6, i)%></td>
					<td><%=IndRslts(11, i)%></td>
					<td><%=ConvertToMinutes(CSng(IndRslts(7, i)))%></td>
                    <td><%=IndRslts(13, i)%>%</td>
                    <td><%=IndRslts(8, i)%></td>
                    <td><%=IndRslts(9, i)%>, <%=IndRslts(10, i)%></td>
				</tr>
		    <%Next%>
	    </table>
    <%Else%>
        <p class="bg-danger text-danger"><%=sErrMsg%></p>
    <%End If%>
	<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>