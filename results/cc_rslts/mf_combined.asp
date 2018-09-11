<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lThisMeet
Dim sMeetName, sRaces, sGradeYear
Dim sngOffset
Dim dMeetDate
Dim CombRslts, CombTms(3, 15), SortArr(3), PartSort(8)

'lThisMeet = Request.QueryString("meet_id")

lThisMeet = 346
sRaces = "1629,1630"

'lThisMeet = 294
'sRaces = "1314,1315"

sngOffSet = 185.85

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get combined teams array
CombTms(0, 0) = "Burnsville"
CombTms(1, 0) = "553, 360"
CombTms(0, 1) = "Charter Stars"
CombTms(1, 1) = "555, 556"
CombTms(0, 2) = "Fair School"
CombTms(1, 2) = "306, 307"
CombTms(0, 3) = "Northfield"
CombTms(1, 3) = "551, 552"
CombTms(0, 4) = "Richfield"
CombTms(1, 4) = "79, 80"
CombTms(0, 5) = "Rosemount"
CombTms(1, 5) = "268, 269"
CombTms(0, 6) = "Forest Lake"
CombTms(1, 6) = "554"
CombTms(0, 7) = "Minnetonka"
CombTms(1, 7) = "242, 243"
CombTms(0, 8) = "Academy of Holy Angels"
CombTms(1, 8) = "10, 11"
CombTms(0, 9) = "Eastview"
CombTms(1, 9) = "233"
CombTms(0, 10) = "Mineapolis Patrick Henry"
CombTms(1, 10) = "662, 663"
CombTms(0, 11) = "St. Agnes"
CombTms(1, 11) = "329, 359"
CombTms(0, 12) = "Eagan"
CombTms(1, 12) = "285"
CombTms(0, 13) = "Visitation"
CombTms(1, 13) = "45"
CombTms(0, 14) = "Prior Lake"
CombTms(1, 14) = "550"
CombTms(0, 15) = "Waseca"
CombTms(1, 15) = "253, 254"

CombTms(2, 0) = 0   'num parts
CombTms(3, 0) = 0   'team points
CombTms(2, 1) = 0   'num parts
CombTms(3, 1) = 0   'team points
CombTms(2, 2) = 0   'num parts
CombTms(3, 2) = 0   'team points
CombTms(2, 3) = 0   'num parts
CombTms(3, 3) = 0   'team points
CombTms(2, 4) = 0   'num parts
CombTms(3, 4) = 0   'team points
CombTms(2, 5) = 0   'num parts
CombTms(3, 5) = 0   'team points
CombTms(2, 6) = 0   'num parts
CombTms(3, 6) = 0   'team points
CombTms(2, 7) = 0   'num parts
CombTms(3, 7) = 0   'team points
CombTms(2, 8) = 0   'num parts
CombTms(3, 8) = 0   'team points
CombTms(2, 9) = 0   'num parts
CombTms(3, 9) = 0   'team points
CombTms(2, 10) = 0   'num parts
CombTms(3, 10) = 0   'team points
CombTms(2, 11) = 0   'num parts
CombTms(3, 11) = 0   'team points
CombTms(2, 12) = 0   'num parts
CombTms(3, 12) = 0   'team points
CombTms(2, 13) = 0   'num parts
CombTms(3, 13) = 0   'team points
CombTms(2, 14) = 0   'num parts
CombTms(3, 14) = 0   'team points
CombTms(2, 15) = 0   'num parts
CombTms(3, 15) = 0   'team points

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing
	
'get year for roster grades
If Month(dMeetDate) <= 7 Then
	sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
	sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

'get results
i = 0
ReDim CombRslts(8, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ir.Bib, r.FirstName, r.LastName, r.Gender, g.Grade" & sGradeYear & ", t.TeamName, ir.RaceTime, ir.FnlScnds, r.TeamsID "
sql = sql & "FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
sql = sql & "INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID IN (" & sRaces & ") AND ir.Place > 0 AND ir.RaceTime > '00:00' "
sql = sql & "AND ir.Excludes = 'n' AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds, ir.Place"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    For j = 0 To 8
        CombRslts(j, i) = rs(j).Value
    Next

    If UCase(CombRslts(3, i)) = "F" Then CombRslts(7, i) = CSng(CombRslts(7, i)) - CSng(sngOffset)

    i = i + 1
    ReDim Preserve CombRslts(8, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

For i = 0 To UBound(CombRslts, 2) - 2
    For j = i + 1 To UBound(CombRslts, 2) - 1
        If CSng(CombRslts(7, i)) > CSng(CombRslts(7, j)) Then
            For k = 0 To 8
                PartSort(k) = CombRslts(k, i)
                CombRslts(k, i) = CombRslts(k, j)
                CombRslts(k, j) = PartSort(k)
            Next
        End If
    Next
Next   

'get points into team array
k = 1
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT t.TeamName FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
sql = sql & "WHERE ir.RacesID IN (" & sRaces & ") AND ir.RaceTime > '00:00' AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
rs.Open sql, conn, 1, 2 
Do While Not rs.EOF
    For i = 0 To UBound(CombTms, 2)
        If CStr(CombTms(0, i)) = CStr(rs(0).Value) Then     'if the team id is in the CombTms array
            CombTms(2, i) = CInt(CombTms(2, i)) + 1         'increment the team counter by 1

            If CInt(CombTms(2, i)) <= 5 Then                'only count the top 5
                CombTms(3, i) = CSng(CombTms(3, i)) + k     'add the place to team score
            End If

            Exit For
        End If
    Next

    k = k + 1

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'sort team array
For i = 0 To UBound(CombTms, 2) - 1
    For j = i + 1 To UBound(CombTms, 2)
        If CSng(CombTms(3, i)) > CSng(CombTms(3, j)) Then
            For k = 0 To 3
                SortArr(k)  = CombTms(k, i)
                CombTms(k, i) = CombTms(k, j)
                CombTms(k, j) = SortArr(k)
            Next
        End If
    Next
Next
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CC/Nordic Results Manager: Boys-Girls Combined Results</title>
<<meta name="description" content="Cross-Country & Nordic Ski Results by Gopher State Events, a conventional timing service offererd by H51 Software, LLC in Minnetonka, MN.">
<!--#include file = "../../includes/js.asp" --> 
</head>
<body>
<div class="container">
    <img src="/graphics/mobile_results.png" alt="Boys/Girls Combined Results">
	<div class="bg-info"><a href="javascript:window.print();">Print</a></div>

	<h4 class="h4">Boys-Girls Varsity Combined Results for <%=sMeetName%> on <%=dMeetDate%></h4>
					
    <h4 class="h4">Team Places</h4>  
	<table class="table table-striped">
        <tr><th>Pl</th><th>School</th><th>Score</th></tr>
        <%k = 1%>
        <%For i = 0 To UBound(CombTms, 2)%>
            <%If CInt(CombTms(2, i)) >= 5 Then%>
                <tr>
                    <td><%=k%></td>
                    <td><%=CombTms(0, i)%></td>
                    <td><%=CombTms(3, i)%></td>
                </tr>

                <%k = k + 1%>
            <%End If%>
        <%Next%>
    </table>
				
    <h4 class="h4">Individual Places</h4>            	
	<table class="table table-striped">
        <tr><th>Pl</th><th>Bib</th><th>First</th><th>Last</th><th>MF</th><th>Gr</th><th>School</th><th>Time</th></tr>
        <%For i = 0 To UBound(CombRslts, 2) - 1%>
            <tr>
                <td><%=i + 1%></td>
                <%For j = 0 To 6%>
                    <td><%=CombRslts(j, i)%></td>
                <%Next%>
            </tr>
        <%Next%>
    </table>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
