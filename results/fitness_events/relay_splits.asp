<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID
Dim sEventName, sRaceName
Dim dEventDate
Dim RelayTeams(), TeamMembers()

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

sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
Set rs = Nothing

Private Sub GetResults(sGender)
    Dim x
            
    x = 0
    ReDim RelayTeams(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RelayTeamsID, TeamName, AgeGrp, FnlTime FROM RelayTeams WHERE RaceID = " &lRaceID & " AND Gender = '" & sGender 
    sql = sql & "' AND FnlTime > '00:00:00.000' ORDER BY Place, AgeGrp, TeamName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        RelayTeams(0, x) = rs(0).Value
        RelayTeams(1, x) = Replace(rs(1).Value, "''", "'")
        RelayTeams(2, x) = rs(2).Value
        RelayTeams(3, x) = rs(3).Value
        x = x + 1
        ReDim Preserve RelayTeams(3, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTmMmbrs(lRelayTeamID) 
    Dim x

    x = 0
    ReDim TeamMembers(5, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Bib, FirstName, LastName, Gender, Age, Leg, Split FROM RelayTmMmbrs WHERE RelayTeamsID = " & lRelayTeamID & " ORDER BY Leg"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        TeamMembers(0, x) = rs(0).Value
        TeamMembers(1, x) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
        TeamMembers(2, x) = rs(3).Value
        TeamMembers(3, x) = rs(4).Value
        TeamMembers(4, x) = rs(5).Value
        TeamMembers(5, x) = rs(6).Value
        x = x + 1
        ReDim Preserve TeamMembers(5, x)
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
<title>GSE Relay Results with Splits</title>
<meta name="description" content="Gopher State Events Relay w/Results With Splits.">
 <!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Relay Results">

    <div class="bg-warning">
        <a href="javascript:window.print();">Print</a>
        &nbsp;|&nbsp;
        <a href="relay_by_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">Individual Finish By Leg</a>
    </div>
	<h1 class="h1">Gopher State Events Relay Results with Splits</h1>
    <h2 class="h2"><%=sEventName%>&nbsp;On&nbsp;<%=dEventDate%></h2>
    <h3 class="h3"><%=sRaceName%></h3>

    <%For k = 0 To 2%>
        <div class="row">
            <%Select Case k%>
                <%Case "0"%>
                    <h4 class="h4">Combined</h4>
                    <%Call GetResults("Combined")%>
                <%Case "1"%>
                    <h4 class="h4">Male</h4>
                    <%Call GetResults("Male")%>
                <%Case "2"%>
                    <h4 class="h4">Female</h4>
                    <%Call GetResults("Female")%>
            <%End Select%>
        </div>
        <div class="row">
            <%For i = 0 To UBound(RelayTeams, 2) - 1%>
                <table class="table table-condensed">
                    <tr>
                        <th><%=i + 1%>)</th>
                        <th>TEAM NAME</th><td><%=RelayTeams(1, i)%></td>
                        <th>AGE GRP</th><td><%=RelayTeams(2, i)%></td>
                        <th>TIME</th><td><%=RelayTeams(3, i)%></td>
                    </tr>
                    <tr>
                        <td colspan="7">
                            <%Call GetTmMmbrs(RelayTeams(0, i))%>
                            <table class="table-striped" style="margin-left: 25px;">
                                <tr><th>BIB</th><th>NAME</th><th>M/F</th><th>AGE</th><th>LEG</th><th>SPLIT</th></tr>
                                <%For j = 0 To UBound(TeamMembers, 2) - 1%>
                                    <tr>
                                        <td><%=TeamMembers(0, j)%></td>
                                        <td><%=TeamMembers(1, j)%></td>
                                        <td><%=TeamMembers(2, j)%></td>
                                        <td><%=TeamMembers(3, j)%></td>
                                        <td><%=TeamMembers(4, j)%></td>
                                        <td><%=TeamMembers(5, j)%></td>
                                    </tr>
                                <%Next%>
                            </table>
                        </td>
                    </tr>
                </table>
            </div>
        <%Next%>
    <%Next%>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>