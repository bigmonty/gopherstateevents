<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID
Dim sEventName, sRaceName, sGender
Dim dEventDate
Dim Results(), SortArr(5)

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

Private Sub GetResults(sGender, iThisLeg)
    Dim x, y, z
            
    x = 0
    ReDim Results(5, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT rtm.Bib, rtm.FirstName, rtm.LastName, rtm.Age, rt.TeamName, rtm.Split "
    sql = sql & "FROM RelayTmMmbrs rtm INNER JOIN RelayTeams rt ON rtm.RelayTeamsID = rt.RelayTeamsID WHERE rt.RaceID = "
    sql = sql & lRaceID & " AND rtm.Gender = '" & sGender & "' AND rtm.Split > '00:00:00.000' AND rtm.Leg = " & iThisLeg
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Results(0, x) = rs(0).Value
        Results(1, x) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
        Results(2, x) = rs(3).Value
        Results(3, x) = rs(4).Value
        Results(4, x) = rs(5).Value
        Results(5, x) = ConvertToSeconds(rs(5).Value)
        x = x + 1
        ReDim Preserve Results(5, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'sort by time
    For x = 0 To UBound(Results, 2) - 2
        For y = x + 1 To UBound(Results, 2) - 1
            If CLng(Results(5, x)) > CLng(Results(5, y)) Then
                For z = 0 To 5
                    SortArr(z) = Results(z, x)
                    Results(z, x) = Results(z, y)
                    Results(z, y) = SortArr(z)
                Next
            End If
        Next
    Next
End Sub

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Relay Individual Finish By Leg</title>
<meta name="description" content="Gopher State Events Relay Individual Finish By Leg.">
 <!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Relay Results">

    <div class="bg-warning">
        <a href="javascript:window.print();">Print</a>
        &nbsp;|&nbsp;
        <a href="relay_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">Relay Results with Splits</a>
    </div>
	<h1 class="h1">Gopher State Events Relay Individual Finish By Leg</h1>
    <h2 class="h2"><%=sEventName%>&nbsp;On&nbsp;<%=dEventDate%></h2>
    <h3 class="h3"><%=sRaceName%></h3>

 
    <%For j = 0 To 1%>
        <%For k = 1 To 2%>
            <div class="row">
                <%Select Case j%>
                    <%Case "0"%>
                        <h4 class="h4">MALE RESULTS LEG <%=k%></h4>
                        <%Call GetResults("m", k)%>
                    <%Case "1"%>
                        <h4 class="h4">FEMALE RESULTS LEG <%=k%></h4>
                        <%Call GetResults("f", k)%>
                <%End Select%>
            </div>
            <div class="row">
                <table class="table table-striped">
                    <tr>
                        <th>PL</th>
                        <th>BIB</th>
                        <th>NAME</th>
                        <th>AGE</th>
                        <th>TEAM</th>
                        <th>SPLIT</th>
                    </tr>
                    <%For i = 0 To UBound(Results, 2) - 1%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td><%=Results(0, i)%></td>
                            <td><%=Results(1, i)%></td>
                            <td><%=Results(2, i)%></td>
                            <td><%=Results(3, i)%></td>
                            <td><%=Results(4, i)%></td>
                        </tr>
                    <%Next%>
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