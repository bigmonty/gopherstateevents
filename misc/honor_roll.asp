<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k, m, n, p
Dim sDist
Dim HonorRollM(), HonorRollF(), HonorRollRaces(), Distances, TempArr(3), Events()

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
			
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_dist") = "submit_dist" Then sDist = Request.Form.Item("distances")

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AgeGrDistID, Distance FROM AgeGrDist ORDER BY Distance"
rs.Open sql, conn, 1, 2
Distances = rs.GetRows()
rs.Close
Set rs = Nothing

If Not sDist = vbNullString Then
    i = 0
    ReDim HonorRollRaces(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT rd.RaceID, e.EventName, e.EventDate FROM Events e INNER JOIN RaceData rd ON e.EventID = rd.EventID WHERE rd.Dist = '" 
    sql = sql & sDist & "' AND e.EventType = 5 AND e.EventDate <= '" & Date & "' AND rd.Certified = 'y'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        HonorRollRaces(0, i) = rs(0).Value
        HonorRollRaces(1, i) = Replace(rs(1).Value, "''", "'")
        HonorRollRaces(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve HonorRollRaces(2, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    j = 0
    k = 0
    ReDim HonorRollM(3, 0)
    ReDim HonorRollF(3, 0)

    For i = 0 To UBound(HonorRollRaces, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT p.ParticipantID, ir.FnlTime FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " 
        sql = sql & HonorRollRaces(0, i) & " AND p.Gender = 'm' ORDER BY EventPl"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If UBound(HonorRollM, 2) < 10 Then
                HonorRollM(0, j) = PartName(rs(0).Value)                'participant
                HonorRollM(1, j) = HonorRollRaces(1, i)                 'event name
                HonorRollM(2, j) = HonorRollRaces(2, i)                 'event date
                HonorRollM(3, j) = ConvertToSeconds(rs(1).Value)        'time
                j = j + 1
                ReDim Preserve HonorRollM(3, j)
            Else
                If CSng(ConvertToSeconds(rs(1).Value)) < CSng(HonorRollM(3, 9)) Then    'if it is better than at least one time on the list
                    'replace 10th with this one
                    HonorRollM(0, 9) = PartName(rs(0).Value)                'participant
                    HonorRollM(1, 9) = HonorRollRaces(1, i)                 'event name
                    HonorRollM(2, 9) = HonorRollRaces(2, i)                 'event name
                    HonorRollM(3, 9) = ConvertToSeconds(rs(1).Value)        'time

                    'resort the list
                    For m = 0 To 8
                        For n = m + 1 To 9
                            If CSng(HonorRollM(3, m)) > CSng(HonorRollM(3, n)) Then
                                For p = 0 To 3
                                    TempArr(p) = HonorRollM(p, m)
                                    HonorRollM(p, m) = HonorRollM(p, n)
                                    HonorRollM(p, n) = TempArr(p)
                                Next
                            End If
                        Next
                    Next
                Else
                    Exit Do
                End If
            End If

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT p.ParticipantID, ir.FnlTime FROM IndResults ir INNER JOIN Participant p ON ir.ParticipantID = p.ParticipantID WHERE ir.RaceID = " 
        sql = sql & HonorRollRaces(0, i) & " AND p.Gender = 'f' ORDER BY EventPl"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If UBound(HonorRollF, 2) < 10 Then
                HonorRollF(0, k) = PartName(rs(0).Value)                'participant
                HonorRollF(1, k) = HonorRollRaces(1, i)                 'event name
                HonorRollF(2, k) = HonorRollRaces(2, i)                 'event name
                HonorRollF(3, k) = ConvertToSeconds(rs(1).Value)        'time
                k = k + 1
                ReDim Preserve HonorRollF(3, k)
            Else
                If CSng(ConvertToSeconds(rs(1).Value)) < CSng(HonorRollF(3, 9)) Then    'if it is better than at least one time on the list
                    'replace 10th with this one
                    HonorRollF(0, 9) = PartName(rs(0).Value)                'participant
                    HonorRollF(1, 9) = HonorRollRaces(1, i)                 'event name
                    HonorRollF(2, 9) = HonorRollRaces(2, i)                 'event name
                    HonorRollF(3, 9) = ConvertToSeconds(rs(1).Value)        'time

                    'resort the list
                    For m = 0 To 8
                        For n = m + 1 To 9
                            If CSng(HonorRollF(3, m)) > CSng(HonorRollF(3, n)) Then
                                For p = 0 To 3
                                    TempArr(p) = HonorRollF(p, m)
                                    HonorRollF(p, m) = HonorRollF(p, n)
                                    HonorRollF(p, n) = TempArr(p)
                                Next
                            End If
                        Next
                    Next
                Else
                    Exit Do
                End If
            End If

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Next
End If

Private Function PartName(lPartID)
    PartName = "unknown"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID 
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then PartName = Replace(rs2(0).Value, "''", "'") & " " &  Replace(rs2(1).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

%>
<!--#include file = "../includes/convert_to_seconds.asp" -->
<!--#include file = "../includes/convert_to_minutes.asp" -->

<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; Honor Roll</title>
<meta name="description" content="GSE (Gopher State Events) event timing honor roll.">
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<h1 class="h1">Gopher State Events Honor Roll</h1>

    <div>
        Welcome to the Gopher State Events Honor Roll, listing the top performances for road events in races we have timed on 
        <span style="font-weight: bold;">CERTIFIED COURSES ONLY</span>.  If you notice
        any missing data, please <a href="mailto:bob.schneider@gopherstateevents.com" style="font-weight: bold;">contact us</a> and we will look 
        into it.  Look forward to age group-based break-down and cross-country running/nordic ski breakdowns coming soon.
    </div>

    <br>

    <form role="form" class="form-inline" name="get_dist" method="post" action="honor_roll.asp">
    <label for="distances">Distance:</label>	
    <select class="form-control" name="distances" id="distances" onchange="this.form.submit2.click();">
        <option value="">&nbsp;</option>
		<%For i = 0 to UBound(Distances, 2) - 1%>
            <%If CStr(sDist) = CStr(Distances(1, i)) Then%>
                <option value="<%=Distances(1, i)%>" selected><%=Distances(1, i)%></option>
            <%Else%>
				<option value="<%=Distances(1, i)%>"><%=Distances(1, i)%></option>
            <%End If%>
		<%Next%>
	</select>
	<input type="hidden" name="submit_dist" id="submit_dist" value="submit_dist">
	<input type="submit" class="form-control" name="submit2" id="submit2" value="Get This" style="font-size: 0.85em;">
    </form>

    <div class="row">
        <div class="col-sm-6">
            <%If Not sDist = vbNullString Then%>
                <h4 class="h4">GSE Men's <%=Replace(sDist, "_", " ")%> Honor Roll</h4>

                <table class="table table-striped">
                    <tr>
                        <th>No.</th>
                        <th>Competitor</th>
                        <th>Event</th>
                        <th>Date</th>
                        <th>Time</th>
                    </tr>
                    <%For i = 0 To UBound(HonorRollM, 2) - 1%>
                        <tr>
                            <td><%=i + 1%></td>
                            <td><%=HonorRollM(0, i)%></td>
                            <td><%=HonorRollM(1, i)%></td>
                            <td><%=HonorRollM(2, i)%></td>
                            <td><%=ConvertToMinutes(Round(HonorRollM(3, i), 1))%></td>
                        </tr>
                    <%Next%>
                </table>
            <%End If%>
        </div>
        <div class="col-sm-6">
            <%If Not sDist = vbNullString Then%>
                <h4 class="h4">GSE Women's <%=Replace(sDist, "_", " ")%> Honor Roll</h4>

                <table class="table table-striped">
                    <tr>
                        <th>No.</th>
                        <th>Competitor</th>
                        <th>Event</th>
                        <th>Date</th>
                        <th>Time</th>
                    </tr>
                    <%For i = 0 To UBound(HonorRollF, 2) - 1%>
                        <tr>
                            <td><%=i + 1%></td>
                            <td><%=HonorRollF(0, i)%></td>
                            <td><%=HonorRollF(1, i)%></td>
                            <td><%=HonorRollF(2, i)%></td>
                            <td><%=ConvertToMinutes(Round(HonorRollF(3, i), 1))%></td>
                        </tr>
                    <%Next%>
                </table>
            <%End If%>
        </div>
    </div>
</div>
<!--#include file = "../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
