<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID, lCustomFieldsID
Dim sEventName, sRaceName, sCustomFieldName
Dim Results(), CustomFieldParts()
Dim dEventDate

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

lCustomFieldsID = Request.QueryString("custom_fields_id")
If CStr(lCustomFieldsID) = vbNullString Then lCustomFieldsID = 0
If Not IsNumeric(lCustomFieldsID) Then Response.Redirect("http://www.google.com")
If CLng(lCustomFieldsID) < 0 Then Response.Redirect("http://www.google.com")

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

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FieldName FROM CustomFields WHERE CustomFieldsID = " & lCustomFieldsID
rs.Open sql, conn, 1, 2
sCustomFieldName = rs(0).Value
rs.Close
Set rs = Nothing

i = 0
ReDim CustomFieldParts(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ParticipantID FROM CustomFieldsParts WHERE CustomFieldsID = " & lCustomFieldsID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    CustomFieldParts(i) = rs(0).Value
    i = i + 1
    ReDim Preserve CustomFieldParts(i)
    rs.MoveNext
Loop
rs.Close
Set rs=Nothing

Private Sub GetResults(sGender)
    Dim x, y
    Dim sMF

    sMF = Left(sGender, 1)

    x = 0 
    ReDim Results(3, 0)
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT pr.Bib, p.FirstName, p.LastName, pr.Age, ir.FnlScnds, p.ParticipantID FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID "
    sql = sql & "INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID WHERE pr.RaceID = " & lRaceID & " AND ir.RaceID = " & lRaceID
    sql = sql & "AND p.Gender = '" & sMF & "' AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF 
        For y = 0 To UBound(CustomFieldParts) - 1
            If CLng(rs(5).Value) = CLng(CustomFieldParts(y)) Then
                Results(0, x) = rs(0).Value
                Results(1, x) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
                Results(2, x) = rs(3).Value
                Results(3, x) = ConvertToMinutes(rs(4).Value)
                x = x + 1
                ReDim Preserve Results(3, x)
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>

<!--#include file = "../../includes/convert_to_minutes.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Relay Results with Splits</title>
<meta name="description" content="Gopher State Events Custom Field Results.">
 <!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Relay Results">

    <div class="bg-warning">
        <a href="javascript:window.print();">Print</a>
    </div>
	<h1 class="h1">Gopher State Events Results</h1>
    <h2 class="h2"><%=sEventName%>&nbsp;On&nbsp;<%=dEventDate%></h2>
    <h3 class="h3"><%=sRaceName%>: <%=sCustomFieldName%></h3>

    <div class="row">
        <%For k = 0 To 1%>
            <%Select Case k%>
                <%Case "0"%>
                    <h4 class="h4">Male</h4>
                    <%Call GetResults("Male")%>
                <%Case "1"%>
                    <h4 class="h4">Female</h4>
                    <%Call GetResults("Female")%>
            <%End Select%>

            <table class="table table-striped">
                <tr>
                    <th>Pl</th>
                    <th>Bib</th>
                    <th>Name</th>
                    <th>Age</th>
                    <th>Time</th>
                </tr>
                <%For i = 0 To UBound(Results, 2) - 1%>
                    <tr>
                        <td><%=i + 1%></td>
                        <td><%=Results(0, i)%></td>
                        <td><%=Results(1, i)%></td>
                        <td><%=Results(2, i)%></td>
                        <td><%=Results(3, i)%></td>
                    </tr>
                <%Next%>
            </table>
        <%Next%>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>