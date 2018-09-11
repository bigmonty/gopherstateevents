<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, sql2, rs2
Dim i, j, k
Dim lRaceID, lEventID
Dim sEventName, sGender, sRaceName, sPartName, sMF
Dim dEventDate
Dim SortArr(8), Races

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get races
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Races = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
End If

'get num legs
Dim iNumLegs
iNumLegs = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT NumLegs FROM MultiSettingsChip WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iNumLegs = rs(0).Value
rs.Close
Set rs = Nothing

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

'get trans data
i = 0
ReDim TransData(8, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Bib, Trans1In, Trans1Out, Trans1Time, Trans2In, Trans2Out, Trans2Time FROM TransData WHERE RaceID = " & lRaceID & " ORDER BY Bib"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Call PartData(rs(0).Value)

    TransData(0, i) = rs(0).Value
    TransData(1, i) = rs(1).Value
    TransData(2, i) = rs(2).Value
    TransData(3, i) = rs(3).Value
    TransData(4, i) = rs(4).Value
    TransData(5, i) = rs(5).Value
    TransData(6, i) = rs(6).Value
    TransData(7, i) = sPartName
    TransData(8, i) = sMF
    i = i + 1
    ReDim Preserve TransData(8, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub PartData(iThisBib)
    sPartName = vbNullString
    sMF = vbNullString
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT p.FirstName, p.LastName, p.Gender FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID WHERE pr.RaceID = "
    sql2 = sql2 & lRaceID & " AND pr.Bib = " & iThisBib
    rs2.Open sql2, conn, 1, 2
    sPartName = Replace(rs2(0).Value, "''", "'") & " " & Replace(rs2(1).Value, "''", "'")
    sMF = rs2(2).Value
    rs2.Close
    Set rs2 = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Multi-Sport Transition Data</title>
<meta name="description" content="Gopher State Events transition data.">
 <!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Transition Results">

    <div class="bg-warning">
        <a href="javascript:window.print();">Print Page</a>
        &nbsp;|&nbsp;
        <a href="multi_splits.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">Results with Splits</a>
    </div>

	<h1 class="h1">Gopher State Events Multi-Sport Transition Data:&nbsp;<%=sEventName%>&nbsp;On&nbsp;<%=dEventDate%></h1>

    <%If UBound(Races, 2) > 0 Then%>
		<form class="form-inline" name="get_races" method="post" action="trans_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
        <div class="form-group">
		    <label for="races">Race:</label>
		    <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()" style="font-size:0.9em;">
			    <%For i = 0 to UBound(Races, 2)%>
				    <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
					    <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
				    <%Else%>
					    <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
				    <%End If%>
			    <%Next%>
		    </select>
		    <input type="hidden" class="form-control" name="submit_race" id="submit_race" value="submit_race">
		    <input type="submit" class="form-control" name="get_race" id="get_race" value="View">
        </div>
		</form>
    <%Else%>
        <h2 class="h2"><%=sRaceName%></h2>
    <%End If%>

    <div class="row">
        <%If CInt(iNumLegs) > 2 Then%>
            <table class="table table-striped">
                <tr>
                    <th>No</th>
                    <th>Bib</th>
                    <th>Name</th>
                    <th>MF</th>
                    <th>Trans 1 In</th>
                    <th>Trans 1 Out</th>
                    <th>Time</th>
                    <th>Trans 2 In</th>
                    <th>Trans 2 Out</th>
                    <th>Time</th>
                </tr>
                <%For i = 0 To UBound(TransData, 2) - 1%>
                    <tr>
                        <td><%=i + 1%></td>
                        <td><%=TransData(0, i)%></td>
                        <td><%=TransData(7, i)%></td>
                        <td><%=TransData(8, i)%></td>
                        <td><%=TransData(1, i)%></td>
                        <td><%=TransData(2, i)%></td>
                        <td><%=TransData(3, i)%></td>
                        <td><%=TransData(4, i)%></td>
                        <td><%=TransData(5, i)%></td>
                        <td><%=TransData(6, i)%></td>
                    </tr>
                <%Next%>
            </table>
        <%Else%>
            <table class="table table-striped">
                <tr>
                    <th>Bib</th>
                    <th>Name</th>
                    <th>MF</th>
                    <th>Trans 1 In</th>
                    <th>Trans 1 Out</th>
                    <th>Transition Time</th>
                </tr>
                <%For i = 0 To UBound(TransData, 2) - 1%>
                    <tr>
                        <td><%=TransData(0, i)%></td>
                        <td><%=TransData(7, i)%></td>
                        <td><%=TransData(8, i)%></td>
                        <td><%=TransData(1, i)%></td>
                        <td><%=TransData(2, i)%></td>
                        <td><%=TransData(3, i)%></td>
                    </tr>
                <%Next%>
            </table>
        <%End If%>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>