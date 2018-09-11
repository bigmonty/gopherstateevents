<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lPartID
Dim sEventName, sLocation, sRaceName
Dim iRaceAge
Dim dEventDate
Dim RaceHist(), PartData(7), PartRace(), PartReg()

lPartID = Request.QueryString("part_id")
If Not IsNumeric(lPartID) Then Response.Redirect "htttp://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_trim") = "submit_trim" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
    rs.Open sql, conn, 1, 2
    rs(0).Value = Trim(rs(0).Value)
    rs(1).Value = Trim(rs(1).Value)
    rs.Update
    rs.Close
    Set rs = Nothing

    'refresh parent
    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
ElseIf Request.Form.Item("submit_delete") = "submit_delete" Then
    sql = "DELETE FROM IndResults WHERE ParticipantID = " & lPartID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sql = "DELETE FROM PartReg WHERE ParticipantID = " & lPartID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sql = "DELETE FROM PartRace WHERE ParticipantID = " & lPartID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sql = "DELETE FROM PreRaceRecips WHERE ParticipantID = " & lPartID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sql = "DELETE FROM PartReminders WHERE PartID = " & lPartID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sql = "DELETE FROM ResultsSent WHERE ParticipantID = " & lPartID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    sql = "DELETE FROM Participant WHERE ParticipantID = " & lPartID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
    Response.Write("<script type='text/javascript'>{window.close();}</script>")
End If

'part details
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, City, St, Phone, Email, DOB, Gender FROM Participant WHERE ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
For i = 0 to 7
	PartData(i) = rs(i).Value
Next
rs.Close
Set rs = Nothing

'get part race history
i = 0
ReDim PartRace(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PartRaceID, RaceID, Age, Bib FROM PartRace WHERE ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    PartRace(0, i) = rs(0).Value
    PartRace(1, i) = RaceName(rs(1).Value)
    PartRace(2, i) = rs(2).Value
    PartRace(3, i) = rs(3).Value
    i = i + 1
    ReDim Preserve PartRace(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get part reg history
i = 0
ReDim PartReg(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PartRegID, RaceID, DateReg, WhereReg FROM PartReg WHERE ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    PartReg(0, i) = rs(0).Value
    PartReg(1, i) = RaceName(rs(1).Value)
    PartReg(2, i) = rs(2).Value
    PartReg(3, i) = rs(3).Value
    i = i + 1
    ReDim Preserve PartReg(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get race history
i = 0
ReDim RaceHist(5, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, EventPl, FnlTime FROM IndResults WHERE ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Call EventInfo(rs(0).Value)

    RaceHist(0, i) = sEventName
    RaceHist(1, i) = dEventDate
    RaceHist(2, i) = sLocation
    RaceHist(3, i) = sRaceName
    RaceHist(4, i) = rs(1).Value
    RaceHist(5, i) = rs(2).Value

    i = i + 1
    ReDim Preserve RaceHist(5, i)

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub EventInfo(lRaceID)
    sEventName = vbNullString
    dEventDate = "1/1/1900"
    sLocation = vbNullString
    sRaceName = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT e.EventName, e.EventDate, e.Location, rd.RaceName FROM Events e INNER JOIN RaceData rd ON e.EventID = rd.EventID WHERE rd.RaceID = "
    sql2 = sql2 & lRaceID
    rs2.Open sql2, conn, 1, 2
    sEventName = Replace(rs2(0).Value, "''", "'")
    dEventDate = rs2(1).Value
    If Not rs2(2).Value & "" = "" Then sLocation = Replace(rs2(2).Value, "''", "'")
    If Not rs2(3).Value & "" = "" Then sRaceName = lRaceID & "-" & Replace(rs2(3).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Sub

Private Function RaceName(lRaceID)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
    rs2.Open sql2, conn, 1, 2
    RaceName = lRaceID & "-" & Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Participant Details</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
	th{
		text-align:right;
		padding:5px 0 0 5px;
        white-space: nowrap;
	}
	
	td{
		padding:5px 0 0 5px;
		text-align:left;
        white-space: nowrap;
	}
</style>
</head>

<body>
<div class="container">
    <h4 class="h4">Edit Participant ID <%=lPartID%></h4>
    <table class="table">
	    <tr>
		    <th>First Name:</th>
		    <td><%=PartData(0)%></td>
		    <th>Last Name:</th>
		    <td><%=PartData(1)%></td>
		    <th>City:</th>
		    <td><%=PartData(2)%></td>
		    <th>State:</th>
		    <td><%=PartData(3)%></td>
	    </tr>
	    <tr>
		    <th>Phone:</th>
		    <td><%=PartData(4)%></td>
		    <th>Email:</th>
		    <td><%=PartData(5)%></td>
		    <th>DOB:</th>
		    <td><%=PartData(6)%></td>
		    <th>Gender:</th>
		    <td colspan="5"><%=PartData(7)%></td>
	    </tr>
        <tr>
             <td style="text-align:center;" colspan="4">
                <form class="form" name="trim_part" method="post" action="part_details.asp?part_id=<%=lPartID%>">
                <input type="hidden" name="submit_trim" id="submit_trim" value="submit_trim">
                <input type="submit" name="submit2" id="submit2" value="Trim Participant Name">
                </form>
            </td>
           <td style="text-align:center;" colspan="4">
                <form class="form" name="delete_part" method="post" action="part_details.asp?part_id=<%=lPartID%>">
                <input type="hidden" name="submit_delete" id="submit_delete" value="submit_delete">
                <input type="submit" name="submit1" id="submit1" value="Delete This Participant">
                </form>
            </td>
        </tr>
    </table>
	
    <%If UBound(RaceHist, 2) > 0 Then%>
	    <h4 class="h4">Race History</h4>

        <table class="table">
            <tr>
                <th style="text-align: left;">Event</th>
                <th style="text-align: left;">Date</th>
                <th style="text-align: left;">Location</th>
                <th style="text-align: left;">Race</th>
                <th style="text-align: center;">Pl</th>
                <th style="text-align: center;">Time</th>
            </tr>
            <%For i = 0 To UBound(RaceHist, 2) - 1%>
                <tr>
                    <%For j = 0 To 5%>
                        <td><%=RaceHist(j, i)%></td>
                    <%Next%>
                </tr>
            <%Next%>
        </table>
    <%End If%>

    <hr>

    <table class="table">
        <tr>
            <td style="width: 300px;" valign="top">
                <h4 class="h4">Part Race Data</h4>
                <table class="table">
                    <tr>
                        <th style="text-align: left;">ID</th>
                        <th style="text-align: left;">Race</th>
                        <th style="text-align: left;">Age</th>
                        <th style="text-align: left;">Bib</th>
                    </tr>
                    <%For i = 0 To UBound(PartRace, 2) - 1%>
                        <tr>
                            <%For j = 0 To 3%>
                                <td><%=PartRace(j, i)%></td>
                            <%Next%>
                        </tr>
                    <%Next%>
                </table>
            </td>
            <td style="width: 300px;" valign="top">
                <h4 class="h4">Part Reg Data</h4>
                <table class="table">
                    <tr>
                        <th style="text-align: left;">ID</th>
                        <th style="text-align: left;">Race</th>
                        <th style="text-align: left;">Date Reg</th>
                        <th style="text-align: left;">Where Reg</th>
                    </tr>
                    <%For i = 0 To UBound(PartReg, 2) - 1%>
                        <tr>
                            <%For j = 0 To 3%>
                                <td><%=PartReg(j, i)%></td>
                            <%Next%>
                        </tr>
                    <%Next%>
                </table>
            </td>
        </tr>
    </table>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>