<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisTeam
Dim iFirstBib, iLastBib
Dim sTeamName, sTeamGender, sCoachName, sCoachPhone, sCoachEmail
Dim BibRange(), Delete()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisTeam = Request.QueryString("this_team")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamBibsID, FirstBib, LastBib FROM TeamBibs WHERE TeamsID = " & lThisTeam
    rs.Open sql, conn, 1,  2
    Do While Not rs.EOF
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
            rs(1).Value = Request.Form.Item("first_bib_" & rs(0).Value)
            rs(2).Value = Request.Form.Item("last_bib_" & rs(0).Value)
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM TeamBibs WHERE TeamBibsID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
    Set rs = Nothing
ElseIf Request.Form.Item("submit_range") = "submit_range" Then
	iLastBib = Request.Form.Item("last_bib")
	iFirstBib = Request.Form.Item("first_bib")

	sql = "INSERT INTO TeamBibs (TeamsID, FirstBib, LastBib) VALUES (" & lThisTeam & ", " & iFirstBib & ", " & iLastBib & ")"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

'get team name
sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lThisTeam
Set rs = conn.Execute(sql)
sTeamName = Replace(rs(0).Value, "''", "'") & " (Gender: " & rs(1).Value & ")"
If rs(1).Value = "F" Then
	sTeamGender = "Female"
Else
	sTeamGender = "Male"
End If
Set rs = Nothing
	
sql = "SELECT c.FirstName, c.LastName, c.Phone, c.Email FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID WHERE t.TeamsID = " & lThisTeam
Set rs = conn.Execute(sql)
sCoachName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
sCoachPhone = rs(2).Value
sCoachEmail = rs(3).Value
Set rs = Nothing

'get bib range
i = 0
ReDim BibRange(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamBibsID, FirstBib, LastBib FROM TeamBibs WHERE TeamsID = " & lThisTeam
rs.Open sql, conn, 1,  2
Do While Not rs.EOF
    BibRange(0, i) = rs(0).Value
    BibRange(1, i) = rs(1).Value
    BibRange(2, i) = rs(2).Value
    i = i + 1
    ReDim Preserve BibRange(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Nordic Ski Bib Range Assignment</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body onload="javascript:set_bib_range.first_bib.focus()">
<div class="container">
	<h4 class="h4">Bib Range Assignments for <%=sTeamName%></h4>
							
	<form name="set_bib_range" method="Post" action="bib_range.asp?this_team=<%=lThisTeam%>">
	<table>
        <tr><th colspan="5">Add New Range:</th></tr>
		<tr>
			<th>From</th>
			<td><input type="text" name="first_bib" id="first_bib" size="4"></td>
			<th>To</th>
			<td><input type="text" name="last_bib" id="last_bib" size="4"></td>
			<td>
				<input type="hidden" name="submit_range" id="submit_range" value="submit_range">
				<input type="submit" name="submit1" id="submit1" value="Add Bib Range">
			</td>
		</tr>
	</table>
	</form>

    <br><br>

	<form name="edit_bib_range" method="Post" action="bib_range.asp?this_team=<%=lThisTeam%>">
	<table>
        <tr><th colspan="4">Existing Bib Ranges:</th></tr>
		<%For i = 0 To UBound(BibRange, 2) - 1%>
            <tr>
			    <th>From</th>
			    <td><input type="text" name="first_bib_<%=BibRange(0, i)%>" id="first_bib_<%=BibRange(0, i)%>" size="4" value="<%=BibRange(1, i)%>"></td>
			    <th>To</th>
			    <td><input type="text" name="last_bib_<%=BibRange(0, i)%>" id="last_bib_<%=BibRange(0, i)%>" size="4" value="<%=BibRange(2, i)%>"></td>
                <td style="color: red;"><input type="checkbox" name="delete_<%=BibRange(0, i)%>" id="delete_<%=BibRange(0, i)%>">Delete</td>
            </tr>
        <%Next%>
        <tr>
			<td  style="text-align:center;" colspan="4">
				<input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
				<input type="submit" name="submit2" id="submit2" value="Submit Changes">
			</td>
		</tr>
	</table>
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
