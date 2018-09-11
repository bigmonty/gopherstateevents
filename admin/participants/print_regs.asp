<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID
Dim sEventName, sRace, dEventDate
Dim PartArray(), RaceArray(), Races(), TempArray(14)

lEventID = Request.QueryString("event_id")

lRaceID = Request.QueryString("race_id")
If Not IsNumeric(lRaceID) Then Response.Redirect "htttp://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
sEventName = rs(0).Value
dEventDate = rs(1).Value
Set rs = Nothing

i = 0
ReDim Races(0)
sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(i) = rs(0).Value
	i = i + 1
	ReDim Preserve Races(i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim PartArray(14, 0)					
If lRaceID = "0" Then
	sRace = "All Races"

	i = 0
	For k = 0 to UBound(Races) - 1
		sql="SELECT p.ParticipantID, p.FirstName, p.LastName, rc.Bib, p.Gender, rc.Age, p.City, p.St, "
		sql = sql & " p.Phone, rg.ShrtSize, rg.ShrtStyle, rg.WhereReg, rg.DateReg, p.DOB, rg.AmtPd, p.Email FROM "
		sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
		sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE (rc.RaceID = " & Races(k) & " AND rg.RaceID = " 
		sql = sql & Races(k) & ") ORDER BY p.LastName, p.FirstName"
		Set rs=conn.Execute(sql)
		Do While Not rs.EOF
			PartArray(0, i) = rs(0).value
			PartArray(1, i) = rs(2).Value & ", " & rs(1).value
			For j = 2 to 14
				PartArray(j, i) = rs(j + 1).Value
			Next
			i = i + 1
			ReDim Preserve PartArray(14, i)
			rs.MoveNext
		Loop
		Set rs=Nothing
	Next
		
	'sort the array
	For i = 0 to UBound(PartArray, 2) - 2
		For j = i + 1 to UBound(PartArray, 2) - 1
			If CStr(PartArray(1, i)) > CStr(PartArray(1, j)) Then
				For k = 0 to 14
					TempArray(k) = PartArray(k, i)
					PartArray(k, i) = PartArray(k, j)
					PartArray(k, j) = TempArray(k)
				Next
			End If
		Next
	Next
Else
	sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	sRace = rs(0).Value
	Set rs = Nothing

	sql="SELECT p.ParticipantID, p.FirstName, p.LastName, rc.Bib, p.Gender, rc.Age, p.City, p.St, "
	sql = sql & " p.Phone, rg.ShrtSize, rg.ShrtStyle, rg.WhereReg, rg.DateReg, p.DOB, rg.AmtPd, p.Email FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE (rc.RaceID = " & lRaceID & " AND rg.RaceID = " 
	sql = sql & lRaceID & ") ORDER BY p.LastName, p.FirstName"
	Set rs=conn.Execute(sql)
	Do While Not rs.EOF
		PartArray(0, i) = rs(0).value
		PartArray(1, i) = rs(2).Value & ", " & rs(1).value
		For j = 2 to 14
			PartArray(j, i) = rs(j + 1).Value
		Next
		i = i + 1
		ReDim Preserve PartArray(14, i)
		rs.MoveNext
	Loop
	Set rs=Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title>Print <%=sEventName%> Participants</title>
<!--#include file = "../../includes/meta2.asp" -->




<style type="text/css">
<!--
td,th{
	padding-right:5px;
	padding-left:5px;
	font-size:0.8em;
	}
-->
</style>
</head>

<body style="background-image:none;background-color:#fff;">
<div style="text-align:left;margin:10px;">
	<a href="javascript:window.print()" style="font-size:0.8em;">Print This</a>
	
	<h5 style="margin:5px 0 5px 0;background-color:#ececd8;">Registration Data For <%=sEventName%>&nbsp;(<%=sRace%>) on <%=dEventDate%></h5>

	<table style="border-collapse:collapse;width:790px;">
		<tr>
			<th style="text-align:center;width:10px">No</th>
			<th>Name</th>
			<th style="text-align:center;">Bib</th>
			<th style="text-align:center;">M/F</th>
			<th style="text-align:center;">Age</th>
			<th>City</th>
			<th style="text-align:center;">St</th>
			<th style="text-align:center;">Phone</th>
			<th style="text-align:center;">Size</th>
			<th style="text-align:center;">Style</th>
			<th style="text-align:center;">Where</th>
			<th style="text-align:center;">Date</th>
			<th style="text-align:center;">DOB</th>
			<th style="text-align:center;">Paid</th>
		</tr>
		<%For j = 0 to UBound(PartArray, 2) - 1%>
			<%If j mod 2 = 0 Then%>
				<tr>
					<td class="alt" style="text-align:right;width:10px;">
						<%=j+1%>)
					</td>
					<%For i = 1 to 14%>
						<td class="alt" style="white-space:nowrap;">
							<%=PartArray(i, j)%>
						</td>
					<%Next%>
				</tr>
			<%Else%>
				<tr>
					<td style="text-align:right;width:10px">
						<%=j+1%>)
					</td>
					<%For i = 1 to 14%>
						<td style="white-space:nowrap;">
							<%=PartArray(i, j)%>
						</td>
					<%Next%>
				</tr>
			<%End If%>
		<%Next%>
	</table>
</div>
</body>
<%
conn.Close
Set conn=Nothing
%>
</html>
