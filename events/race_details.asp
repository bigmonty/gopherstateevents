<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lRaceID, lMeetID
Dim sMeetName
Dim dMeetDate
Dim RaceDetails(11)
Dim i

lRaceID = Request.QueryString("race_id")

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect "htttp://www.google.com"
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
If Not rs(0).Value & "" = "" Then sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

sql = "SELECT RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, ScoreMethod, NumAllow, NumScore, Comments, "
sql = sql & "TmAwds, IndAwds, StartType FROM Races WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
For i = 0 to 11
	RaceDetails(i) = rs(i).Value
Next
Set rs = Nothing

If RaceDetails(6) = 0 Then RaceDetails(6) = "Unlimited"

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Race Details</title>
<!--#include file = "../includes/js.asp" -->
<style type="text/css">
    td{
        text-align: left;
    }
    th{
        text-align: right;
    }
</style>
</head>
<body>
<div style="margin:10px;background-color:#fff;">
    <div style="text-align: right;background-color: #ececec;margin: 10px;padding: 5px;font-size: 0.8em;">
        <a href="javascript:print();">Print</a>
    </div>

    <h3><%=sMeetName%> on <%=dMeetDate%></h3>
    <h4 style="border: none;background: none;">Race Details For <%=RaceDetails(0)%></h4>

	<table>
		<tr>
			<th>Race Time:</th>
			<td><%=RaceDetails(1)%></td>
		</tr>
		<tr>
			<th>Race Distance:</th>
			<td><%=RaceDetails(2)%> <%=RaceDetails(3)%></td>
		</tr>
		<tr>
			<th>Race Gender:</th>
			<td style="width:60%;text-align:left"><%=RaceDetails(4)%></td>
		</tr>
		<tr>
            <th>Scoring Method:</th>
			<td>by <%=RaceDetails(5)%></td>
		</tr>
		<tr>
			<th>Start Type:</th>
			<td><%=RaceDetails(11)%></td>
		</tr>
		<tr>
			<th>Participants per Team:</th>
			<td><%=RaceDetails(6)%></td>
		</tr>
		<tr>
			<th>Number Scoring per Team:</th>
			<td><%=RaceDetails(7)%></td>
		</tr>
		<tr>
			<th valign="top">Team Awards:</th>
			<td><%=RaceDetails(9)%></td>
		</tr>
		<tr>
			<td style="width:40%;font-weight:bold;text-align:right" valign="top">
				Individual Awards:
			</td>
			<td style="width:60%;width:60%;text-align:left">
				<%=RaceDetails(10)%>
			</td>
		</tr>
		<tr>
			<td style="width:40%;font-weight:bold;text-align:right" valign="top">
				Comments:
			</td>
			<td style="width:60%;text-align:left">
				<%If Not RaceDetails(8) = vbNull Then Response.Write(Replace(RaceDetails(8), "''", "'"))%>
			</td>
		</tr>
	</table>
</div>
</body>
</html>
