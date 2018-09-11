<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim sMeetName, sTeamName, sGender, sSport
Dim RaceArr(), RacePartsArr()
Dim iNumRaces
Dim i, j, k
Dim lMeetID, lTeamID
Dim sGradeYear

If Not Session("role") = "coach" Then Response.Redirect "/default.asp?sign_out=y"

lMeetID = Request.QueryString("meet_id")
lTeamID = Request.QueryString("team_id")
	
Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate, Sport FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value & " on " & rs(1).Value 
If Month(rs(1).Value) <=7 Then
	sGradeYear = Right(CStr(Year(rs(1).Value) - 1), 2)
Else
	sGradeYear = Right(CStr(Year(rs(1).Value)), 2)	
End If
sSport = rs(2).Value
Set rs = Nothing

'get team name, gender
sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = rs(0).Value
sGender = rs(1).Value
Set rs = Nothing

'convert gender to full word
Select Case sGender
	Case "M"
		sGender = "Male"
	Case "F"
		sGender = "Female"
End Select

'get num races
iNumRaces = 1
i = 0
ReDim RaceArr(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceName, RaceTime FROM Races WHERE MeetsID = " & lMeetID & " AND (Gender = '" & sGender 
sql = sql & "' OR Gender = 'Open')"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	iNumRaces = iNumRaces + 1
	RaceArr(0, i) = rs(0).Value
	RaceArr(1, i) = rs(1).Value
	RaceArr(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve RaceArr(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Sub GetRaceParts(lRaceID)
	j = 0
	ReDim RacePartsArr(3, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.Bib, ir.IndDelay, ir.Gate FROM Roster r INNER JOIN Grades g "
	sql = sql & "ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
	sql = sql & "WHERE TeamsID = " & lTeamID & " AND ir.RacesID = " & lRaceID & " AND r.Archive = 'n' "
	sql = sql & "ORDER BY ir.IndDelay, r.LastName, g.Grade" & sGradeYear
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		RacePartsArr(0, j) = rs(1).Value & ", " & rs(0).Value & " (" & rs(3).Value & ")"
		RacePartsArr(1, j) = rs(2).Value
        RacePartsArr(2, j) = ConvertToMinutes(rs(4).Value)
        RacePartsArr(3, j) = rs(5).Value
		j = j + 1
		ReDim Preserve RacePartsArr(3, j)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End Sub

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If sglScnds <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(sglScnds \ 60)
    sSecondPart = CStr(((sglScnds / 60) - (sglScnds \ 60)) * 60)
    
    'add leading zero to seconds if necessary
    If CSng(sSecondPart) < 10 Then
        sSecondPart = "0" & sSecondPart
    End If
    
    'make sure there are exactly two decimal places
    If Len(sSecondPart) < 5 Then
        If Len(sSecondPart) = 2 Then
            sSecondPart = sSecondPart & ".00"
        ElseIf Len(sSecondPart) = 4 Then
            sSecondPart = sSecondPart & "0"
        End If
    Else
        sSecondPart = Left(sSecondPart, 5)
    End If
    
    'do the conversion
    If CInt(sMinutePart) <= 60 Then
        ConvertToMinutes = sMinutePart & ":" & sSecondPart
    Else
        sHourPart = CStr(CSng(sMinutePart) \ 60)
        sMinutePart = CStr(CSng(sMinutePart) Mod 60)

        If Len(sMinutePart) = 1 Then
            sMinutePart = "0" & sMinutePart
        End If

        ConvertToMinutes = sHourPart & ":" & sMinutePart & ":" & sSecondPart
    End If
End Function
%>
<html lang="en">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>GSE CC Meet Sheet</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="keywords" content="running, nordic skiing, cross-country, mountain biking, road races, snowshoe, race, timing, ">
<meta name="description" content="A Fitness Event Timing Service specializing in road racing, nordic ski events, showshoe events, mountain bike events, and high school and college cross-country meet timing.">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">

<script type="text/javascript" src="../../../misc/vira.js"></script>
<link rel="stylesheet" type="text/css" href="../../../misc/vira.css">

<style type="text/css">
<!--
td{
	color:#000000;
	font-size:10pt;
	padding:5px;
	border:1px solid #000000
	}
-->
</style>

</head>
<body>
<a href="javascript:window.print()">Print</a><br>
<%If sSport = "Nordic Ski" Then%>
    <%For i = 0 to UBound(RaceArr, 2) - 1%>
	    <table style="page-break-after:always;">
		    <tr>
			    <th style="color:#000000;border: none;font-size: 1.1em;text-align:left;" colspan="11">Meet Sheet for <%=sTeamName%> for <%=sMeetName%></th>
		    </tr>
		    <%Call GetRaceParts(RaceArr(0, i))%>
		    <tr>
			    <td style="font-weight:bold;border:none" colspan="2">Race Name: <%=RaceArr(1, i)%></td>
			    <td style="text-align:right;font-weight:bold;border:none" colspan="7">Race Time: <%=RaceArr(2, i)%></td>
		    </tr>
		    <tr>
			    <td style="font-weight:bold;padding:2px;width:10px" rowspan="2" valign="bottom">No.</td>
			    <td style="font-weight:bold;padding:2px" rowspan="2" valign="bottom">Name</td>
			    <td style="font-weight:bold;padding:2px" rowspan="2" valign="bottom">Gr</td>
                <td style="font-weight:bold;padding:2px" rowspan="2" valign="bottom">Start</td>
                <td style="font-weight:bold;padding:2px" rowspan="2" valign="bottom">Gate</td>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" rowspan="2" valign="bottom">Pl</td>
			    <td style="font-weight:bold;padding-left:25px;padding-right:25px" rowspan="2" valign="bottom">Time</td>
			    <td style="font-weight:bold;padding:2px;text-align:center" colspan="3">Splits</td>
			    <td style="text-align:center;font-weight:bold;padding:2px;width:200px" rowspan="2" valign="bottom">Comments</td>
		    </tr>
		    <tr>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">Split 1</td>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">Split 2</td>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">Split 3</td>
		    </tr>
		    <%For j = 0 to UBound(RacePartsArr, 2) - 1%>
			    <tr>
				    <td style="width:10px"><%=j + 1%></td>
				    <td nowrap="nowrap"><%=RacePartsArr(0, j)%></td>
				    <td><%=RacePartsArr(1, j)%></td>
                    <td><%=RacePartsArr(2, j)%></td>
                    <td><%=RacePartsArr(3, j)%></td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
			    </tr>
		    <%Next%>
		    <%For k = j to j + 4%>
			    <tr>
				    <td style="width:10px"><%=k + 1%></td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
			    </tr>
		    <%Next%>
	    </table>
    <%Next%>
<%Else%>
    <%For i = 0 to UBound(RaceArr, 2) - 1%>
	    <table style="border:1px solid #000000;page-break-after:always;">
		    <tr>
			    <td class="table_head" style="color:#000000" colspan="9">Meet Sheet for <%=sTeamName%> for <%=sMeetName%></td>
		    </tr>
		    <%Call GetRaceParts(RaceArr(0, i))%>
		    <tr>
			    <td style="font-weight:bold;border:none" colspan="2">Race Name: <%=RaceArr(1, i)%></td>
			    <td style="text-align:right;font-weight:bold;border:none" colspan="7">Race Time: <%=RaceArr(2, i)%></td>
		    </tr>
		    <tr>
			    <td style="font-weight:bold;padding:2px;width:10px" rowspan="2" valign="bottom">No.</td>
			    <td style="font-weight:bold;padding:2px" rowspan="2" valign="bottom">Name</td>
			    <td style="font-weight:bold;padding:2px" rowspan="2" valign="bottom">Gr</td>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" rowspan="2" valign="bottom">Pl</td>
			    <td style="font-weight:bold;padding-left:25px;padding-right:25px" rowspan="2" valign="bottom">Time</td>
			    <td style="font-weight:bold;padding:2px;text-align:center" colspan="3">Splits</td>
			    <td style="text-align:center;font-weight:bold;padding:2px;width:200px" rowspan="2" valign="bottom">Comments</td>
		    </tr>
		    <tr>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">Split 1</td>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">Split 2</td>
			    <td style="font-weight:bold;padding-left:10px;padding-right:10px" nowrap="nowrap">Split 3</td>
		    </tr>
		    <%For j = 0 to UBound(RacePartsArr, 2) - 1%>
			    <tr>
				    <td style="width:10px"><%=j + 1%></td>
				    <td nowrap="nowrap"><%=RacePartsArr(0, j)%></td>
				    <td><%=RacePartsArr(1, j)%></td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
			    </tr>
		    <%Next%>
		    <%For k = j to j + 4%>
			    <tr>
				    <td style="width:10px"><%=k + 1%></td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
			    </tr>
		    <%Next%>
	    </table>
    <%Next%>
<%End If%>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
