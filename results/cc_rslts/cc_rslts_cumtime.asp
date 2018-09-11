<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs
Dim i, j, k
Dim lMeetID
Dim iNumRnrs
Dim sMeetName, sRaces, sRaceNames, sShowInd, sRaceID, sGradeYear
Dim MeetTeams,  MeetRaces(), SortArr(3), OurIndiv(), SelRaces()
Dim dMeetDate

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

sShowInd = Request.QueryString("show_ind")
If sShowInd = vbNullString Then sShowInd = "n"

iNumRnrs = Request.QueryString("num_rnrs")
If CStr(iNumRnrs) = vbNullString Then iNumRnrs = 7

sRaces = Request.QueryString("races")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value
Set rs = Nothing
		
'get year for roster grades
If Month(dMeetDate) <= 7 Then
	sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
	sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

i = 0
ReDim MeetRaces(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceDesc,RaceDist, RaceUnits FROM Races WHERE MeetsID = " & lMeetID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    MeetRaces(0, i) = rs(0).Value
    MeetRaces(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & " " & rs(3).Value & ")"
    i = i + 1
    ReDim Preserve MeetRaces(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

ReDim SelRaces(0)
If Request.Form.Item("submit_this") = "submit_this" Then
    iNumRnrs = Request.Form.Item("num_rnrs")
	sRaces = Request.Form.Item("races")
End If

If Not sRaces = vbNullString Then

    'get included races by name
    If InStr(sRaces, ",") = 0 Then
        GetRaceName(CLng(sRaces))
    Else
        j = 0
        For i = 1 To Len(sRaces)
            If Mid(sRaces, i, 1) = "," Then
                SelRaces(j) = sRaceID
                j = j + 1
                ReDim Preserve SelRaces(j)

                If sRaceNames = vbNullString Then
                    sRaceNames = GetRaceName(CLng(sRaceID))
                Else
                    sRaceNames = sRaceNames & ", " & GetRaceName(CLng(sRaceID))
                End If
                sRaceID = vbNullString
            Else
                sRaceID = sRaceID & Mid(sRaces, i, 1)
                If i = Len(sRaces) Then 
                    SelRaces(j) = sRaceID
                    j = j + 1
                    ReDim Preserve SelRaces(j)

                    sRaceNames = sRaceNames & ", " & GetRaceName(CLng(sRaceID))
                End If
            End If
        Next
    End If
End If

ReDim MeetTeams(3, 0)

If Not sRaces = vbNullString Then
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT t.TeamsID, t.TeamName, t.Gender FROM Teams t INNER JOIN MeetTeams mt ON t.TeamsID = mt.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lMeetID & " ORDER BY t.TeamName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MeetTeams(0, i) = rs(0).Value
        MeetTeams(1, i) = Replace(rs(1).Value, "''", "'")
        MeetTeams(2, i) = rs(2).Value
        MeetTeams(3, i) = OurTime(rs(0).Value)
        i = i + 1
        ReDim Preserve MeetTeams(3, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(MeetTeams, 2) - 2
        For j = i + 1 To UBound(MeetTeams, 2) - 1
            If CSng(MeetTeams(3, i)) > CSng(MeetTeams(3, j)) Then
                For k = 0 To 3
                    SortArr(k) = MeetTeams(k, i)
                    MeetTeams(k, i) = MeetTeams(k, j)
                    MeetTeams(k, j) = SortArr(k)
                Next    
            End If
        Next
    Next
End If

If CStr(iNumRnrs) = vbNullString Then iNumRnrs = 2

Private Function OurTime(lThisTeam)
    Dim rs2, sql2
    Dim x

    x = 0
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT ir.FnlScnds FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE r.TeamsID = " & lThisTeam & " AND ir.RacesID IN ("
    sql2 = sql2 & sRaces & ") AND ir.FnlScnds > 0 AND ir.Excludes = 'n' ORDER BY ir.FnlScnds, ir.Place"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        x = x + 1
        OurTime = CSng(OurTime) + CSng(rs2(0).Value)
        If CInt(x) = CInt(iNumRnrs) Then Exit Do
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If CInt(x) < CInt(iNumRnrs) Then OurTime = 0
End Function

Private Sub OurIndividuals(lThisTeam)
    Dim rs2, sql2
    Dim x

    x = 0
    ReDim OurIndiv(2, 0)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT r.FirstName, r.LastName, ir.FnlScnds, g.Grade" & sGradeYear & " FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID "
    sql2 = sql2 & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE r.TeamsID = " & lThisTeam & " AND ir.RacesID IN ("& sRaces 
    sql2 = sql2 & ") AND ir.FnlScnds > 0 ORDER BY ir.FnlScnds, ir.Place"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        OurIndiv(0, x) = Replace(rs2(0).Value, "''", "'") & " " & Replace(rs2(1).Value, "''", "'")
        OurIndiv(1, x) = ConvertToMinutes(rs2(2).Value)
        OurIndiv(2, x) = rs2(3).Value
        If CInt(x) = CInt(iNumRnrs) Then Exit Do
        x = x + 1
        ReDim Preserve OurIndiv(2, x)
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End Sub

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If CSng(sglScnds) <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(CSng(sglScnds) \ 60)
    sSecondPart = CStr(((CSng(sglScnds) / 60) - (CSng(sglScnds) \ 60)) * 60)
    
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

        ConvertToMinutes = Replace(ConvertToMinutes, "-", "")
    End If
End Function

Private Function GetRaceName(lThisRace)
    GetRaceName = "undetermined"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceDesc FROM Races WHERE RacesID = " & lThisRace
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetRaceName = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Gopher State Events CC/Nordic Results By Cumulative Time</title>
<meta name="description" content="Cross-Country & Nordic Ski Results by Gopher State Events, a conventional timing service offererd by H51 Software, LLC in Minnetonka, MN.">
<!--#include file = "../../includes/js.asp" --> 

<style type="text/css">
    .form-inline .form-control {
        vertical-align: top;
    }
</style>
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-sm-6">
            <br>
            <img src="/graphics/html_header.png" class="img-responsive" alt="Individual Results">
        </div>
        <div class="col-sm-6">
	        <h2 class="h2">GSE Cumulative Time Results</h2>
            <h3 class="h3"><%=sMeetName%></h3>
            <h4 class="h4"><%=dMeetDate%></h4>
        </div>
    </div>
    <div class="row">
        <p>This utility will allow you to determine cumulative times for teams over one or more races in a cross-country or Nordic ski meet.</p>

	    <div class="bg-warning">
		    <a href="javascript:window.print();">Print</a>
	    </div>
    </div>

    <div class="row bg-success">
        <form class="form-inline" name="get_race_parts" method="post" action="cc_rslts_cumtime.asp?meet_id=<%=lMeetID%>&amp;show_ind=<%=sShowInd%>">
		<div class="form-group">
            <label for="races">Select Races:</label>
			<select class="form-control" name="races" id="races" multiple size="8">
				<%For i = 0 to UBound(MeetRaces, 2)%>
					<option value="<%=MeetRaces(0, i)%>"><%=MeetRaces(1, i)%></option>
				<%Next%>
			</select>
            &nbsp;
            <label for="num_rnrs">Runners per Team to Score:</label>
            <select class="form-control" name="num_rnrs" id="num_rnrs">
                <%For i = 2 To 10%>
                    <%If CInt(i) = CInt(iNumRnrs) Then%>
                        <option value="<%=i%>" selected><%=i%></option>
                    <%Else%>
                        <option value="<%=i%>"><%=i%></option>
                    <%End If%>
                <%Next%>
            </select>
			<input class="form-control" type="hidden" name="submit_this" id="submit_this" value="submit_this">
			<input class="form-control" type="submit" name="get_race" id="get_race" value="Get Results">
        </div>
        </form>
    </div>

    <%If Not sRaces = vbNullString Then%>
        <div class="row">
            <h4 class="h4">Cumulative Time Results for <%=sRaceNames%> (<%=iNumRnrs%> runners per team)</h4>
            <div class="bg-info">
                <%If sShowInd = "y" Then%>
                    <a href="cc_rslts_cumtime.asp?meet_id=<%=lMeetID%>&amp;num_rnrs=<%=iNumRnrs%>&amp;races=<%=sRaces%>&amp;show_ind=n">Hide Individuals</a>
                <%Else%>
                    <a href="cc_rslts_cumtime.asp?meet_id=<%=lMeetID%>&amp;num_rnrs=<%=iNumRnrs%>&amp;races=<%=sRaces%>&amp;show_ind=y">Show Individuals</a>
                <%End If%> 
            </div>
	        <table class="table table-striped">
		        <tr>
			        <th>Pl</th>
			        <th>Team</th>
                    <th>Gender</th>
			        <th>Time</th>
		        </tr>
                <%k = 1%>
			    <%For i = 0 to UBound(MeetTeams, 2) - 1%>
                    <%If CSng(MeetTeams(3, i)) > 0 Then%>
				        <tr>
					        <td><%=k%></td>
					        <td><%=MeetTeams(1, i)%>
					        <td><%=MeetTeams(2, i)%></td>
					        <td><%=ConvertToMinutes(MeetTeams(3, i))%></td>
				        </tr>
                        <%If sShowInd = "y" Then%>
                            <%Call OurIndividuals(MeetTeams(0, i))%>
                            <tr>
                                <td style="padding-left: 150px;" colspan="4">
                                    <ol class="list-group">
                                        <%For j = 0 To UBound(OurIndiv, 2) - 1%>
                                            <li class="list-group-item"><%=OurIndiv(0, j)%> (<%=OurIndiv(2, j)%>): <%=OurIndiv(1, j)%></li>
                                        <%Next%>
                                    </ol>
                                </td>
                            </tr>
                        <%End If%>
                        <%k = k + 1%>
                    <%End If%>
			    <%Next%>
		    </table>
        </div>
    <%End If%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
