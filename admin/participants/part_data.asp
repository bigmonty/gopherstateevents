<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID
Dim sEventName, sRace, sSortBy, sOrderBy, sEventRaces, sAlphaStart, sAlphaEnd, iTtlParts, sAgeGrpUpdate, sCleanData
Dim iNumRcds
Dim dEventDate
Dim PartArray, RaceArray(), Races(), TempArray(10), AlphaArr(25)

lEventID = Request.QueryString("event_id")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0

sAgeGrpUpdate = Request.QueryString("age_grp_update")
If sAgeGrpUpdate = vbNullString Then sAgeGrpUpdate = "n"

sAlphaStart = Request.QueryString("alpha_start")
sAlphaStart = "A"

sAlphaEnd = Request.QueryString("alpha_end")
sAlphaEnd = "Z"

iNumRcds = Request.QueryString("num_rcds")
If CStr(iNumRcds) = vbNullString Then iNumRcds = 1500

iTtlParts = 0

sSortBy = Request.QueryString("sort_by")

sCleanData = Request.QueryString("clean_data")
If sCleanData = vbNullString Then sCleanData = "n"

If sSortBy = "bib" Then
	sOrderBy = "ORDER BY rc.Bib"
Else
	sOrderBy = "ORDER BY p.LastName, p.FirstName"
End If

AlphaArr(0) = "A"
AlphaArr(1) = "B"
AlphaArr(2) = "C"
AlphaArr(3) = "D"
AlphaArr(4) = "E"
AlphaArr(5) = "F"
AlphaArr(6) = "G"
AlphaArr(7) = "H"
AlphaArr(8) = "I"
AlphaArr(9) = "J"
AlphaArr(10) = "K"
AlphaArr(11) = "L"
AlphaArr(12) = "M"
AlphaArr(13) = "N"
AlphaArr(14) = "O"
AlphaArr(15) = "P"
AlphaArr(16) = "Q"
AlphaArr(17) = "R"
AlphaArr(18) = "S"
AlphaArr(19) = "T"
AlphaArr(20) = "U"
AlphaArr(21) = "V"
AlphaArr(22) = "W"
AlphaArr(23) = "X"
AlphaArr(24) = "Y"
AlphaArr(25) = "Z"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim Events
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing
		
i = 0
ReDim Races(1, 0)
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Races(1, i)

    sEventRaces = sEventRaces & rs(0).Value & ", "

	rs.MoveNext
Loop
Set rs = Nothing

If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

If UBound(Races, 2) = 1 Then lRaceID = Races(0, 0)

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_alpha") = "submit_alpha" Then
    sAlphaStart = Request.Form.Item("alpha_start")
    sAlphaEnd = Request.Form.Item("alpha_end")
    iNumRcds = Request.Form.Item("num_rcds")

    If CStr(iNumRcds) = vbNullString Then iNumRcds = GetNumRcds(lRaceID)
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
	If CStr(lRaceID) = vbNullString Then lRaceID = 0

    iNumRcds = GetNumRcds(lRaceID)
ElseIf Request.Form.Item("submit_bibs") = "submit_bibs" Then
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT ParticipantID, Bib FROM PartRace WHERE RaceID IN (" & sEventRaces & ")"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		If Request.Form.Item("edit_" & rs(0).Value) = "y" Then
            rs(1).Value = Request.Form.Item("bib_" & rs(0).Value)
			rs.Update
        End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs=Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT p.ParticipantID FROM Participant p INNER JOIN PartRace rc ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces & ")"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iTtlParts = rs.RecordCount
rs.Close
Set rs=Nothing

If sCleanData = "y" Then
    For i = 0 To UBound(Races, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT p.FirstName, p.LastName, p.Email FROM PartRace pr INNER JOIN Participant p ON "
        sql = sql & "pr.ParticipantID = p.ParticipantID WHERE pr.RaceID = " & Races(0, i) & " ORDER BY p.LastName, p.FirstName"
        rs.Open sql, conn, 1, 2
        For j = 0 To rs.RecordCount - 1
            rs(0).Value = Replace(rs(0).Value, Chr(34), "")
            rs(1).Value = Replace(rs(1).Value, Chr(34), "")
            rs(2).Value = Replace(rs(2).Value, Chr(34), "")
            rs.Update
            rs.MoveNext
        Next
        rs.Close
        Set rs = Nothing
    Next
End If

If sAgeGrpUpdate = "y" Then
    'fix age groups for all participants in each race
    For i = 0 To UBound(Races, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT p.Gender, pr.Age, pr.AgeGrp FROM PartRace pr INNER JOIN Participant p ON "
        sql = sql & "pr.ParticipantID = p.ParticipantID WHERE pr.RaceID = " & Races(0, i) & " ORDER BY pr.Age DESC"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 1 Then
            For j = 0 To rs.RecordCount - 1
                rs(2).Value = GetAgeGrp(rs(0).Value, rs(1).Value, Races(0, i))
                rs.Update
                rs.MoveNext
            Next
        End If
        rs.Close
        Set rs = Nothing
    Next
End If

If lRaceID = "0" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, rc.Bib, p.Gender, rc.Age, p.City, p.St, p.DOB, p.Email, rg.RaceID FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces & ") AND rc.RAceID = rg.RaceID AND Left(p.LastName, 1) BETWEEN '" 
    sql = sql & sAlphaStart & "' AND '" & sAlphaEnd & "' " & sOrderBy
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        PartArray=rs.GetRows(iNumRcds)		
    Else
        ReDim PartArray(10, 0)
    End If
    rs.Close
	Set rs=Nothing

	'sort the array
	If sSortBy = "bib" Then
		For i = 0 to UBound(PartArray, 2) - 2
			For j = i + 1 to UBound(PartArray, 2) - 1
				If CInt(PartArray(3, i)) > CInt(PartArray(3, j)) Then
					For k = 0 to 10
						TempArray(k) = PartArray(k, i)
						PartArray(k, i) = PartArray(k, j)
						PartArray(k, j) = TempArray(k)
					Next
				End If
			Next
		Next
	End If
Else
	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, rc.Bib, p.Gender, rc.Age, p.City, p.St, p.DOB, p.Email, rg.RaceID FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID = " & lRaceID & " AND rc.RAceID = rg.RaceID AND Left(p.LastName, 1) BETWEEN '" 
    sql = sql & sAlphaStart & "' AND '" & sAlphaEnd & "' " & sOrderBy
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        PartArray=rs.GetRows(iNumRcds)		
    Else
        ReDim PartArray(10, 0)
    End If
	rs.Close
	Set rs=Nothing
End If
	
If Not lRaceID = "0" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
	rs.Open sql, conn, 1, 2
	sRace = Replace(rs(0).Value, "''", "'")
	rs.Close
	Set rs = Nothing
End If

Private Function GetNumRcds(lWhichRace)
    If CLng(lWhichRace) = 0 Then
		sql="SELECT p.LastName FROM Participant p INNER JOIN PartRace rc ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces 
        sql = sql & ") AND Left(p.LastName, 1) BETWEEN '" & sAlphaStart & "' AND '" & sAlphaEnd & "'"
    Else
		sql="SELECT p.LastName FROM Participant p INNER JOIN PartRace rc ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID = " & lWhichRace
        sql = sql & " AND Left(p.LastName, 1) BETWEEN '" & sAlphaStart & "' AND '" & sAlphaEnd & "'"
    End If

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 2

    If rs.RecordCount > 500 Then
        GetNumRcds = 500
    Else
        GetNumRcds = rs.RecordCount
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Function GetRaceName(lWhichRace)
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql="SELECT RaceName FROM RaceData WHERE RaceID = " & lWhichRace
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetRaceName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

Public Function GetAgeGrp(sMF, iAge, lThisRace)
    Dim sql_agegrp, rs_agegrp
    Dim iBegAge, iEndAge
    
    iBegAge = 0
    
    Set rs_agegrp = Server.CreateObject("ADODB.REcordset")
    sql_agegrp = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sMF & "' AND RaceID = " & lThisRace & " ORDER BY EndAge DESC"
    rs_agegrp.Open sql_agegrp, conn, 1, 2
    Do While Not rs_agegrp.EOF
        If CInt(iAge) <= CInt(rs_agegrp(0).Value) Then
            iEndAge = rs_agegrp(0).Value
        Else
            iBegAge = CInt(rs_agegrp(0).Value) + 1
            Exit Do
        End If
        rs_agegrp.MoveNext
    Loop
    rs_agegrp.Close
    Set rs_agegrp = Nothing

    If iBegAge = 0 Then
        GetAgeGrp = iEndAge & " and Under"
    Else
        If iEndAge = 110 Then
            GetAgeGrp = CInt(iBegAge) & " and Over"
        Else
            GetAgeGrp = CInt(iBegAge) & " - " & iEndAge
        End If
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Participant Data</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h3><%=sEventName%> Participant Data</h3>
			
			<form class="form-inline" name="which_event" method="post" action="part_data.asp?event_id=<%=lEventID%>">
			<label for="events">Events:</label>
			<select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Events, 2)%>
					<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%Else%>
						<option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			<input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
			</form>
			<br>
            <%If Not CLng(lEventID) = 0 Then%>
			    <!--#include file = "../../includes/event_nav.asp" -->
			    <!--#include file = "part_nav.asp" -->

                <div class="col-sm-5">
                    <%If UBound(Races, 2) > 1 Then%>
				        <form class="form-inline" name="get_race" method="post" action="part_data.asp?event_id=<%=lEventID%>">
					    <label for="races">Race:</label>
					    <select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
						    <option value="">View All</option>
						    <%For i = 0 to UBound(Races, 2) - 1%>
							    <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
								    <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
							    <%Else%>
								    <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
							    <%End If%>
						    <%Next%>
					    </select>
					    <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
					    <input type="submit" class="form-control" name="get_race" id="get_race" value="Get These">
				        </form>
			        <%End If%>
                    <h5 class="h5">Number of Registrants:&nbsp;<%=iTtlParts%></h5>
                </div>
                <div class="col-sm-7">
 			        <form class="form-inline" name="view_what" method="post" action="part_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
				    <label for="alpha_start">From</label>
				    <select class="form-control" name="alpha_start" id="alpha_start">
					    <%For i = 0 To 25%>
                            <%If sAlphaStart = AlphaArr(i) Then%>
                                <option value="<%=AlphaArr(i)%>" selected><%=AlphaArr(i)%></option>
                            <%Else%>
                                <option value="<%=AlphaArr(i)%>"><%=AlphaArr(i)%></option>
                            <%End If%>
                        <%Next%>
				    </select>
                    <label for="alpha_start">To</label>
				    <select class="form-control" name="alpha_end" id="alpha_end">
					    <%For i = 0 To 25%>
                            <%If sAlphaEnd = AlphaArr(i) Then%>
                                <option value="<%=AlphaArr(i)%>" selected><%=AlphaArr(i)%></option>
                            <%Else%>
                                <option value="<%=AlphaArr(i)%>"><%=AlphaArr(i)%></option>
                            <%End If%>
                        <%Next%>
				    </select>
                    Num Rcds:
                    <input type="text" class="form-control" name="num_rcds" id="num_rcds" value="<%=iNumRcds%>" size="4">
				    <input type="hidden" name="submit_alpha" id="submit_alpha" value="submit_alpha">
				    <input type="submit" class="form-control" name="get_alpha" id="get_alpha" value="Get These">
			        </form>
                </div>
                
			    <ul class="nav">
				    <%If sSortBy = "bib" Then%>
					    <li class="nav-item"><a class="nav-link" href="part_data.asp?race_id=<%=lRaceID%>&amp;sort_by=name&amp;event_id=<%=lEventID%>">Sort By Name</a></li>
				    <%Else%>
					    <li class="nav-item"><a class="nav-link" href="part_data.asp?race_id=<%=lRaceID%>&amp;sort_by=bib&amp;event_id=<%=lEventID%>">Sort By Bib</a></li>
				    <%End If%>
				    <li class="nav-item"><a class="nav-link" href="javascript:pop('/admin/participants/bib_assign.asp?event_id=<%=lEventID%>',400,200)">Assign Missing Bibs</a></li>
				    <li class="nav-item"><a class="nav-link" href="javascript:pop('print_regs.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>',1000,750)">Print</a></li>
				    <li class="nav-item"><a class="nav-link" href="dwnld_regs.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>" 
                        onclick="openThis(this.href,1024,768);return false;">Download</a></li>
				    <li class="nav-item"><a class="nav-link" href="part_data.asp?race_id=<%=lRaceID%>&amp;sort_by=<%=sSortBy%>&amp;event_id=<%=lEventID%>">Refresh Page</a></li>
			    </ul>
			
			    <form class="form" name="get_bibs" method="post" action="part_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
			    <table class="table table-striped table-condensed">
				    <tr>
					    <th>No</th>
					    <th>Edit</th>
					    <th>Race</th>
					    <th>First</th>
                        <th>Last</th>
					    <th>Bib</th>
					    <th>M/F</th>
					    <th>Age</th>
					    <th>City</th>
					    <th>St</th>
					    <th>DOB</th>
                        <th>Email</th>
				    </tr>

                    <%If UBound(PartArray, 2) > 0 Then%>
				        <%For j = 0 to UBound(PartArray, 2)%>
						    <tr>
							    <td>
								    <%=j+1%>)
							    </td>
							    <td>
								    <a href="javascript:pop('edit_part.asp?part_id=<%=PartArray(0, j)%>&amp;race_id=<%=PartArray(10, j)%>&amp;event_id=<%=lEventID%>',950,700)">
									    <img src="/graphics/edit.png" style="width:20px;height:17px;border:none;" alt="Edit">
								    </a>
							    </td>
							    <td><%=PartArray(10, j)%>-<%=GetRaceName(PartArray(10, j))%></td>
							    <%For i = 1 to 9%>
								    <%If i = 3 Then%>
									    <td style="width: 65px;">
                                            <input type="hidden" name="edit_<%=PartArray(0, j)%>" id="edit_<%=PartArray(0, j)%>" value="y">
										    <input type="text" class="form-control" name="bib_<%=PartArray(0, j)%>" id="bib_<%=PartArray(0, j)%>" 
                                                 style="width: 65px;" value="<%=PartArray(3, j)%>">
									    </td>	
								    <%Else%>
									    <td><%=PartArray(i, j)%></td>
								    <%End If%>
							    <%Next%>
						    </tr>
				        <%Next%>
                    <%End If%>
			    </table>
                <input type="hidden" name="submit_bibs" id="submit_bibs" value="submit_bibs">
			    <input type="submit" class="form-control" name="submit1" id="submit1" value="Assign Bibs" style="color:#d62002">
			    </form>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>