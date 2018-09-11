<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID
Dim sEventName, sRace, sEventRaces, sWhichTab, sInfoLink, sThisPage
Dim iNumRcds, iTtlParts
Dim sngDeposit
Dim dEventDate
Dim PartArray, Races(), Events()
Dim bChangesLocked

If Not Session("role") = "event_dir" Then Response.Redirect "/default.asp?sign_out=y"

sThisPage = "part_data.asp"
lEventID = Request.QueryString("event_id")
lRaceID = Request.QueryString("race_id")
sWhichTab = Request.QueryString("which_tab")

sngDeposit = 0
iTtlParts = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
    iNumRcds = GetNumRcds(lRaceID)
End If

i = 0
ReDim Events(1, 0)
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDirID = " & Session("my_id") & " ORDER By EventDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.eOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
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
If CStr(lRaceID) = vbNullString Then lRaceID = 0

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

bChangesLocked = False
If Date >= CDate(dEventDate) - 5 Then bChangesLocked = True

Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT p.ParticipantID FROM Participant p INNER JOIN PartRace rc ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces & ")"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iTtlParts = rs.RecordCount
rs.Close
Set rs=Nothing

If lRaceID = "0" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, rc.Bib, p.Gender, rc.Age, p.City, p.St, p.Phone, p.DOB, p.Email, rg.RaceID FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces & ") AND rc.RAceID = rg.RaceID ORDER BY p.LastName, p.FirstName" 
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        PartArray=rs.GetRows()		
    Else
        ReDim PartArray(11, 0)
    End If
    rs.Close
	Set rs=Nothing
Else
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, rc.Bib, p.Gender, rc.Age, p.City, p.St, p.Phone, p.DOB, p.Email, rg.RaceID FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID = " & lRaceID & " AND rc.RAceID = rg.RaceID ORDER BY p.LastName, p.FirstName"
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        PartArray=rs.GetRows()		
    Else
        ReDim PartArray(11, 0)
    End If
	rs.Close
	Set rs=Nothing

   	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
	rs.Open sql, conn, 1, 2
	sRace = Replace(rs(0).Value, "''", "'")
	rs.Close
	Set rs = Nothing
End If

Private Function GetNumRcds(lWhichRace)
    If CLng(lWhichRace) = 0 Then
		sql="SELECT p.LastName FROM Participant p INNER JOIN PartRace rc ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces & ")"
    Else
		sql="SELECT p.LastName FROM Participant p INNER JOIN PartRace rc ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID = " & lWhichRace
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        GetNumRcds = rs.RecordCount
    Else
        GetNumRcds = 0
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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Participant Data</title>
<!--#include file = "../../includes/js.asp" -->
<!--#include file = "event_css.asp" -->
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../includes/event_dir_menu.asp" -->
		<div class="col-md-10">
			<h3>GSE Edit/Manage Event Information: <span style="color:#000;"><%=sEventName%></span></h3>
			
            <!--#include file = "event_select.asp" -->	

            <div>
                <!--#include file = "event_dir_tabs.asp" -->

                <%If UBound(Races, 2) > 1 Then%>
				    <form name="get_race" method="post" action="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=<%=sWhichTab%>">
				    <div>	
					    <span style="font-weight:bold;">Select Race:</span>
					    <select name="races" id="races" onchange="this.form.get_race.click()">
                            <option value="">&nbsp;</option>
						    <%For i = 0 to UBound(Races, 2) - 1%>
							    <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
								    <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
							    <%Else%>
								    <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
							    <%End If%>
						    <%Next%>
					    </select>
					    <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
					    <input type="submit" name="get_race" id="get_race" value="Get Race Participants">
				    </div>
				    </form>
			    <%End If%>
				
			    <div style="float:left;width:225px;font-weight:bold;padding:5px 0 5px 10px;font-size:0.9em;">
				    Number of Registrants:&nbsp;<%=iTtlParts%>
			    </div>
			    <div style="margin-left:250px;text-align:right;width:550px;font-size:0.9em;padding:5px 0 5px 10px;">
				    <a href="javascript:pop('print_regs.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>',1000,750)">Print</a>
				    &nbsp;|&nbsp;
				    <a href="/admin/participants/dwnld_regs.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>" 
                    onclick="openThis(this.href,1024,768);return false;">Download</a>
				    <%If bChangesLocked = False Then%>
                        &nbsp;|&nbsp;
				        <a href="javascript:pop('/staff/enter_parts.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>',500,700)" >Enter Participants</a>
				        &nbsp;|&nbsp;
				        <a href="part_data.asp?race_id=<%=lRaceID%>&amp;event_id=<%=lEventID%>&amp;which_tab=<%=sWhichTab%>" >Refresh List</a>
                    <%End If%>
			    </div>

			    <table style="border-collapse:collapse;width:800px;font-size:0.8em;margin: 10px;">
				    <tr>
					    <th style="text-align:center;width:10px">No</th>
					    <th>Race</th>
					    <th>First</th>
                        <th>Last</th>
					    <th>Bib</th>
					    <th>M/F</th>
					    <th>Age</th>
					    <th>City</th>
					    <th>St</th>
					    <th>Phone</th>
					    <th>DOB</th>
                        <th>Email</th>
				    </tr>
                    <%If UBound(PartArray, 2) > 0 Then%>
				        <%For j = 0 to UBound(PartArray, 2)%>
					        <%If j/2 = j\2 Then%>
						        <tr>
							        <td class="alt">
								        <%=j+1%>)
							        </td>
							        <td class="alt" style="white-space:nowrap;"><%=PartArray(11, j)%>-<%=GetRaceName(PartArray(11, j))%></td>
							        <%For i = 1 to 10%>
								        <%If i = 3 Then%>
									        <td class="alt" style="white-space:nowrap;padding-right:2px;">
                                                <input type="hidden" name="edit_<%=PartArray(0, j)%>" id="edit_<%=PartArray(0, j)%>" value="y">
										        <input name="bib_<%=PartArray(0, j)%>" id="bib_<%=PartArray(0, j)%>" size="3" value="<%=PartArray(3, j)%>" style="text-align:center">
									        </td>	
								        <%Else%>
									        <td class="alt" style="white-space:nowrap"><%=PartArray(i, j)%></td>
								        <%End If%>
							        <%Next%>
						        </tr>
					        <%Else%>
						        <tr>
							        <td>
								        <%=j+1%>)
							        </td>
							        <td style="white-space:nowrap;"><%=PartArray(11, j)%>-<%=GetRaceName(PartArray(11, j))%></td>
							        <%For i = 1 to 10%>
								        <%If i = 3 Then%>
									        <td style="white-space:nowrap;padding-right:2px;">
                                                <input type="hidden" name="edit_<%=PartArray(0, j)%>" id="edit_<%=PartArray(0, j)%>" value="y">
										        <input name="bib_<%=PartArray(0, j)%>" id="bib_<%=PartArray(0, j)%>" size="3" value="<%=PartArray(3, j)%>" style="text-align:center">
									        </td>	
								        <%Else%>
									        <td style="white-space:nowrap"><%=PartArray(i, j)%></td>
								        <%End If%>
							        <%Next%>
						        </tr>
					        <%End If%>
				        <%Next%>
                    <%End If%>
			    </table>
            </div>
		</div>
	</div>
	<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>