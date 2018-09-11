<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lThisMeet, lRaceID
Dim sRaceName, sRaceDesc, sRaceTime, iRaceDist, sRaceUnits, sMeetName, sGender, sScoreMethod, iNumAllow, iNumScore, sComments, sStartType
Dim sTmAwds, sIndAwds
Dim iTotalParts
Dim cdoMessage, cdoConfig
Dim sMsg
Dim Races(), MeetArr()
Dim dMeetDate

Dim sMapLink, sMeetInfoSheet, sCourseMap
Dim dWhenShutdown

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")
lRaceID = Request.QueryString("race_id")

If CStr(lRaceID) = vbNullString Then lRaceID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim MeetArr(1, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDirID = " & Session("my_id") & " ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetArr(0, i) = rs(0).Value
	MeetArr(1, i) = rs(1).Value & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve MeetArr(1, i)
	rs.MoveNext
Loop

If UBound(MeetArr, 2) = 1 Then lThisMeet = MeetArr(0, 0)

If Request.Form.Item("submit_meet") = "submit_meet" Then 
    lThisMeet = Request.Form.Item("meets")
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
    sRaceName = Replace(Request.Form.Item("race_name"), "'", "''")
    sRaceDesc = Replace(Request.Form.Item("race_desc"), "'", "''")
    sRaceTime = Request.Form.Item("race_time")
	iRaceDist = Request.Form.Item("race_dist")
    sRaceUnits = Request.Form.Item("race_units")
    sGender = Request.Form.Item("gender")
    sScoreMethod = Request.Form.Item("score_meth")
	iNumAllow = Request.Form.Item("num_allow")
 	iNumScore = Request.Form.Item("num_score")
	sComments = Replace(Request.Form.Item("comments"), "'", "''")
    sTmAwds = Replace(Request.Form.Item("tm_awds"), "'", "''")
    sIndAwds = Replace(Request.Form.Item("ind_awds"), "'", "''")
    sStartType = Request.Form.Item("start_type")
    
    'add the race
    sql = "INSERT INTO Races(MeetsID, RaceName, RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, ScoreMethod, "
    sql = sql & "NumAllow, NumScore, Comments, TmAwds, IndAwds, StartType) VALUES (" & lThisMeet & ", '" & sRaceName 
    sql = sql & "', '" & sRaceDesc & "', '" & sRaceTime & "', " & iRaceDist & ", '" & sRaceUnits & "', '" & sGender & "', '" 
    sql = sql & sScoreMethod & "', " & iNumAllow & ", " & iNumScore & ", '" & sComments & "', '" & sTmAwds & "', '" 
    sql = sql & sIndAwds & "', '" & sStartType & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
    
    'get race id
    sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lThisMeet & " AND RaceName = '" & sRaceName 
    sql = sql & "' AND RaceDesc = '" & sRaceDesc & "' AND RaceTime = '" & sRaceTime & "' AND RaceDist = " 
    sql = sql & iRaceDist & " AND RaceUnits = '" & sRaceUnits & "' AND Gender = '" & sGender & "' AND ScoreMethod = '" 
    sql = sql & sScoreMethod & "' AND NumAllow = " & iNumAllow & " AND NumScore = " & iNumScore 
    sql = sql & " AND Comments = '" & sComments & "' AND TmAwds = '" & sTmAwds
    sql = sql & "' AND IndAwds = '" & sIndAwds & "' AND StartType = '" & sStartType & "'"
    Set rs = conn.Execute(sql)
    lRaceID = rs(0).Value
    Set rs = Nothing
    
    'insert bib range value
    sql = "INSERT INTO RaceDelay (RacesID) VALUES (" & lRaceID & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
		
	sMsg = vbCrLf
	sMsg = sMsg & "A new race has been added to " & sMeetName & vbCrLf & vbCrLf
		
	sMsg = sMsg & "DETAILS: " & vbCrLf
	sMsg = sMsg & "Race Name: " & sRaceName & vbCrLf
	sMsg = sMsg & "Race Description: " & sRaceDesc & vbCrLf
	sMsg = sMsg & "Race Time: " & sRaceTime & vbCrLf
	sMsg = sMsg & "Race Dist: " & iRaceDist &" " & sRaceUnits & vbCrLf
	sMsg = sMsg & "Gender: " & sGender & vbCrLf
	sMsg = sMsg & "Start Type: " & sStartType & vbCrLf
		
%>
<!--#include file = "../../../includes/cdo_connect.asp" -->
<%

	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.com"
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "GSE New CC Race"
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
End If

If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

If Not CLng(lThisMeet) = 0 Then
    sql = "SELECT MeetName, MeetDate, WhenShutdown FROM Meets WHERE MeetsID = " & lThisMeet
    Set rs = conn.Execute(sql)
    sMeetName = Replace(rs(0).Value, "''", "'")
    dMeetDate = rs(1).Value
    dWhenShutdown = rs(2).Value
    Set rs = Nothing

	'get maplink
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMapLink = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT InfoSheet FROM MeetInfo WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMeetInfoSheet = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get course map
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Map FROM CourseMap WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sCourseMap = rs(0).Value
	rs.Close
	Set rs = Nothing

    iTotalParts = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RosterID FROM IndRslts WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then iTotalParts = rs.RecordCount
    rs.Close
    Set rs = Nothing

    'get race information
    i = 0
    ReDim Races(5, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RacesID, RaceName, RaceTime, RaceDist, RaceUnits FROM Races WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    For j = 0 to 4
		    Races(j, i) = rs(j).Value
	    Next
	    Races(5, i) = FieldSize(rs(0).Value)
	    i = i + 1
	    ReDim Preserve Races(5,i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function FieldSize(lThisRaceID)
	FieldSize = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRaceID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then FieldSize = rs2.RecordCount
	rs2.Close
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE  Edit CC Races</title>
<!--#include file = "../../../includes/js.asp" -->

<script>
function chkFields(){
	if (document.add_race.race_name.value==''){
		alert('You must supply a race name!');
		return false;
	}
	else
		if (document.add_race.race_desc.value==''){
			alert('You must supply a race description!');
			return false;
		}
	else
		if (document.add_race.race_time.value==''){
			alert('You must supply a race time!');
			return false;
		}
	else
		if (document.add_race.race_dist.value==''){
			alert('You must supply a race distance!');
			return false;
		}
	else
		if (document.add_race.race_units.value==''){
			alert('You must supply units for this race!');
			return false;
		}
	else
		if (document.add_race.score_meth.value==''){
			alert('You must supply a scoring method for this race!');
			return false;
		}
	else
		if (document.add_race.num_allow.value==''){
			alert('You must submit the number of runners allowed per team (select 0 for unlimited)!');
			return false;
		}
	else
		if (document.add_race.num_score.value==''){
			alert('You must submit the number of runners that score for each team (select 0 if the race is not scored)!');
			return false;
		}
	else
		return true;
}
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	<!--#include file = "../../../includes/meet_dir_menu.asp" -->

	<h4 class="h4">CC/Nordic Meet Director: Edit Teams</h4>

	<form name="get_meets" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>">
	<div>
		<span style="font-weight:bold;">Select Meet:</span>
		<select name="meets" id="meets" onchange="this.form.submit1.click();">
			<option value="">&nbsp;</option>
			<%For i = 0 to UBound(MeetArr, 2) - 1%>
				<%If CLng(lThisMeet) = CLng(MeetArr(0, i)) Then%>
					<option value="<%=MeetArr(0, i)%>" selected><%=MeetArr(1, i)%></option>
				<%Else%>
					<option value="<%=MeetArr(0, i)%>"><%=MeetArr(1, i)%></option>
				<%End If%>
			<%Next%>
		</select>
		<input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
		<input type="submit" name="submit1" id="submit1" value="Get This">
	</div>
	</form>
			
    <%If Not CLng(lThisMeet) = 0 Then%>
		<!--#include file = "../meet_dir_nav.asp" -->
			
        <table>
            <tr>
                <td valign="top">
			        <h4 class="h4">Existing Races</h4>
				
			        <table>
				        <tr>
					        <th style="text-align:left;">Race:</th>
					        <th style="text-align:left;">Time</th>
					        <th style="text-align:left;">Dist</th>
					        <th style="text-align:left;">Entries</th>
				        </tr>
				        <%For i = 0 to UBound(Races, 2) - 1%>
                            <%If i mod 2 = 0 Then%>
						        <tr>
							        <td class="alt"><a href="javascript:pop('/ccmeet_admin/manage_meet/edit_this_race.asp?race_id=<%=Races(0, i)%>',800,300)"><%=Races(1, i)%></a></td>
							        <td class="alt"><%=Races(2, i)%></td>
							        <td class="alt"><%=Races(3, i)%> <%=Races(4, i)%></td>
							        <td class="alt" style="text-align:right;"><%=Races(5, i)%></td>
						        </tr>                        
                            <%Else%>
						        <tr>
							        <td><a href="javascript:pop('/ccmeet_admin/manage_meet/edit_this_race.asp?race_id=<%=Races(0, i)%>',800,300)"><%=Races(1, i)%></a></td>
							        <td><%=Races(2, i)%></td>
							        <td><%=Races(3, i)%> <%=Races(4, i)%></td>
							        <td style="text-align:right;"><%=Races(5, i)%></td>
						        </tr>
                            <%End If%>
				        <%Next%>
				        <tr>
					        <th style="text-align:right;" colspan="4">
						        Total Entries:&nbsp;<%=iTotalParts%>
					        </th>
				        </tr>
			        </table>
                </td>
                <td valign="top">
				    <h4 class="h4">Add Race</h4>
				
				    <form name="add_race" method="post" action="add_race.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkFields()">
				    <table>
					    <tr>
						    <th>Race Name:</th>
						    <td><input type="text" name="race_name" id="race_name" size="5" maxlength="5"></td>
						    <th>Description:</th>
						    <td><input type="text" name="race_desc" id="race_desc" size="20" maxlength="50"></td>
					    </tr>
					    <tr>
						    <th>Race Time:</th>
						    <td><input type="text" name="race_time" id="race_time" size="8" maxlength="12"></td>
						    <th>Distance:</th>
						    <td>
							    <input type="text" name="race_dist" id="race_dist" size="5" maxlength="12">
							    <select name="race_units" id="race_units" tabindex="5">
								    <option value="">&nbsp;</option>
								    <option value="miles">miles</option>
								    <option value="kms">km</option>
								    <option value="yds">yds</option>
								    <option value="meters">meters</option>
							    </select>
						    </td>
					    </tr>
					    <tr>
						    <th>Gender:</th>
						    <td>
							    <select name="gender" id="gender">
								    <option value="">&nbsp;</option>
								    <option value="Male">Male</option>
								    <option value="Female">Female</option>
								    <option value="Open">Open</option>
							    </select>
						    </td>
						    <th>Score By:</th>
						    <td>
							    <select name="score_meth" id="score_meth">
								    <option value="">&nbsp;</option>
								    <option value="Place">Place</option>
								    <option value="Time">Time</option>
								    <option value="None">None</option>
							    </select>
						    </td>
					    </tr>
					    <tr>
						    <th>Start Type:</th>
						    <td colspan="3">
							    <select name="start_type" id="start_type">
								    <option value="Mass">Mass</option>
								    <option value="Interval">Interval</option>
								    <option value="Wave">Wave</option>
							    </select>
						    </td>
					    </tr>
					    <tr>
						    <th valign="top">Team Awds:</th>
						    <td><textarea name="tm_awds" id="tm_awds" rows="2" cols="35" style="font-size:1.1em;"></textarea></td>
						    <th valign="top">Ind Awds:</th>
						    <td><textarea name="ind_awds" id="ind_awds" rows="2" cols="35" style="font-size:1.1em;"></textarea></td>
					    </tr>
					    <tr>
						    <th>Parts/Team:</th>
						    <td>
							    <select name="num_allow" id="num_allow">
								    <option value="0">Unlimited</option>
								    <%For i = 1 To 100%>
									    <option value="<%=i%>"><%=i%></option>
								    <%Next%>
							    </select>
						    </td>
						    <th>Num Score:</th>
						    <td>
							    <select name="num_score" id="num_score">
								    <%For i = 0 To 100%>
									    <option value="<%=i%>"><%=i%></option>
								    <%Next%>
							    </select>
						    </td>
					    </tr>
					    <tr>
						    <th>Comments:</th>
						    <td colspan="3"><textarea name="comments" id="comments"  rows="2" cols="35" style="font-size:1.1em;"></textarea></td>
					    </tr>
					    <tr>
						    <td style="text-align:center" colspan="4">
							    <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
							    <input type="submit" name="submit" id="submit" tabindex="13" value="Submit This Race">
						    </td>
					    </tr>
				    </table>
				    </form>
                </td>
            </tr>
        </table>
	<%End If%>
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
