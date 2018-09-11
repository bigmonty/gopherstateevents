<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lRaceID, lEventID
Dim iAgeGrpAwds
Dim sEventName, sStartTime, sRaceName, sErrMsg, sWhichTab, sInfoLink, sThisPage, sTeamScore
Dim sngDeposit
Dim InfoArray(14), MaleArray(), FemaleArray(), StartType(2), Delete(), Events()
Dim bBibsOverlap, bFound, bLastGrpExists, bChangesLocked
Dim dEventDate

If Not Session("role") = "event_dir" Then Response.Redirect "/default.asp?sign_out=y"

sThisPage = "race_data.asp"
lEventID = Request.QueryString("event_id")
lRaceID = Request.QueryString("race_id")
sWhichTab = Request.QueryString("which_tab")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_age_grps") = "submit_age_grps" Then
    i = 0
    ReDim Delete(0)

	'write male back to db
    iAgeGrpAwds = 0
    bLastGrpExists = False
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AgeGroupsID, EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'm' ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF	
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
		    If rs(1).Value = "110" Then 
                bLastGrpExists = True
			
			    If IsNumeric(Request.Form.Item("m_awds_" & rs(0).Value)) Then
				    rs(2).Value = Request.Form.Item("m_awds_" & rs(0).Value)
			    Else
				    sErrMsg = "All award values must be numeric.  Some work was not done."
			    End If
            Else
			    If IsNumeric(Request.Form.Item("m_end_age_" & rs(0).Value)) Then
				    rs(1).Value = Request.Form.Item("m_end_age_" & rs(0).Value)	
			    Else
				    sErrMsg = "All ending ages must be numeric.  Some work was not done."
			    End If
			
			    If IsNumeric(Request.Form.Item("m_awds_" & rs(0).Value)) Then
                    iAgeGrpAwds = Request.Form.Item("m_awds_" & rs(0).Value)
				    rs(2).Value = iAgeGrpAwds
			    Else
				    sErrMsg = "All award values must be numeric.  Some work was not done."
			    End If
		    End If
		    rs.Update
        End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

    If bLastGrpExists = False Then
		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'm', 110, " & iAgeGrpAwds & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    End If
	
	If CInt(Request.Form.Item("new_m_end_age")) > 0 Then
		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'm', "
		sql = sql & Request.Form.Item("new_m_end_age") & ", " & Request.Form.Item("new_m_awds") & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	End If

	'write female back to db
    iAgeGrpAwds = 0
    bLastGrpExists = False
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AgeGroupsID, EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'f' ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF		
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
		    If rs(1).Value = "110" Then 
                bLastGrpExists = True
			
			    If IsNumeric(Request.Form.Item("f_awds_" & rs(0).Value)) Then
				    rs(2).Value = Request.Form.Item("f_awds_" & rs(0).Value)
			    Else
				    sErrMsg = "All award values must be numeric.  Some work was not done."
			    End If
            Else
			    If IsNumeric(Request.Form.Item("f_end_age_" & rs(0).Value)) Then
				    rs(1).Value = Request.Form.Item("f_end_age_" & rs(0).Value)	
			    Else
				    sErrMsg = "All ending ages must be numeric.  Some work was not done."
			    End If
			
			    If IsNumeric(Request.Form.Item("f_awds_" & rs(0).Value)) Then
                    iAgeGrpAwds = Request.Form.Item("f_awds_" & rs(0).Value)
				    rs(2).Value = iAgeGrpAwds
			    Else
				    sErrMsg = "All award values must be numeric.  Some work was not done."
			    End If
		    End If
		    rs.Update
        End IF
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

    If bLastGrpExists = False Then
		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'f', 110, " & iAgeGrpAwds & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    End If
	
	If CInt(Request.Form.Item("new_f_end_age")) > 0 Then
		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'f', "
		sql = sql & Request.Form.Item("new_f_end_age") & ", " & Request.Form.Item("new_f_awds") & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	End If

    For i = 0 To UBound(Delete) - 1
		sql = "DELETE FROM AgeGroups WHERE AgeGroupsID = " & Delete(i)
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    Next
ElseIf Request.Form.Item("submit_race_info") = "submit_race_info" Then
	'male sure the time field is an acceptable format
	sStartTime = Request.Form.Item("start_time")
	For i = 1 to Len(sStartTime)
		If Mid(sStartTime, i, 1) = ":" Then
			If Not (CInt(i) = 2 Or CInt(i) = 3) Then
				sErrMsg = "Please use 'hh:mm' format for this race's start time."
				Exit For
			End If
		Else
			If Not IsNumeric(Mid(sStartTime, i, 1)) Then
				sErrMsg = "Please use 'hh:mm' format for this race's start time."
				Exit For
			End If
		End If
	Next
	
	'if all is good then...
	If sErrMsg = vbNullString Then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RaceName, Dist, StartTime, Certified, StartType, MAwds, FAwds, OnlineRegLink, "
        sql = sql & "AllowDuplAwds, EntryFeePre, EntryFee, ChipStart, NumSplits, StartToFinish FROM RaceData WHERE RaceID = " & lRaceID
		rs.Open sql, conn, 1, 2
		rs(0).Value = Replace(Request.Form.Item("race_name"), "'", "''")
		rs(1).Value = Request.Form.Item("dist")
		rs(2).Value = Request.Form.Item("start_time") & Request.Form.Item("am_pm")
		rs(3).Value = Request.Form.Item("certif")
		rs(4).Value = Request.Form.Item("start_type")
		rs(5).Value = Request.Form.Item("mawds")
		rs(6).Value = Request.Form.Item("fawds")
    	rs(7).Value = Request.Form.Item("online_reg_link")
    	rs(8).Value = Request.Form.Item("allow_dupl_awds")
        rs(9).Value = Request.Form.Item("entry_fee_pre")
        rs(10).Value = Request.Form.Item("entry_fee")
        rs(11).Value = Request.Form.Item("chip_start")
        rs(12).Value = Request.Form.Item("num_splits")
        rs(13).Value = Request.form.Item("start_to_finish")
		rs.Update
		rs.Close
		Set rs = Nothing

        'process team scoring
        sTeamScore = Request.Form.Item("team_score")

        If sTeamScore = "n" Then
            sql = "DELETE FROM TeamScoring WHERE RaceID = " & lRaceID
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        Else
            bFound = False
            sql = "SELECT RaceID FROM TeamScoring WHERE RaceID = " & lRaceID
            Set rs = conn.Execute(sql)
            If Not rs.EOF = rs.BOF Then bFound = True
            Set rs = Nothing

            If bFound = False Then
                sql = "INSERT INTO TeamScoring (RaceID) VALUES (" & lRaceID & ")"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            End If
        End If
	End If
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
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

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Website, Deposit FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    If rs(0).Value & "" = "" Then
        sInfoLink = "http://www.gopherstateevents.com/events/raceware_events.asp?event_id=" & lEventID
    Else
        sInfoLink = rs(0).Value
    End If

    sngDeposit = rs(1).Value
End If
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

bChangesLocked = False
If Date >= CDate(dEventDate) - 5 Then bChangesLocked = True

i = 0
ReDim RaceArray(1, 0)
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RaceArray(0, i) = rs(0).Value
	RaceArray(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve RaceArray(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If CStr(lRaceID) = vbNullString Then lRaceID = RaceArray(0, 0)

'check for last end age = 110
Dim iEndAge
iEndAge = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EndAge FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'm' ORDER BY EndAge DESC"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    iEndAge = rs(0).Value
End If
rs.Close
Set rs = Nothing

If CInt(iEndAge) < 110 Then
   	sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'm', 110, 0)"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

iEndAge = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EndAge FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'f' ORDER BY EndAge DESC"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    iEndAge = rs(0).Value
End If
rs.Close
Set rs = Nothing

If CInt(iEndAge) < 110 Then
   	sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'f', 110, 0)"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

sql = "SELECT RaceName, Dist, StartTime, Certified, StartType, MAwds, FAwds, OnlineRegLink, AllowDuplAwds, "
sql = sql & "EntryFeePre, EntryFee, ChipStart, NumSplits, StartToFinish FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
InfoArray(0) = rs(0).Value
InfoArray(1) = rs(1).Value
	
'split the time field
InfoArray(2) = Left(rs(2).Value, Len(rs(2).Value) - 2)
InfoArray(3) = Right(rs(2).Value, 2)
	
InfoArray(4) = rs(3).Value
InfoArray(5) = rs(4).Value
InfoArray(6) = rs(5).Value
InfoArray(7) = rs(6).Value
InfoArray(8) = rs(7).Value
InfoArray(9) = rs(8).Value
InfoArray(10) = rs(9).Value
InfoArray(11) = rs(10).Value
InfoArray(12) = rs(11).Value
InfoArray(13) = rs(12).Value
InfoArray(14) = rs(13).Value
Set rs = Nothing

If InfoArray(13) & "" = "" Then InfoArray(13) = 0
If InfoArray(14) & "" = "" Then InfoArray(14) = 0

sTeamScore = "n"
Set rs = SErver.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID FROM TeamScoring WHERE RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then sTeamScore = "y"
rs.Close
Set rs = Nothing

'get male age group array
i = 0
ReDim MaleArray(2, 0)
Set rs = SErver.CreateObject("ADODB.Recordset")
sql = "SELECT AgeGroupsID, EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lRaceID
sql = sql & " AND Gender = 'm' ORDER BY EndAge"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	For j = 0 to 2
		MaleArray(j, i) = rs(j).Value
	Next
		
	i = i + 1
	ReDim Preserve MaleArray(2, i)
		
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get female age group array
i = 0
ReDim FemaleArray(2, 0)
Set rs = SErver.CreateObject("ADODB.Recordset")
sql = "SELECT AgeGroupsID, EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lRaceID
sql = sql & " AND Gender = 'f' ORDER BY EndAge"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	For j = 0 to 2
		FemaleArray(j, i) = rs(j).Value
	Next
		
	i = i + 1
	ReDim Preserve FemaleArray(2, i)
		
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

StartType(0) = "mass"
StartType(1) = "wave"
StartType(2) = "interval"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Race Data</title>
<!--#include file = "../../includes/js.asp" -->
<!--#include file = "event_css.asp" -->

<script>
function checkFields() {
 	if (document.update_info.dist.value == '' || 
	 	document.update_info.race_name.value == '' || 
	 	document.update_info.mawds.value == '' || 
	 	document.update_info.fawds.value == '' || 
	 	document.update_info.start_time.value == '')
		{
  		alert('All fields are required!  If a field calls for a numeric value you may enter 0.');
  		return false
  		}
	else
		if (isNaN(document.update_info.mawds.value) ||
		   isNaN(document.update_info.fawds.value))
    		{
			alert('All distance and awards fields must be numeric values');
			return false
			} 	
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../includes/event_dir_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">GSE Edit/Manage Event Information: <span style="color:#000;"><%=sEventName%></span></h3>
			
            <!--#include file = "event_select.asp" -->

            <div>
                <!--#include file = "event_dir_tabs.asp" -->
                                        
			    <%If UBound(RaceArray, 2) > 1 Then%>
				    <form name="get_races" method="post" action="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=<%=sWhichTab%>">
				    <div style="margin-left:10px;">	
					    <span style="font-weight:bold;">Select Race:</span>
					    <select name="races" id="races" onchange="this.form.get_race.click()">
						    <%For i = 0 to UBound(RaceArray, 2) - 1%>
							    <%If CLng(lRaceID) = CLng(RaceArray(0, i)) Then%>
								    <option value="<%=RaceArray(0, i)%>" selected><%=RaceArray(1, i)%></option>
							    <%Else%>
								    <option value="<%=RaceArray(0, i)%>"><%=RaceArray(1, i)%></option>
							    <%End If%>
						    <%Next%>
					    </select>
					    <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
					    <input type="submit" name="get_race" id="get_race" value="Get Race Info">
				    </div>
				    </form>
			    <%End If%>
			
				<%If Not sErrMsg = vbNullString Then%>
					<p style="margin-left:10px;"><%=sErrMsg%></p>
				<%End If%>
				
				<form name="update_info" method="post" action="race_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;which_tab=<%=sWhichTab%>" 
                    onsubmit="return checkFields()">
				<table style="margin-left:10px;">
					<tr>	
						<th class="required" valign="top">Race Name:</th>
						<td class="required" valign="top"><input name="race_name" id="race_name" maxlength="35" value="<%=InfoArray(0)%>"></td>
                        <td class="required" valign="top">
                            This is what will show up on the results pages.  It may be the same as the distance (below) or it can be something like 
                            "5K Run-Walk".  Please do not use the "/" character in the race name.
                        </td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Distance:</th>
						<td class="required" valign="top"><input name="dist" id="dist" maxlength="6" size="3" value="<%=InfoArray(1)%>"></td>
                        <td class="required" valign="top">
                            The formatting for this MUST be "5 km" (number, space, and then km or Miles).  For races like a marathon, half-marathon,
                            and some specialty events, we will manage this on our end
                        </td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Certified:</th>
						<td class="required" valign="top">
							<select name="certif" id="certif">
								<%If InfoArray(4) = "y" Then%>
									<option value="y" selected>Yes</option>
									<option value="n">No</option>
								<%Else%>
									<option value="y">Yes</option>
									<option value="n" selected>No</option>
								<%End If%>
							</select>
						</td>
                        <td class="required" valign="top">&nbsp;</td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Start Type:</th>
						<td class="required" valign="top">
							<select name="start_type" id="start_type">
								<%If InfoArray(5) = "mass" Then%>
									<option value="mass" selected>Mass</option>
									<option value="wave">Wave</option>
									<option value="interval">Interval</option>
								<%ElseIf InfoArray(5) = "wave" Then%>
									<option value="mass">Mass</option>
									<option value="wave" selected>Wave</option>
									<option value="interval">Interval</option>
								<%ElseIf InfoArray(5) = "interval" Then%>
									<option value="mass">Mass</option>
									<option value="wave">Wave</option>
									<option value="interval" selected>Interval</option>
								<%End If%>
							</select>
						</td>
                        <td class="required" valign="top">
                            This is almost always a mass start.  Typically multi-sport, Nordic Ski, and some specialty events have wave or interval
                            starts.
                        </td>
					</tr>
					<tr>	
						<th class="required" valign="top">Start Time:</th>
						<td class="required" valign="top">
							<input name="start_time" id="start_time" size="3" maxlength="5" value="<%=InfoArray(2)%>" onchange="return chkStr(this)">
							<select name="am_pm" id="am_pm">
								<%If InfoArray(3) = "am" Then%>
									<option value="am">am</option>
									<option value="pm">pm</option>
								<%Else%>
									<option value="am">am</option>
									<option value="pm" selected>pm</option>
								<%End If%>
							</select>
						</td>
                        <td class="required" valign="top">&nbsp;</td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Chip Start:</th>
						<td class="required" valign="top">
							<select name="chip_start" id="chip_start">
								<%If InfoArray(12) = "y" Then%>
									<option value="y" selected>Yes</option>
									<option value="n">No</option>
								<%Else%>
									<option value="y">Yes</option>
									<option value="n" selected>No</option>
								<%End If%>
							</select>
						</td>
                        <td class="required" valign="top">
                            All races with over 200 particpants will have a chip start.  Races with under 200 participants will have a chip start free of
                            charge if the starting line and the finish line are the same place.  In other cases a $150 fee will be assessed for a chip 
                            start.
                        </td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Number of Splits:</th>
						<td class="required" valign="top">
							<select name="num_splits" id="num_splits">
								<%For i = 0 To 5%>
                                    <%If CInt(InfoArray(13)) = CInt(i) Then%>
									    <option value="<%=i%>" selected><%=i%></option>
								    <%Else%>
									    <option value="<%=i%>"><%=i%></option>
								    <%End If%>
                                <%Next%>
							</select>
						</td>
                        <td class="required" valign="top">
                            If your event would like splits please indicate the number of splits you would like.  We will contact you for the location
                            of those splits.  Please note that there is a minimum fee of $25 for races that have a split at the existing start or finish
                            line and a $150 fee for each additional timing box required.
                        </td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Duplicate Awards Allowed:</th>
						<td class="required" valign="top">
							<select name="allow_dupl_awds" id="allow_dupl_awds">
								<%If InfoArray(9) = "y" Then%>
									<option value="y" selected>Yes</option>
									<option value="n">No</option>
								<%Else%>
									<option value="y">Yes</option>
									<option value="n" selected>No</option>
								<%End If%>
							</select>
						</td>
                        <td class="required" valign="top">This determines whether or not a person can earn an Open award AND and Age Group Award.</td>
                    </tr>
                    <tr>
 						<th class="required" valign="top">Distance between<br>Start/Finish Lines:</th>
						<td class="required" valign="top"><input name="start_to_finish" id="start_to_finish" size="2"value="<%=InfoArray(14)%>"></td>
                        <td class="required" valign="top">
                            Let us know approximately, in meters, how far apart the start and finish line are.  This helps
                            us plan our equipment availability.
                        </td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Team Scoring:</th>
						<td class="required" valign="top">
							<select name="team_score" id="team_score">
								<%If sTeamScore = "y" Then%>
									<option value="y" selected>Yes</option>
									<option value="n">No</option>
								<%Else%>
									<option value="y">Yes</option>
									<option value="n" selected>No</option>
								<%End If%>
							</select>
						</td>
                        <td class="required" valign="top">
                            If you would like to add a team scoring component to this race select "Yes".  We will get back
                            to you regarding the parameters for scoring.
                        </td>
                    </tr>
                    <tr>
 						<th class="required" valign="top">Male Open Awds:</th>
						<td class="required" valign="top"><input name="mawds" id="mawds" size="2"value="<%=InfoArray(6)%>"></td>
                        <td class="required" valign="top">&nbsp;</td>
                    </tr>
                    <tr>
						<th class="required" valign="top">Female Open Awds:</th>
						<td class="required" valign="top"><input name="fawds" id="fawds" size="2" value="<%=InfoArray(7)%>"></td>
                        <td class="required" valign="top">&nbsp;</td>
                    </tr>
                    <tr>
 						<th valign="top">Pre-Registration Fee:</th>
						<td valign="top"><input name="entry_fee_pre" id="entry_fee_pre" size="2"value="<%=InfoArray(10)%>"></td>
                        <td valign="top">This value can be changed periodically if your entry fee changes at certain intervals.</td>
                    </tr>
                    <tr>
						<th valign="top">Race Day Fee:</th>
						<td valign="top"><input name="entry_fee" id="entry_fee" size="2" value="<%=InfoArray(11)%>"></td>
                        <td valign="top">&nbsp;</td>
                    </tr>
                    <tr>
						<th valign="top">Registration Link:</th>
						<td valign="top" colspan="2"><input type="text" name="online_reg_link" id="online_reg_link" value="<%=InfoArray(8)%>" size="100"></td>
					</tr>
					<tr>
						<td colspan="3">
							<input type="hidden" name="submit_race_info" id="submit_race_info" value="submit_race_info">
							<%If bChangesLocked = False Then%>
                                <input type="submit" name="submit1" id="submit1" value="Make Changes">
                            <%Else%>
                                <input type="submit" name="submit1" id="submit1" value="Make Changes" disabled>
                            <%End If%>
						</td>
					</tr>
				</table>
				</form>
				
				<form name="update_age_grps" method="post" action="race_data.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;which_tab=<%=sWhichTab%>">
				<table style="background-color:#ececec;width: 600px;margin-top: 20px;">
					<tr>
						<td>
							<table style="background-color:#fff;font-size:1.2em;width: 300px;">
								<tr>
									<th style="border-bottom:1px solid #ececd8;" colspan="2">Male Age Groups</th>
								</tr>
								<tr>
									<td style="text-align:center;">End Age</td>
									<td style="text-align:center;">Awds</td>
                                    <td style="text-align:center;">Delete?</td>
								</tr>
								<%If UBound(MaleArray, 2) > 1 Then%>
									<%For i = 0 to UBound(MaleArray, 2) - 1%>
										<tr>
											<td style="text-align:center;">
												<% If MaleArray(1, i) = "110" Then%>
													<input type="text" name="m_end_age_<%=MaleArray(0, i)%>" 
															    id="m_end_age_<%=MaleArray(0, i)%>" 
															    value="<%=CInt(MaleArray(1, i - 1)) + 1%>  And Over" 
															    size="10" style="text-align:right" disabled>
												<%Else%>
													<input type="text" name="m_end_age_<%=MaleArray(0, i)%>" 
															    id="m_end_age_<%=MaleArray(0, i)%>" size="1"
																maxlength="2" value="<%=MaleArray(1, i)%>">
												<%End If%>
											</td>
											<td style="text-align:center;">
												<input type="text" name="m_awds_<%=MaleArray(0, i)%>" id="m_awds_<%=MaleArray(0, i)%>" size="1"
															maxlength="2" value="<%=MaleArray(2, i)%>">
											</td>
											<td style="text-align:center;">
												<input type="checkbox" name="delete_<%=MaleArray(0, i)%>" id="delete_<%=MaleArray(0, i)%>">
											</td>
										</tr>
									<%Next%>
                                <%Else%>
									<tr>
										<td style="text-align:center;">
											<input type="text" name="m_end_age_<%=MaleArray(0, 0)%>" id="m_end_age_<%=MaleArray(0, 0)%>" 
                                                size="1" maxlength="2" value="<%=MaleArray(1, 0)%>">
										</td>
										<td style="text-align:center;">
											<input type="text" name="m_awds_<%=MaleArray(0, 0)%>" id="m_awds_<%=MaleArray(0, 0)%>" size="1"
														maxlength="2" value="<%=MaleArray(2,0)%>">
										</td>
										<td style="text-align:center;">
											<input type="checkbox" name="delete_<%=MaleArray(0, 0)%>" id="delete_<%=MaleArray(0, 0)%>">
										</td>
									</tr>
								<%End If%>
								<tr><th style="text-align:left;" colspan="3">Add Male Age Group:</th></tr>
								<tr>
									<td style="text-align:center;">
										<select name="new_m_end_age" id="new_m_end_age">
											<%For j = 0 To 110%>
												<option value="<%=j%>"><%=j%></option>
											<%Next%>
										</select>
									</td>
									<td style="text-align:center;" colspan="2">
										<select name="new_m_awds" id="new_m_awds">
											<%For j = 0 To 25%>
												<option value="<%=j%>"><%=j%></option>
											<%Next%>
										</select>
									</td>
								</tr>
							</table>
						</td>
						<td>
							<table style="background-color:#fff;font-size:1.2em;width: 300px;">
								<tr>
									<th style="border-bottom:1px solid #ececd8;" colspan="2">Female Age Groups</th>
								</tr>
								<tr>
									<td style="text-align:center;">End Age</td>
									<td style="text-align:center;">Awds</td>
                                    <td style="text-align:center;">Delete?</td>
								</tr>
								<%If UBound(FemaleArray, 2) > 1 Then%>
									<%For i = 0 to UBound(FemaleArray, 2) - 1%>
										<tr>
											<td style="text-align:center;">
												<%If FemaleArray(1, i) = "110" Then%>
													<input type="text" name="f_end_age_<%=FemaleArray(0, i)%>" 
															    id="f_end_age_<%=FemaleArray(0, i)%>" 
															    value="<%=CInt(FemaleArray(1, i - 1)) + 1%>  And Over" 
															    size="10" style="text-align:right" disabled>
												<%Else%>
													<input type="text" name="f_end_age_<%=FemaleArray(0, i)%>" 
															    id="f_end_age_<%=FemaleArray(0, i)%>" size="1"
																maxlength="2" value="<%=FemaleArray(1, i)%>">
												<%End If%>
											</td>
											<td style="text-align:center;">
												<input type="text" name="f_awds_<%=FemaleArray(0, i)%>" id="f_awds_<%=FemaleArray(0, i)%>" size="1"
															maxlength="2" value="<%=FemaleArray(2, i)%>">
											</td>
											<td style="text-align:center;">
												<input type="checkbox" name="delete_<%=FemaleArray(0, i)%>" id="delete_<%=FemaleArray(0, i)%>">
											</td>
										</tr>
									<%Next%>
                                <%Else%>
									<tr>
										<td style="text-align:center;">
											<input type="text" name="f_end_age_<%=FemaleArray(0, 0)%>" id="f_end_age_<%=FemaleArray(0, 0)%>" 
                                                size="1" maxlength="2" value="<%=FemaleArray(1, 0)%>">
										</td>
										<td style="text-align:center;">
											<input type="text" name="f_awds_<%=FemaleArray(0, 0)%>" id="f_awds_<%=FemaleArray(0, 0)%>" size="1"
														maxlength="2" value="<%=FemaleArray(2, 0)%>">
										</td>
										<td style="text-align:center;">
											<input type="checkbox" name="delete_<%=FemaleArray(0, 0)%>" id="delete_<%=FemaleArray(0, 0)%>">
										</td>
									</tr>
								<%End If%>
								<tr>
									<th style="text-align:left;" colspan="3">Add Female Age Group:</th>
								</tr>
								<tr>
									<td style="text-align:center;">
										<select name="new_f_end_age" id="new_f_end_age">
											<%For j = 0 To 110%>
												<option value="<%=j%>"><%=j%></option>
											<%Next%>
										</select>
									</td>
									<td style="text-align:center;" colspan="2">
										<select name="new_f_awds" id="new_f_awds">
											<%For j = 0 To 25%>
												<option value="<%=j%>"><%=j%></option>
											<%Next%>
										</select>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td style="text-align:center;" colspan="4">
							<input type="hidden" name="submit_age_grps" id="submit_age_grps" value="submit_age_grps">
							<%If bChangesLocked = False Then%>
                                <input type="submit" name="submit2" id="submit2" value="Make Changes">
                            <%Else%>
                                <input type="submit" name="submit2" id="submit2" value="Make Changes" disabled>
                            <%End If%>
						</td>
					</tr>
				</table>
				</form>
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