<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisType, lRaceType, lRaceID, lEventID, lCloneEvent, lCloneRace
Dim EventTypes(), Events(), Races(), StartType(2), TShirtArray(8), MaleArray(), FemaleArray(), InfoArray(13)
Dim sDist, sEntryFeePre, sEntryFee, sStartTime, sCertified, sRaceName, sRaceDist, sEventName, sStartType
Dim sSmall, sMedium, sLarge, sXLarge, sXXLarge, sShort, sLong, sChooseNone, sAllowDuplAwds
Dim iMAwds, iFAwds, iNumAwds, iEndAge
Dim dEventDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lCloneEvent = Request.QueryString("clone_event")
lCloneRace = Request.QueryString("clone_race")
lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
	
i = 0
ReDim Events(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.eOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
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

'get event types
i = 0
ReDim EventTypes(1, 0)
sql = "SELECT EvntRaceTypesID, EvntRaceType FROM EvntRaceTypes ORDER BY EvntRaceType"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventTypes(0, i) = rs(0).Value
	EventTypes(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve EventTypes(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("submit_clone_race") = "submit_clone_race" Then
    lCloneRace = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_clone_event") = "submit_clone_event" Then
    lCloneEvent = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	'get variables from add_race
	sRaceName = Replace(Request.Form.Item("race_name"), "'", "''")
	sDist = Request.Form.Item("race_dist")
	lRaceType = Request.Form.Item("race_type")
	sEntryFeePre = Request.Form.Item("pre_fee")
	sEntryFee = Request.Form.Item("fee")
	sStartTime = Request.Form.Item("start_time") & Request.Form.Item("am_pm")
	sCertified = Request.Form.Item("certified")
	iMAwds = Request.Form.Item("mawds")
	iFAwds = Request.Form.Item("fawds")
	sStartType = Request.Form.Item("start_type")
    sAllowDuplAwds = Request.Form.Item("allow_dupl_awds")

	'insert into racedata
	sql = "INSERT INTO RaceData (EventID, Dist, Type, EntryFeePre, EntryFee, StartTime, Certified, RaceName, StartType, MAwds, FAwds, BibsFrom, "
    sql = sql & "BibsTo, AllowDuplAwds) VALUES (" & lEventID & ", '" & sDist & "', '" & lRaceType & "', " & sEntryFeePre & ", " & sEntryFee & ", '" 
    sql = sql & sStartTime & "', '" & sCertified & "', '" & sRaceName & "', '" & sStartType & "', " & iMAwds & ", " & iFAwds 
    sql = sql & ", 0, 0, '" & sAllowDuplAwds & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
	
	'get race id
	sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID & " AND Dist = '" & sDist & "' AND Type = " & lRaceType
	Set rs = conn.Execute(sql)
	lRaceID = rs(0).Value
	Set rs = Nothing
	
	'write to tshirt tables if necessary
	If Request.Form.Item("has_shrts") = "y" Then
		If Request.Form.Item("small") ="on" Then
			sSmall = "y"
		Else
			sSmall = "n"
		End If
		
		If Request.Form.Item("medium") ="on" Then
			sMedium = "y"
		Else
			sMedium = "n"
		End If
		
		If Request.Form.Item("large") ="on" Then
			sLarge = "y"
		Else
			sLarge = "n"
		End If
		
		If Request.Form.Item("xlarge") ="on" Then
			sXLarge = "y"
		Else
			sXLarge = "n"
		End If
		
		If Request.Form.Item("xxlarge") ="on" Then
			sXXLarge = "y"
		Else
			sXXLarge = "n"
		End If
		
		If Request.Form.Item("short") ="on" Then
			sShort = "y"
		Else
			sShort = "n"
		End If
		
		If Request.Form.Item("long") ="on" Then
			sLong = "y"
		Else
			sLong = "n"
		End If
		
		If Request.Form.Item("choose_none") ="on" Then
			sChooseNone = "y"
		Else
			sChooseNone = "n"
		End If
		
		sql = "INSERT INTO TShirtData (RaceID, IsOption, Small, Medium, Large, XLarge, XXLarge, Short, Long, ChooseNone) "
		sql = sql & "VALUES (" & lRaceID & ", 'y', '" & sSmall & "', '" & sMedium & "', '" & sLarge & "', '" & sXLarge & "', '"
		sql = sql & sXXLarge & "', '" & sShort & "', '" & sLong & "', '" & sChooseNone & "')"
	Else
		sql = "INSERT INTO TShirtData (RaceID, IsOption) VALUES (" & lRaceID & ", 'n')"
	End If
	Set rs = conn.Execute(sql)
	Set rs =Nothing
	
	For i = 0 To 14
        If Not (Request.form.Item("m_end_age_" & i) & "" = "" AND Request.form.Item("m_end_age_" & i) = 0) Then
		    sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'm', "
		    sql = sql & Request.Form.Item("m_end_age_" & i) & ", " & Request.Form.Item("m_awds_" & i) & ")"
		    Set rs = conn.Execute(sql)
		    Set rs = Nothing
        End If
    Next

	For i = 0 To 14
        If Not (Request.form.Item("f_end_age_" & i) & "" = "" AND Request.form.Item("f_end_age_" & i) = 0) Then
   		    sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'f', "
		    sql = sql & Request.Form.Item("f_end_age_" & i) & ", " & Request.Form.Item("f_awds_" & i) & ")"
		    Set rs = conn.Execute(sql)
		    Set rs = Nothing
        End If
	Next

    'check for last end age = 110
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

	Response.Redirect "race_info.asp?event_id=" & lEventID
End If

If CStr(lCloneEvent) = vbNullString Then lCloneEvent = 0
If CStr(lCloneRace) = vbNullString Then lCloneRace = 0

If Not CLng(lCloneEvent) = 0 Then
    i = 0
    ReDim Races(1, 0)
    sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lCloneEvent
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    Races(0, i) = rs(0).Value
	    Races(1, i) = rs(1).Value
	    i = i + 1
	    ReDim Preserve Races(1, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing

    If UBound(Races, 2) = 1 Then lCloneRace = Races(0, 0)

    If Not CLng(lCloneRace) = 0 Then
	    sql = "SELECT RaceName, Dist, Type, EntryFeePre, EntryFee, StartTime, Certified, StartType, MAwds, "
	    sql = sql & "FAwds, OnlineRegLink, AllowDuplAwds FROM RaceData WHERE RaceID = " & lCloneRace
	    Set rs = conn.Execute(sql)
	
	    InfoArray(0) = rs(0).Value
		InfoArray(1) = rs(1).Value
        'no need for InfoArray(2)
	    InfoArray(3) = rs(2).Value
	    InfoArray(4) = rs(3).Value
	    InfoArray(5) = rs(4).Value
	
	    'split the time field
	    InfoArray(6) = Left(rs(5).Value, Len(rs(5).Value) - 2)
	    InfoArray(7) = Right(rs(5).Value, 2)
	
	    InfoArray(8) = rs(6).Value
	    InfoArray(9) = rs(7).Value
	    InfoArray(10) = rs(8).Value
	    InfoArray(11) = rs(9).Value
	    InfoArray(12) = rs(10).Value
	    InfoArray(13) = rs(11).Value
	    Set rs = Nothing

	    'get male age group array
	    i = 0
	    ReDim MaleArray(2, 0)
	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lCloneRace & " AND Gender = 'm' ORDER BY EndAge"
	    rs.Open sql, conn, 1, 2
	    Do While Not rs.EOF
            MaleArray(0, i) = i
		    MaleArray(1, i) = rs(0).Value
            MaleArray(2, i) = rs(1).Value
		    i = i + 1
		    ReDim Preserve MaleArray(2, i)
		    rs.MoveNext
	    Loop
	    rs.Close
	    Set rs = Nothing
	
        If UBound(MaleArray, 2) = 1 Then
            MaleArray(0, 0) = "0"
		    MaleArray(1, 0) = "110"
            MaleArray(2, 0) = "0"
		    ReDim Preserve MaleArray(2, 1)
        End If

	    'get female age group array
	    i = 0
	    ReDim FemaleArray(2, 0)
	    Set rs = SErver.CreateObject("ADODB.Recordset")
	    sql = "SELECT EndAge, NumAwds FROM AgeGroups WHERE RaceID = " & lCloneRace & " AND Gender = 'f' ORDER BY EndAge"
	    rs.Open sql, conn, 1, 2
	    Do While Not rs.EOF
            FemaleArray(0, i) = i
		    FemaleArray(1, i) = rs(0).Value
            FemaleArray(2, i) = rs(1).Value
    	    i = i + 1
		    ReDim Preserve FemaleArray(2, i)
		    rs.MoveNext
	    Loop
	    rs.Close
	    Set rs = Nothing
	
        If UBound(FemaleArray, 2) = 1 Then
            FemaleArray(0, 0) = "0"
		    FemaleArray(1, 0) = "110"
            FemaleArray(2, 0) = "0"
		    ReDim Preserve FemaleArray(2, 1)
        End If

	    'get t-shirt information
        Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT IsOption, Small, Medium, Large, XLarge, XXLarge, Short, Long, ChooseNone FROM TShirtData WHERE RaceID = " & lCloneRace
	    rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
	        For i = 0 to 8
		        TShirtArray(i) = rs(i).Value
	        Next
        End If
        rs.Close
	    Set rs = Nothing
    End If
End If

Private Function GetThisType(lEventType)
	sql = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs = conn.Execute(sql)
	GetThisType = rs(0).Value
	Set rs = Nothing
End Function

StartType(0) = "mass"
StartType(1) = "wave"
StartType(2) = "interval"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Clone a Race for <%=sEventName%></title>

<script>
function chkFlds(){
 	if (document.race_info.race_dist.value == '' || 
 	    document.race_info.race_name.value == '' ||
 	    document.race_info.race_type.value == '' ||
	 	document.race_info.hours.value == '' || 
	 	document.race_info.minutes.value == '' || 
	 	document.race_info.am_pm.value == '' ||
	 	document.race_info.fee.value == '' ||
	 	document.race_info.pre_fee.value == '' ||
	 	document.race_info.certified.value == '')
		{
  		alert('Please fill in all required fields');
  		return false
  		}
 	else
		if (isNaN(document.race_info.race_dist.value) ||
		   isNaN(document.race_info.fee.value) ||
		   isNaN(document.race_info.pre_fee.value))
    		{
			alert('The race distance, entry fee and pre-reg entry fee must be numeric values!');
			return false
			} 	
  	else
  		if (document.race_info.choose_none.checked != true &&
	 	   document.race_info.short_sleeve.checked != true &&
	 	   document.race_info.long_sleeve.checked != true)
	 	   {
	 	   alert('Please select a shirt style for this race!');
	 	   return false
	 	   }
  	else
  		if (document.race_info.choose_none.checked != true &&
	 	   (document.race_info.small.checked != true &&
			document.race_info.medium.checked != true &&	 	   
			document.race_info.large.checked !=  true && 	   
			document.race_info.xlarge.checked !=  true &&	 	   
			document.race_info.xxlarge.checked !=  true))
	 	   {
	 	   alert('Please select shirt sizes for this race!');
	 	   return false
	 	   }
	else
   		return true
}

    var OK = true;
    for(var i=0; i<=14; i++){
        if (isNaN(document.race_info['magegrp'+i].value) ||
            isNaN(document.race_info['mawds'+i].value)   ||
            isNaN(document.race_info['fagegrp'+i].value) ||
            isNaN(document.race_info['fawds'+i].value)) {
            OK = false;
            break; 
        }
    }
    if (!OK) alert('All data entered on this page must be numeric!');
    return OK;
}
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4"><%=sEventName%>: Clone Race (<span style="font-weight:normal">* = required fields</span></h4>
			
			<!--#include file = "../../includes/event_nav.asp" -->

			<div style="margin:10px;">
				<form name="which_clone_event" method="post" action="clone_race.asp?event_id=<%=lEventID%>">
				<span style="font-weight:bold;">Select Event To Clone From:</span>
				<select name="events" id="events" onchange="this.form.get_event.click()">
					<option value="">&nbsp;</option>
					<%For i = 0 to UBound(Events, 2) - 1%>
						<%If CLng(lCloneEvent) = CLng(Events(0, i)) Then%>
							<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
						<%Else%>
							<option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
						<%End If%>
					<%Next%>
				</select>
				<input type="hidden" name="submit_clone_event" id="submit_clone_event" value="submit_clone_event">
				<input type="submit" name="get_event" id="get_event" value="Clone From This Event">
				</form>
			</div>

            <%If Not CLng(lCloneEvent) = 0 Then%>
			    <div style="margin:10px;">
  				    <form name="which_race" method="post" action="clone_race.asp?event_id=<%=lEventID%>&amp;clone_event=<%=lCloneEvent%>">
				    <span style="font-weight:bold;">Select Race To Clone:</span>
				    <select name="races" id="races" onchange="this.form.get_race.click()">
					    <option value="">&nbsp;</option>
					    <%For i = 0 to UBound(Races, 2) - 1%>
						    <%If CLng(lCloneRace) = CLng(Races(0, i)) Then%>
							    <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
						    <%Else%>
							    <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
						    <%End If%>
					    <%Next%>
				    </select>
				    <input type="hidden" name="submit_clone_race" id="submit_clone_race" value="submit_clone_race">
				    <input type="submit" name="get_race" id="get_race" value="Clone This Race">
				    </form>

                    <br>

                    <hr>

                    </div>
                <%If Not CLng(lCloneRace) = 0 Then%>
   			        <form name="race_info" method="post" action="clone_race.asp?event_id=<%=lEventID%>&amp;clone_event=<%=lCloneEvent%>&amp;clone_race=<%=lCloneRace%>" onsubmit="return chkFlds()">
				    <table style="margin-left:10px;">
					    <tr>	
						    <th>Race Name:</th>
						    <td><input name="race_name" id="race_name" maxlength="35" value="<%=InfoArray(0)%>"></td>
						    <th>Distance:</th>
						    <td><input name="race_dist" id="race_dist" maxlength="8" size="4" value="<%=InfoArray(1)%>"></td>
						    <th>Race Type:</th>
						    <td colspan="3">
							    <select name="race_type" id="race_type">
								    <%For i = 0 to UBound(EventTypes, 2) - 1%>
									    <%If CInt(EventTypes(0, i)) = CInt(InfoArray(3)) Then%>
										    <option value="<%=EventTypes(0, i)%>" selected><%=EventTypes(1, i)%></option>
									    <%Else%>
										    <option value="<%=EventTypes(0, i)%>"><%=EventTypes(1, i)%></option>
									    <%End If%>
								    <%Next%>
							    </select>
						    </td>
					    </tr>
					    <tr>	
						    <th>Start Type:</th>
						    <td>
							    <select name="start_type" id="start_type">
								    <%If InfoArray(9) = "mass" Then%>
									    <option value="mass" selected>Mass</option>
									    <option value="wave">Wave</option>
									    <option value="interval">Interval</option>
								    <%ElseIf InfoArray(9) = "wave" Then%>
									    <option value="mass">Mass</option>
									    <option value="wave" selected>Wave</option>
									    <option value="interval">Interval</option>
								    <%ElseIf InfoArray(9) = "interval" Then%>
									    <option value="mass">Mass</option>
									    <option value="wave">Wave</option>
									    <option value="interval" selected>Interval</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Start Time:</th>
						    <td>
							    <input name="start_time" id="start_time" size="3" maxlength="5" value="<%=InfoArray(6)%>" onchange="return chkStr(this)">
							    <select name="am_pm" id="am_pm">
								    <%If InfoArray(7) = "am" Then%>
									    <option value="am">am</option>
									    <option value="pm">pm</option>
								    <%Else%>
									    <option value="am">am</option>
									    <option value="pm" selected>pm</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Certified:</th>
						    <td>
							    <select name="certified" id="certified">
								    <%If InfoArray(8) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Allow Dupl Awds:</th>
						    <td>
							    <select name="allow_dupl_awds" id="allow_dupl_awds">
								    <%If InfoArray(13) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
					    </tr>
					    <tr>	
						    <th>Pre-reg Fee:&nbsp;$</th>
						    <td><input name="pre_fee" id="pre_fee" size="3" maxlength="5" value="<%=InfoArray(4)%>"></td>
						    <th>Race Day Fee:&nbsp;$</th>
						    <td><input name="fee" id="fee" size="3" maxlength="5" value="<%=InfoArray(5)%>"></td>
						    <th>M Open Awards:</th>
						    <td><input name="mawds" id="mawds" size="2" maxlength="2" value="<%=InfoArray(10)%>"></td>
						    <th>F Open Awards:</th>
						    <td><input name="fawds" id="fawds" size="1" maxlength="2" value="<%=InfoArray(11)%>"></td>
					    </tr>
					    <tr>	
						    <th>Online Reg Link:</th>
						    <td colspan="7"><input type="text" name="online_reg_link" id="online_reg_link" value="<%=InfoArray(12)%>" size="50"></td>
					    </tr>
				    </table>
				                    
                    <hr style="margin: 10px;">

				    <table style="margin-left:10px;">
                        <tr>
					        <td style="padding-right:25px;" valign="top">
						        <h4 style="margin-bottom: 0px;padding: 0px;">Age Groups</h4>
						        <table style="margin-top: 0px;padding: 0px;">
							        <tr>
								        <td style="margin:0px;padding: 0px;" colspan="2">
									        <table style="background-color:#fff;font-size:1.2em;">
										        <tr>
											        <th style="border-bottom:1px solid #ececd8;" colspan="2">Male Age Groups</th>
										        </tr>
										        <tr>
											        <td style="text-align:center;">End Age</td>
											        <td style="text-align:center;">Awds</td>
										        </tr>
										        <%If UBound(MaleArray, 2) > 1 Then%>
											        <%For i = 0 to UBound(MaleArray, 2) - 1%>
												        <tr>
													        <td style="text-align:center;">
														        <%If MaleArray(1, i) = "110" Then%>
                                                                    <%If i = 0 Then%>
															            <input type="text" name="m_end_age_0" id="m_end_age_0" 
															                      value="110 And Over" size="10" style="text-align:right" disabled>
                                                                    <%Else%>
															            <input type="text" name="m_end_age_<%=MaleArray(0, i)%>" 
															                      id="m_end_age_<%=MaleArray(0, i)%>" 
															                      value="<%=CInt(MaleArray(1, i - 1)) + 1%>  And Over" 
															                      size="10" style="text-align:right" disabled>
                                                                    <%End If%>
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
												        </tr>
											        <%Next%>
                                                <%Else%>
												    <tr>
													    <td style="text-align:center;">
															<input type="text" name="m_end_age_0" id="m_end_age_0" 
															            value="110 And Over" 
															            size="10" style="text-align:right" disabled>
													    </td>
													    <td style="text-align:center;">
														    <input type="text" name="m_awds_<%=MaleArray(0, 0)%>" id="m_awds_<%=MaleArray(0, 0)%>" size="1"
																        maxlength="2" value="<%=MaleArray(2, 0)%>">
													    </td>
												    </tr>
										        <%End If%>
									        </table>
								        </td>
								        <td colspan="2">
									        <table style="background-color:#fff;font-size:1.2em;">
										        <tr>
											        <th style="border-bottom:1px solid #ececd8;" colspan="2">Female Age Groups</th>
										        </tr>
										        <tr>
											        <td style="text-align:center;">End Age</td>
											        <td style="text-align:center;">Awds</td>
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
												        </tr>
											        <%Next%>
                                                <%Else%>
												    <tr>
													    <td style="text-align:center;">
															<input type="text" name="f_end_age_0" id="f_end_age_0" 
															            value="110 And Over" 
															            size="10" style="text-align:right" disabled>
													    </td>
													    <td style="text-align:center;">
														    <input type="text" name="f_awds_<%=FemaleArray(0, 0)%>" id="f_awds_<%=FemaleArray(0, 0)%>" size="1"
																        maxlength="2" value="<%=FemaleArray(2, 0)%>">
													    </td>
												    </tr>
										        <%End If%>
									        </table>
								        </td>
							    </tr>
						    </table>
					    </td>
					    <td valign="top">
						    <h4 style="margin-bottom: 0px;">T-Shirt Data</h4>
						    <table style="font-size:1.0em;">
							    <tr>
								    <td style="background-color:#fff;">
									    Are T-Shirts Available for the race?&nbsp;
									    <select name="has_shirts" id="has_shirts">
										    <%If TShirtArray(0) = "y" Then%>
											    <option value="y" selected>Yes</option>
											    <option value="n">No</option>
										    <%Else%>
											    <option value="y">Yes</option>
											    <option value="n" selected>No</option>
										    <%End If%>
									    </select>
								    </td>
							    </tr>
							    <tr>
								    <td style="background-color:#fff;">
										<span style="font-weight: bold;">Sleeve Length?</span><br>
										<%If TShirtArray(6) = "y" Then%>
											<input type="checkbox" name="short" checked>&nbsp;Short Sleeve&nbsp;&nbsp;&nbsp;
										<%Else%>
											<input type="checkbox" name="short">&nbsp;Short Sleeve&nbsp;&nbsp;&nbsp;
										<%End If%>
										
										<%If TShirtArray(7) = "y" Then%>
											<input type="checkbox" name="long" checked>&nbsp;Long Sleeve
										<%Else%>
											<input type="checkbox" name="long">&nbsp;Long Sleeve
										<%End If%>
								    </td>
							    </tr>
							    <tr>
								    <td style="background-color:#fff;">
										<span style="font-weight: bold;">Which sizes are available?</span><br>

										<%If TShirtArray(1) = "y" Then%>
											<input type="checkbox" name="small" checked>&nbsp;Small&nbsp;&nbsp;&nbsp;
										<%Else%>
											<input type="checkbox" name="small">&nbsp;Small&nbsp;&nbsp;&nbsp;
										<%End If%>
										
										<%If TShirtArray(2) = "y" Then%>
											<input type="checkbox" name="medium" checked>&nbsp;Medium&nbsp;&nbsp;&nbsp;
										<%Else%>
											<input type="checkbox" name="medium">&nbsp;Medium&nbsp;&nbsp;&nbsp;
										<%End If%>
										
										<%If TShirtArray(3) = "y" Then%>
											<input type="checkbox" name="large" checked>&nbsp;Large&nbsp;&nbsp;&nbsp;
										<%Else%>
											<input type="checkbox" name="large">&nbsp;Large&nbsp;&nbsp;&nbsp;
										<%End If%>
										
										<%If TShirtArray(4) = "y" Then%>
											<input type="checkbox" name="xlarge" checked>&nbsp;X-Large&nbsp;&nbsp;&nbsp;
										<%Else%>
											<input type="checkbox" name="xlarge">&nbsp;X-Large&nbsp;&nbsp;&nbsp;
										<%End If%>
										
										<%If TShirtArray(5) = "y" Then%>
											<input type="checkbox" name="xxlarge" checked>&nbsp;XX-Large
										<%Else%>
											<input type="checkbox" name="xxlarge">&nbsp;XX-Large
										<%End If%>
								    </td>
							    </tr>
							    <tr>
								    <td style="background-color:#fff;">
									    Can participants choose not to have a t-shirt?&nbsp;
									    <select name="choose_none" id="choose_none">
										    <%If TShirtArray(8) = "y" Then%>
											    <option value="y" selected>Yes</option>
											    <option value="n">No</option>
										    <%Else%>
											    <option value="y">Yes</option>
											    <option value="n" selected>No</option>
										    <%End If%>
									    </select>
								    </td>
							    </tr>
							   </table>
                            </td>
                        </tr>
					</table>

					<div style="text-align:center;background-color: #ececd8;">
						<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
						<input type="submit" name="submit3" id="submit3" value="Save Changes">
					</div>
			        </form>
                <%End If%>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
