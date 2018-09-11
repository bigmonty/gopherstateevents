<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lRaceID, lEventID
Dim iBibsFrom, iBibsTo, iAgeGrpAwds, iNumSplits, iEndAge, iNewEndAge
Dim sEventName, sStartTime, sRaceName, sErrMsg, sSortRsltsBy, sInSeries, sAgeGrpName, sNewGender, sThisGender
Dim dEventDate
Dim InfoArray(23), EventTypes(), MaleArray(), FemaleArray(), StartType(2), Delete()
Dim bBibsOverlap, bFound, bLastGrpExists

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
lRaceID = Request.QueryString("race_id")

iBibsFrom = 0
iBibsTo = 0

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

If Request.Form.Item("submit_multi") = "submit_multi" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT NumLegs, Leg1Name, Leg2Name, Leg3Name, Leg1Dist, Leg2Dist, Leg3Dist FROM MultiSettingsChip WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("num_legs")
    rs(1).Value = Request.Form.Item("leg_1_name")
    rs(2).Value = Request.Form.Item("leg_2_name")
    rs(3).Value = Request.Form.Item("leg_3_name")
    rs(4).Value = Request.Form.Item("leg_1_dist")
    rs(5).Value = Request.Form.Item("leg_2_dist")
    rs(6).Value = Request.Form.Item("leg_3_dist")
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_age_grps") = "submit_age_grps" Then
    i = 0
    ReDim Delete(0)

	'write male back to db
    iAgeGrpAwds = 0
    bLastGrpExists = False
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AgeGroupsID, EndAge, NumAwds, AgeGrpName FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'm' ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF	
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
		    If rs(1).Value = "110" Then 
                bLastGrpExists = True
            Else
			    If IsNumeric(Request.Form.Item("m_end_age_" & rs(0).Value)) Then
				    rs(1).Value = Request.Form.Item("m_end_age_" & rs(0).Value)	
			    Else
                    rs(1).Value = rs(1).OriginalValue
				    sErrMsg = "All ending ages must be numeric.  Some work was not done."
			    End If
		    End If
			
			If IsNumeric(Request.Form.Item("m_awds_" & rs(0).Value)) Then
                iAgeGrpAwds = Request.Form.Item("m_awds_" & rs(0).Value)
				rs(2).Value = iAgeGrpAwds
			Else
                iAgeGrpAwds = rs(2).OriginalValue
                rs(2).Value = iAgeGrpAwds
				sErrMsg = "All award values must be numeric.  Some work was not done."
			End If
			
			If Request.Form.Item("m_name_" & rs(0).Value) & "" = "" Then
				rs(3).Value = rs(3).OriginalValue
			Else
				rs(3).Value = Request.Form.Item("m_name_" & rs(0).Value)
			End If
		    rs.Update
        End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM AgeGroups WHERE AgeGroupsID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next

    sThisGender = "m"
    If UBound(Delete) > 0 Then Call ReNameAgeGrps(sThisGender)

    If bLastGrpExists = False Then
        sAgeGrpName = vbNullString

        Set rs = Serer.CreateObject("ADODB.Recordset")
        sql = "SELECT EndAge FROM AgeGroups WHERE Gender = 'm' AND RaceID - " & lRaceID & " ORDER BY EndAge DESC"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            sAgeGrpName = CInt(rs(0).Value) + 1 & " and Over"
        Else
            sAgeGrpName = "110 and Under"
        End If
        rs.Close
        Set rs = Nothing

		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds, AgeGrpName) VALUES (" & lRaceID & ", 'm', 110, " & iAgeGrpAwds 
        sql = sql & ", '" & sAgeGrpName & "')"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    End If
	
	If CInt(Request.Form.Item("new_m_end_age")) > 0 Then
        iNewEndAge = Request.Form.Item("new_m_end_age")
        sNewGender = "m"

		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds, AgeGrpName) VALUES (" & lRaceID & ", 'm', " & iNewEndAge & ", " 
        sql = sql & Request.Form.Item("new_m_awds") & ", '" & NewAgeGrpName(iNewEndAge, sNewGender) & "')"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	End If

	'write female back to db
    i = 0
    ReDim Delete(0)
    iAgeGrpAwds = 0
    bLastGrpExists = False
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AgeGroupsID, EndAge, NumAwds, AgeGrpName FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'f' ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF		
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
		    If rs(1).Value = "110" Then 
                bLastGrpExists = True
            Else
			    If IsNumeric(Request.Form.Item("f_end_age_" & rs(0).Value)) Then
				    rs(1).Value = Request.Form.Item("f_end_age_" & rs(0).Value)	
			    Else
                    rs(1).Value = rs(1).OriginalValue
				    sErrMsg = "All ending ages must be numeric.  Some work was not done."
			    End If
		    End If
			
			If IsNumeric(Request.Form.Item("f_awds_" & rs(0).Value)) Then
                iAgeGrpAwds = Request.Form.Item("f_awds_" & rs(0).Value)
				rs(2).Value = iAgeGrpAwds
			Else
                iAgeGrpAwds = rs(2).OriginalValue
                rs(2).Value = iAgeGrpAwds
				sErrMsg = "All award values must be numeric.  Some work was not done."
			End If
			
			If Request.Form.Item("f_name_" & rs(0).Value) & "" = "" Then
				rs(3).Value = rs(3).OriginalValue
			Else
				rs(3).Value = Request.Form.Item("f_name_" & rs(0).Value)
			End If
		    rs.Update
        End IF
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	
    If bLastGrpExists = False Then
        sAgeGrpName = vbNullString

        Set rs = Serer.CreateObject("ADODB.Recordset")
        sql = "SELECT EndAge FROM AgeGroups WHERE Gender = 'f' AND RaceID - " & lRaceID & " ORDER BY EndAge DESC"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            sAgeGrpName = CInt(rs(0).Value) + 1 & " and Over"
        Else
            sAgeGrpName = "110 and Under"
        End If
        rs.Close
        Set rs = Nothing

		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds, AgeGrpName) VALUES (" & lRaceID & ", 'f', 110, " & iAgeGrpAwds 
        sql = sql & ", '" & sAgeGrpName & "')"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    End If

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM AgeGroups WHERE AgeGroupsID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next

    sThisGender = "f"
    If UBound(Delete) > 0 Then Call ReNameAgeGrps(sThisGender)
	
	If CInt(Request.Form.Item("new_f_end_age")) > 0 Then
        iNewEndAge = Request.Form.Item("new_f_end_age")
        sNewGender = "f"

		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds, AgeGrpName) VALUES (" & lRaceID & ", 'f', " & iNewEndAge & ", " 
        sql = sql & Request.Form.Item("new_f_awds") & ", '" & NewAgeGrpName(iNewEndAge, sNewGender) & "')"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	End If
ElseIf Request.Form.Item("submit_race_info") = "submit_race_info" Then
    If Request.Form.Item("delete_race") = "on" Then
        sql = "DELETE FROM RaceData WHERE RaceID = " & lRaceID
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        lRaceID = 0
    Else
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
	
	    'check to see if bibs overlap
	
	    'if all is good then...
	    If sErrMsg = vbNullString Then
		    Set rs=Server.CreateObject("ADODB.Recordset")
		    sql = "SELECT RaceName, Dist, Type, StartTime, Certified, StartType, MAwds, FAwds, BibsFrom, BibsTo, OnlineRegLink, "
            sql = sql & "AllowDuplAwds, ChipStart, SortRsltsBy, EntryFeePre, EntryFee, InSeries, NumSplits, Timed, ShowAge, StartToFinish, RaceDelay, "
            sql = sql & "Numlaps, MinLap, IndivRelay FROM RaceData WHERE RaceID = " & lRaceID
		    rs.Open sql, conn, 1, 2
		    rs(0).Value = Replace(Request.Form.Item("race_name"), "'", "''")
		    rs(1).Value = Request.Form.Item("dist")
		    rs(2).Value = Request.Form.Item("race_type")
		    rs(3).Value = Request.Form.Item("start_time") & Request.Form.Item("am_pm")
		    rs(4).Value = Request.Form.Item("certif")
		    rs(5).Value = Request.Form.Item("start_type")
		    rs(6).Value = Request.Form.Item("mawds")
		    rs(7).Value = Request.Form.Item("fawds")
    	    rs(8).Value = Request.Form.Item("bibs_from")
    	    rs(9).Value = Request.Form.Item("bibs_to")
    	    rs(10).Value = Request.Form.Item("online_reg_link")
    	    rs(11).Value = Request.Form.Item("allow_dupl_awds")
            rs(12).Value = Request.Form.Item("chip_start")
            rs(13).Value = Request.Form.Item("sort_rslts_by")
            rs(14).Value = Request.Form.Item("entry_fee_pre")
            rs(15).Value = Request.Form.Item("entry_fee")
            rs(16).Value = Request.Form.Item("in_series")
            rs(17).Value = Request.Form.Item("num_splits")
            rs(18).Value = Request.Form.Item("timed")
            rs(19).Value = Request.Form.Item("show_age")
            rs(20).Value = Request.Form.Item("start_to_finish")
            rs(21).Value = Request.Form.Item("race_delay")
            rs(22).Value = Request.Form.Item("num_laps")
            rs(23).Value = Request.Form.Item("min_lap")
            rs(24).Value = Request.Form.Item("indiv_relay")
		    rs.Update
		    rs.Close
		    Set rs = Nothing
	    End If
    End If
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
End If

If CStr(lRaceID) = vbNullString Then lRaceID = 0
	
'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

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

If UBound(RaceArray, 2) = 1 Then lRaceID = RaceArray(0, 0)

If Not CLng(lRaceID) = 0 Then
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

	sql = "SELECT RaceName, Dist, Type, StartTime, Certified, StartType, MAwds, FAwds, OnlineRegLink, AllowDuplAwds, ChipStart, "
    sql = sql & "SortRsltsBy, EntryFeePre, EntryFee, InSeries, NumSplits, Timed, ShowAge, StartToFinish, RaceDelay, Numlaps, MinLap, IndivRelay "
    sql = sql & "FROM RaceData WHERE RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	
	InfoArray(0) = rs(0).Value
	InfoArray(1) = rs(1).Value
	InfoArray(2) = rs(2).Value
	
	'split the time field
	InfoArray(3) = Left(rs(3).Value, Len(rs(3).Value) - 2)
	InfoArray(4) = Right(rs(3).Value, 2)
	
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
    InfoArray(15) = rs(14).Value
    InfoArray(16) = rs(15).Value
    InfoArray(17) = rs(16).Value
    InfoArray(18) = rs(17).Value
    InfoArray(19) = rs(18).Value
	InfoArray(20) = rs(19).Value
    InfoArray(21) = rs(20).Value
    InfoArray(22) = rs(21).Value
    InfoArray(23) = rs(22).Value
    Set rs = Nothing

    If InfoArray(19) & "" = "" Then InfoArray(19) = "0"

    If InfoArray(2) = "9" Then
        Dim iNumLegs
        Dim sLeg1Name, sLeg2Name, sLeg3Name
        Dim sngLeg1Dist, sngLeg2Dist, sngLeg3Dist

        iNumLegs = 3

        sLeg1Name = "Swim"
        sLeg2Name = "Bike"
        sLeg3Name = "Run"
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT NumLegs, Leg1Name, Leg2Name, Leg3Name, Leg1Dist, Leg2Dist, Leg3Dist FROM MultiSettingsChip WHERE RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then 
            iNumLegs = rs(0).Value

            If Not rs(1).Value & "" = "" Then sLeg1Name = rs(1).Value
            If Not rs(2).Value & "" = "" Then sLeg2Name = rs(2).Value
            If Not rs(3).Value & "" = "" Then sLeg3Name = rs(3).Value

            sngLeg1Dist = rs(4).Value
            sngLeg2Dist = rs(5).Value
            sngLeg3Dist = rs(6).Value
        End If
        rs.Close
        Set rs = Nothing

        If sngLeg1Dist & "" = "" Then sngLeg1Dist = "unknown"
        If sngLeg2Dist & "" = "" Then sngLeg2Dist = "unknown"
        If sngLeg3Dist & "" = "" Then sngLeg3Dist = "unknown"
    End if

	'get male age group array
	i = 0
	ReDim MaleArray(3, 0)
	Set rs = SErver.CreateObject("ADODB.Recordset")
	sql = "SELECT AgeGroupsID, EndAge, NumAwds, AgeGrpName FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'm' ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		For j = 0 to 3
			MaleArray(j, i) = rs(j).Value
		Next
		
		i = i + 1
		ReDim Preserve MaleArray(3, i)
		
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

	'get female age group array
	i = 0
	ReDim FemaleArray(3, 0)
	Set rs = SErver.CreateObject("ADODB.Recordset")
	sql = "SELECT AgeGroupsID, EndAge, NumAwds, AgeGrpName FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'f' ORDER BY EndAge"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		For j = 0 to 3
			FemaleArray(j, i) = rs(j).Value
		Next
		
		i = i + 1
		ReDim Preserve FemaleArray(3, i)
		
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

	'get bib range
	Set rs=Server.CreateObject("ADODB.Recordset")
    sql = "SELECT BibsFrom, BibsTo FROM RaceData WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    iBibsFrom = rs(0).Value
    iBibsTo = rs(1).Value
    rs.Close
    Set rs = Nothing
End If

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

StartType(0) = "mass"
StartType(1) = "wave"
StartType(2) = "interval"

Private Function NewAgeGrpName(iThisEndAge, sThisGender)
    Dim iBegAge, rs2, sql2
    Dim x

    x = 0                   'to determine if we are changing the first age group

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT EndAge, AgeGrpName FROM AgeGroups WHERE Gender = '" & sThisGender & "' AND RaceID = " & lRaceID & " ORDER BY EndAge"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        If CInt(iThisEndAge) < CInt(rs2(0).Value) Then
            If x = 0 Then
                rs2(1).Value = iThisEndAge + 1 & " - " & rs2(0).Value
                NewAgeGrpName = iThisEndAge & " and Under"
            Else
                NewAgeGrpName = " - " & iThisEndAge
                If rs2(0).Value = "110" Then
                    rs2(1).Value = CInt(iThisEndAge) + 1 & " and Over"
                Else
                    rs2(1).Value = CInt(iThisEndAge) + 1 & " - " & rs2(0).Value
                End If
            End If

            rs2.Update
            Exit Do
        End If

        x = x + 1

        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If x > 0 Then
        ' finish naming this age group and update age group name below this one
        Set rs2 = Server.CreateObject("ADODB.Recordset")
        sql2 = "SELECT EndAge, AgeGrpName FROM AgeGroups WHERE Gender = '" & sThisGender & "' AND RaceID = " & lRaceID & " ORDER BY EndAge DESC"
        rs2.Open sql2, conn, 1, 2
        Do While Not rs2.EOF
            If CInt(iThisEndAge) > CInt(rs2(0).Value) Then 
                NewAgeGrpName = CInt(rs2(0).Value) + 1 & " - " & iThisEndAge    'finish naming the new age group
                Exit Do
            End If
            rs2.MoveNext
        Loop
        rs2.Close
        Set rs2 = Nothing
    End If
End Function

Private Sub ReNameAgeGrps(sAgeGrpGender)
    Dim sql2, rs2
    Dim iBegAge, iEndAge

    iBegAge = 0

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT AgeGroupsID, EndAge, AgeGrpName FROM AgeGroups WHERE Gender = '" & sAgeGrpGender & "' AND RaceID = " & lRaceID & " ORDER BY EndAge"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        iEndAge = rs2(1).Value

        If CInt(iBegAge) = 0 Then 
            rs2(2).Value = iEndAge & " and Under"
            iBegAge = CInt(iEndAge) + 1            'get beg age for next age group
        Else
            If CInt(iEndAge) = 110 Then
                rs2(2).Value = iBegAge & " and Over"
            Else
                rs2(2).Value = iBegAge & " - " & iEndAge
            End If

            iBegAge = CInt(iEndAge) + 1            'get beg age for next age group
        End If

        rs2.Update
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Race Info</title>
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

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">Race Info:&nbsp;<%=sEventName%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="race_info.asp?event_id=<%=lEventID%>">
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

			<%If Not Clng(lEventID) = 0 Then%>
			    <!--#include file = "../../includes/event_nav.asp" -->

			    <div style="margin: 0;padding: 0;font-size: 0.85em;">
				    <a href="/admin/events/add_race.asp?event_id=<%=lEventID%>" rel="nofollow">Add Race</a>
                </div>
				
			    <%If UBound(RaceArray, 2) > 1 Then%>
				    <form class="form-inline" name="get_races" method="post" action="race_info.asp?event_id=<%=lEventID%>">
				    <label for="races">Select Race:</label>
					<select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
						<option value="">&nbsp;</option>
						<%For i = 0 to UBound(RaceArray, 2) - 1%>
							<%If CLng(lRaceID) = CLng(RaceArray(0, i)) Then%>
								<option value="<%=RaceArray(0, i)%>" selected><%=RaceArray(1, i)%></option>
							<%Else%>
								<option value="<%=RaceArray(0, i)%>"><%=RaceArray(1, i)%></option>
							<%End If%>
						<%Next%>
					</select>
					<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
					<input type="submit" class="form-control" name="get_race" id="get_race" value="Get Race Info">
				    </form>
			    <%End If%>
			
			    <%If Not CLng(lRaceID) = 0 Then%>
				    <%If Not sErrMsg = vbNullString Then%>
					    <p style="margin-left:10px;"><%=sErrMsg%></p>
				    <%End If%>

				    <form role="form" class="form" name="update_info" method="post" action="race_info.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" 
                        onsubmit="return checkFields()">
				    <table style="margin-left:10px;">
					    <tr>	
						    <th>Race Name:</th>
						    <td><input name="race_name" id="race_name" maxlength="35" value="<%=InfoArray(0)%>"></td>
						    <th>Distance:</th>
						    <td><input name="dist" id="dist" maxlength="12" size="3" value="<%=InfoArray(1)%>"></td>
						    <th>Type:</th>
						    <td>
							    <select name="race_type" id="race_type">
								    <%For i = 0 to UBound(EventTypes, 2) - 1%>
									    <%If CInt(EventTypes(0, i)) = CInt(InfoArray(2)) Then%>
										    <option value="<%=EventTypes(0, i)%>" selected><%=EventTypes(1, i)%></option>
									    <%Else%>
										    <option value="<%=EventTypes(0, i)%>"><%=EventTypes(1, i)%></option>
									    <%End If%>
								    <%Next%>
							    </select>
						    </td>
					    </tr>
					    <tr>	
						    <th>Certified:</th>
						    <td>
							    <select name="certif" id="certif">
								    <%If InfoArray(5) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Start Type:</th>
						    <td>
							    <select name="start_type" id="start_type">
								    <%If InfoArray(6) = "mass" Then%>
									    <option value="mass" selected>Mass</option>
									    <option value="wave">Wave</option>
									    <option value="interval">Interval</option>
								    <%ElseIf InfoArray(6) = "wave" Then%>
									    <option value="mass">Mass</option>
									    <option value="wave" selected>Wave</option>
									    <option value="interval">Interval</option>
								    <%ElseIf InfoArray(6) = "interval" Then%>
									    <option value="mass">Mass</option>
									    <option value="wave">Wave</option>
									    <option value="interval" selected>Interval</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Start Time:</th>
						    <td>
							    <input name="start_time" id="start_time" size="3" maxlength="5" value="<%=InfoArray(3)%>" onchange="return chkStr(this)">
							    <select name="am_pm" id="am_pm">
								    <%If InfoArray(4) = "am" Then%>
									    <option value="am">am</option>
									    <option value="pm">pm</option>
								    <%Else%>
									    <option value="am">am</option>
									    <option value="pm" selected>pm</option>
								    <%End If%>
							    </select>
						    </td>
                        </tr>
                        <tr>
 						    <th>Chip Start:</th>
						    <td>
							    <select name="chip_start" id="chip_start">
								    <%If InfoArray(11) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Dupl Awds:</th>
						    <td>
							    <select name="allow_dupl_awds" id="allow_dupl_awds">
								    <%If InfoArray(10) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Sort Rslts By:</th>
						    <td>
							    <select name="sort_rslts_by" id="sort_rslts_by">
								    <%If InfoArray(12) = "FnlTime" Then%>
									    <option value="FnlTime" selected>FnlTime</option>
                                        <option value="place">place</option>
								    <%Else%>
									    <option value="FnlTime">FnlTime</option>
                                        <option value="place" selected>place</option>
								    <%End If%>
							    </select>
						    </td>
                        </tr>
                        <tr>
 						    <th>In Series?</th>
						    <td>
							    <select name="in_series" id="in_series">
								    <%If InfoArray(15) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>Pre-Reg Fee:</th>
                            <td><input name="entry_fee_pre" id="entry_fee_pre" size="2"value="<%=InfoArray(13)%>"></td>
						    <th>Race Day Fee:</th>
						    <td><input name="entry_fee" id="entry_fee" size="2" value="<%=InfoArray(14)%>"></td>
                        </tr>
                        <tr>
 						    <th>Race Delay:</th>
						    <td><input name="race_delay" id="race_delay" size="8"value="<%=InfoArray(20)%>"></td>
 						    <th>M Open Awds:</th>
						    <td><input name="mawds" id="mawds" size="2"value="<%=InfoArray(7)%>"></td>
						    <th>F Open Awds:</th>
						    <td><input name="fawds" id="fawds" size="2" value="<%=InfoArray(8)%>"></td>
					    </tr>
					    <tr>	
						    <th>Timed:</th>
						    <td>
							    <select name="timed" id="timed">
								    <%If InfoArray(17) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
						    <th>First Bib:</th>
						    <td>
							    <select name="bibs_from" id="bibs_from">
								    <%For i = 0 To 5900%>
									    <%If CInt(i) = CInt(iBibsFrom) Then%>
										    <option value="<%=i%>" selected><%=i%></option>
									    <%Else%>
										    <option value="<%=i%>"><%=i%></option>
									    <%End If%>
								    <%Next%>
							    </select>
						    </td>
						    <th>Last Bib:</th>
						    <td>
							    <select name="bibs_to" id="bibs_to">
								    <%For i = 0 To 7000%>
									    <%If CInt(i) = CInt(iBibsTo) Then%>
										    <option value="<%=i%>" selected><%=i%></option>
									    <%Else%>
										    <option value="<%=i%>"><%=i%></option>
									    <%End If%>
								    <%Next%>
							    </select>
						    </td>
                        </tr>
                        <tr>
						    <th>Show Age:</th>
						    <td>
							    <select name="show_age" id="show_age">
								    <%If InfoArray(18) = "y" Then%>
									    <option value="y" selected>Yes</option>
									    <option value="n">No</option>
								    <%Else%>
									    <option value="y">Yes</option>
									    <option value="n" selected>No</option>
								    <%End If%>
							    </select>
						    </td>
 						    <th>Num Splits:</th>
						    <td>
							    <select name="num_splits" id="num_splits">
                                    <%For i = 0 To 4%>
								        <%If CInt(InfoArray(16)) = CInt(i) Then%>
									        <option value="<%=i%>" selected><%=i%></option>
								        <%Else%>
									        <option value="<%=i%>"><%=i%></option>
								        <%End If%>
                                    <%Next%>
							    </select>
						    </td>
 						    <th>Num Laps:</th>
						    <td>
							    <select name="num_laps" id="num_laps">
                                    <%For i = 1 To 6%>
								        <%If CInt(InfoArray(21)) = CInt(i) Then%>
									        <option value="<%=i%>" selected><%=i%></option>
								        <%Else%>
									        <option value="<%=i%>"><%=i%></option>
								        <%End If%>
                                    <%Next%>
							    </select>
						    </td>
                        </tr>
                        <tr>
 						    <th>Dist Start-Finish:</th>
                            <td>
						        <input name="start_to_finish" id="start_to_finish" size="2" value="<%=InfoArray(19)%>">
                            </td>
 						    <th>Indiv/Relay:</th>
                            <td>
                                <select name="indiv_relay" id="indiv_relay">
								    <%If InfoArray(23) = "indiv" Then%>
									    <option value="indiv" selected>indiv</option>
									    <option value="relay">relay</option>
								    <%Else%>
									    <option value="indiv">indiv</option>
									    <option value="relay" selected>relay</option>
								    <%End If%>
                                </select>
                            </td>
 						    <th>Min Lap Time:</th>
                            <td>
						        <input name="min_lap" id="min_lap" size="2" value="<%=InfoArray(22)%>">
                            </td>
					    </tr>
                        <tr>
						    <th>Regist Link:</th>
						    <td colspan="5"><input type="text" name="online_reg_link" id="online_reg_link" value="<%=InfoArray(9)%>" size="60"></td>
                        </tr>
					    <tr>
						    <td style="color: red;text-align: center;background-color: #ececd8;" colspan="8">
							    <input type="checkbox" name="delete_race" id="delete_race">&nbsp;<span style="font-weight: bold;">Delete Race</span>
						    </td>
					    </tr>
					    <tr>
						    <td colspan="8">
							    <input type="hidden" name="submit_race_info" id="submit_race_info" value="submit_race_info">
							    <input type="submit" name="submit1" id="submit1" value="Make Changes">
						    </td>
					    </tr>
				    </table>
				    </form>
				
                    <%If InfoArray(2) = "9" Then%>
                        <h5 class="h5">Multi-Sport Leg Settings</h5>
                        <form role="form" class="form-horizontal" name="multi_settings" method="post" action="race_info.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
                        <div class="form-group">
                            <label for="num_legs" class="control-label col-xs-2">Num Legs:</label>
                            <div class="col-xs-10">
                                <input type="text" class="form-control" name="num_legs" id="num_legs" value="<%=iNumLegs%>">
                            </div>
                        </div>
			            <div class="form-group">
				            <label for="leg_1_name" class="control-label col-xs-2">Leg 1 Name:</label>
				            <div class="col-xs-4">
                                <input type="text" class="form-control" name="leg_1_name" id="leg_1_name" value="<%=sLeg1Name%>">
                            </div>
				            <label for="leg_1_dist" class="control-label col-xs-2">Leg 1 Dist:</label>
				            <div class="col-xs-4">
                                <input type="text" class="form-control" name="leg_1_dist" id="leg_1_dist" value="<%=sngLeg1Dist%>">
                            </div>
			            </div>
			            <div class="form-group">
				            <label for="leg_2_name" class="control-label col-xs-2">Leg 2 Name:</label>
				            <div class="col-xs-4">
                                <input type="text" class="form-control" name="leg_2_name" id="leg_2_name" value="<%=sLeg2Name%>">
                            </div>
				            <label for="leg_2_dist" class="control-label col-xs-2">Leg 2 Dist:</label>
				            <div class="col-xs-4">
                                <input type="text" class="form-control" name="leg_2_dist" id="leg_2_dist" value="<%=sngLeg2Dist%>">
                            </div>
			            </div>
			            <div class="form-group">
				            <label for="leg_3_name" class="control-label col-xs-2">Leg 3 Name:</label>
				            <div class="col-xs-4">
                                <input type="text" class="form-control" name="leg_3_name" id="leg_3_name" value="<%=sLeg3Name%>">
                            </div>
				            <label for="leg_3_dist" class="control-label col-xs-2">Leg 3 Dist:</label>
				            <div class="col-xs-4">
                                <input type="text" class="form-control" name="leg_3_dist" id="leg_3_dist" value="<%=sngLeg3Dist%>">
                            </div>
			            </div>
                        <div class="form-group">
                            <input type="hidden" name="submit_multi" id="submit_multi" value="submit_multi">
                            <input type="submit" class="form-control" name="submitMulti" id="submitMulti" value="Save Changes">
                        </div>
                        </form>
                    <%End If%>

				    <table style="margin-left:10px;">
                        <tr>
					        <td style="padding-right:25px;" valign="top">
						        <h4 class="h4">Edit Age Groups</h4>
						        <form name="update_age_grps" method="post" action="race_info.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
						        <table style="background-color:#ececec;">
							        <tr>
								        <td>
									        <table class="table table-striped">
										        <tr>
											        <th colspan="4">Male Age Groups</th>
										        </tr>
										        <tr>
											        <td>End Age</td>
											        <td>Awds</td>
                                                    <td>Name</td>
                                                    <td>Delete?</td>
										        </tr>
										        <%If UBound(MaleArray, 2) > 1 Then%>
											        <%For i = 0 to UBound(MaleArray, 2) - 1%>
												        <tr>
													        <td>
														        <% If MaleArray(1, i) = "110" Then%>
															        <input type="text" name="m_end_age_<%=MaleArray(0, i)%>" 
															                  id="m_end_age_<%=MaleArray(0, i)%>" size="1"
																	          maxlength="2" value="<%=MaleArray(1, i)%>" disabled>
														        <%Else%>
															        <input type="text" name="m_end_age_<%=MaleArray(0, i)%>" 
															                  id="m_end_age_<%=MaleArray(0, i)%>" size="1"
																	          maxlength="2" value="<%=MaleArray(1, i)%>">
														        <%End If%>
													        </td>
													        <td>
														        <input type="text" name="m_awds_<%=MaleArray(0, i)%>" id="m_awds_<%=MaleArray(0, i)%>" size="1"
																          maxlength="2" value="<%=MaleArray(2, i)%>">
													        </td>
													        <td>
														        <input type="text" name="m_name_<%=MaleArray(0, i)%>" id="m_name_<%=MaleArray(0, i)%>"
																    value="<%=MaleArray(3,i)%>">
													        </td>
													        <td>
														        <% If MaleArray(1, i) = "110" Then%>
														            <input type="checkbox" name="delete_<%=MaleArray(0, i)%>" 
                                                                        id="delete_<%=MaleArray(0, i)%>" disabled>

														        <%Else%>
														            <input type="checkbox" name="delete_<%=MaleArray(0, i)%>" 
                                                                        id="delete_<%=MaleArray(0, i)%>">

														        <%End If%>
													        </td>
												        </tr>
											        <%Next%>
                                                <%Else%>
												    <tr>
													    <td>
												            <input type="text" name="m_end_age_<%=MaleArray(0, 0)%>" id="m_end_age_<%=MaleArray(0, 0)%>" 
                                                                size="1" maxlength="2" value="<%=MaleArray(1, 0)%>">
													    </td>
													    <td>
														    <input type="text" name="m_awds_<%=MaleArray(0, 0)%>" id="m_awds_<%=MaleArray(0, 0)%>" size="1"
																        maxlength="2" value="<%=MaleArray(2,0)%>">
													    </td>
													    <td>
														    <input type="text" name="m_name_<%=MaleArray(0, 0)%>" id="m_name_<%=MaleArray(0, 0)%>"
																value="<%=MaleArray(3,0)%>">
													    </td>
													    <td>
														    <input type="checkbox" name="delete_<%=MaleArray(0, 0)%>" id="delete_<%=MaleArray(0, 0)%>">
													    </td>
												    </tr>
										        <%End If%>
										        <tr><th colspan="3">Add Male Age Group:</th></tr>
										        <tr>
											        <td>
												        <select name="new_m_end_age" id="new_m_end_age">
													        <%For j = 0 To 110%>
														        <option value="<%=j%>"><%=j%></option>
													        <%Next%>
												        </select>
											        </td>
											        <td colspan="2">
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
									        <table class="table table-striped">
										        <tr>
											        <th colspan="4">Female Age Groups</th>
										        </tr>
										        <tr>
											        <td>End Age</td>
											        <td>Awds</td>
                                                    <td>Name</td>
                                                    <td>Delete?</td>
										        </tr>
										        <%If UBound(FemaleArray, 2) > 1 Then%>
											        <%For i = 0 to UBound(FemaleArray, 2) - 1%>
												        <tr>
													        <td>
														        <%If FemaleArray(1, i) = "110" Then%>
															        <input type="text" name="f_end_age_<%=FemaleArray(0, i)%>" 
															                  id="f_end_age_<%=FemaleArray(0, i)%>" size="1"
																	          maxlength="2" value="<%=FemaleArray(1, i)%>" disabled>
														        <%Else%>
															        <input type="text" name="f_end_age_<%=FemaleArray(0, i)%>" 
															                  id="f_end_age_<%=FemaleArray(0, i)%>" size="1"
																	          maxlength="2" value="<%=FemaleArray(1, i)%>">
														        <%End If%>
													        </td>
													        <td>
														        <input type="text" name="f_awds_<%=FemaleArray(0, i)%>" id="f_awds_<%=FemaleArray(0, i)%>" size="1"
																          maxlength="2" value="<%=FemaleArray(2, i)%>">
													        </td>
													        <td>
														        <input type="text" name="f_name_<%=FemaleArray(0, i)%>" id="f_name_<%=FemaleArray(0, i)%>"
																    value="<%=FemaleArray(3,i)%>">
													        </td>
													        <td>
														        <%If FemaleArray(1, i) = "110" Then%>
														            <input type="checkbox" name="delete_<%=FemaleArray(0, i)%>" 
                                                                        id="delete_<%=FemaleArray(0, i)%>" disabled>
														        <%Else%>
														            <input type="checkbox" name="delete_<%=FemaleArray(0, i)%>" 
                                                                        id="delete_<%=FemaleArray(0, i)%>">
														        <%End If%>
													        </td>
												        </tr>
											        <%Next%>
                                                <%Else%>
												    <tr>
													    <td>
												            <input type="text" name="f_end_age_<%=FemaleArray(0, 0)%>" id="f_end_age_<%=FemaleArray(0, 0)%>" 
                                                                size="1" maxlength="2" value="<%=FemaleArray(1, 0)%>">
													    </td>
													    <td>
														    <input type="text" name="f_awds_<%=FemaleArray(0, 0)%>" id="f_awds_<%=FemaleArray(0, 0)%>" size="1"
																        maxlength="2" value="<%=FemaleArray(2, 0)%>">
													    </td>
													    <td>
														    <input type="text" name="f_name_<%=FemaleArray(0, 0)%>" id="f_name_<%=FemaleArray(0, 0)%>"
																        value="<%=FemaleArray(3,0)%>">
													    </td>
													    <td>
														    <input type="checkbox" name="delete_<%=FemaleArray(0, 0)%>" id="delete_<%=FemaleArray(0, 0)%>">
													    </td>
												    </tr>
										        <%End If%>
										        <tr>
											        <th colspan="3">Add Female Age Group:</th>
										        </tr>
										        <tr>
											        <td>
												        <select name="new_f_end_age" id="new_f_end_age">
													        <%For j = 0 To 110%>
														        <option value="<%=j%>"><%=j%></option>
													        <%Next%>
												        </select>
											        </td>
											        <td colspan="2">
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
								        <td colspan="4">
									        <input type="hidden" name="submit_age_grps" id="submit_age_grps" value="submit_age_grps">
									        <input type="submit" name="submit2" id="submit2" value="Save Changes">
								        </td>
							        </tr>
						        </table>
						        </form>
					        </td>
                        </tr>
				    </table>
			    <%End If%>
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