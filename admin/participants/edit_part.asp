<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2, conn2
Dim i, j, k
Dim lRaceID, lPartID, lEventID, lProvider
Dim sEventName, sRace, sMobileNum
Dim dEventDate
Dim PartReg(10), RaceArray(), PartArray(), TempArray(1), CellProviders
Dim sErrMsg

lEventID = Request.QueryString("event_id")

lPartID = Request.QueryString("part_id")
If Not IsNumeric(lPartID) Then Response.Redirect "htttp://www.google.com"

lRaceID = Request.QueryString("race_id")
If Not IsNumeric(lRaceID) Then Response.Redirect "htttp://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn2, 1, 2
CellProviders = rs.GetRows()
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

'get race information
i = 0
ReDim RaceArray(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	RaceArray(0, i) = rs(0).Value
	RaceArray(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve RaceArray(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_part") = "submit_part" Then
	lPartID = Request.Form.Item("participants")
	
	If Not CLng(lPartID) = vbNullString Then
		'get this race id
		For i = 0 To UBound(RaceArray, 2) - 1
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql="SELECT RaceID FROM PartRace WHERE ParticipantID = " & lPartID & " AND RaceID = " & RaceArray(0, i)
			rs.Open sql, conn, 1, 2
			If rs.RecordCount > 0 Then
				lRaceID = rs(0).Value
				Exit For
			Else
				lRaceID = 0
			End If
			rs.Close
			Set rs=Nothing
		Next
	End If
ElseIf Request.Form.Item("submit_race_change") = "submit_race_change" Then
    'update partreg table
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM PartReg WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("races")
    rs.Update
    rs.Close
    Set rs = Nothing

    'update partrace table
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM PartRace WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("races")
    rs.Update
    rs.Close
    Set rs = Nothing

    'update indresults table if necessary
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM IndResults WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs(0).Value = Request.Form.Item("races")
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
	
    lRaceID = Request.Form.Item("races")

    'adjust age groups
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT p.Gender, pr.Age, pr.AgeGrp FROM PartRace pr INNER JOIN Participant p ON "
    sql = sql & "pr.ParticipantID = p.ParticipantID WHERE p.ParticipantID = " & lPartID
    rs.Open sql, conn, 1, 2
    rs(2).Value = GetAgeGrp(rs(0).Value, rs(1).Value, Request.Form.Item("races"))
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("delete_part") = "delete_part" Then
	sql = "DELETE FROM PartRace WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	Set rs = Nothing
	
	sql = "DELETE FROM PartReg WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	Set rs = Nothing
	 
	sql = "DELETE FROM IndResults WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	Set rs = Nothing
	
	lRaceID = 0
	lPartID = 0
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
	PartReg(0) = Replace(Request.Form.Item("first_name"), "'", "''")
	PartReg(1) = Replace(Request.Form.Item("last_name"), "'", "''")
	If Not Request.Form.Item("city") & "" = "" Then PartReg(2) = Replace(Request.Form.Item("city"), "'", "''")
	PartReg(3) = Request.Form.Item("state")
	PartReg(4) = Request.Form.Item("phone")
	PartReg(5) = Request.Form.Item("email")
	PartReg(6) = Request.Form.Item("dob_month") & "/" & Request.Form.Item("dob_day") & "/" &Request.Form.Item("dob_year")
	PartReg(7) = Request.Form.Item("gender")
	PartReg(8) = Request.Form.Item("fbook")
	PartReg(9) = Request.Form.Item("twitter")
  	PartReg(10) = Request.Form.Item("age")
	
    sMobileNum = Request.Form.Item("mobile_num")
    If Not sMobileNum = vbNullString Then
        sMobileNum = Replace(sMobileNum, "-", "")
        sMobileNum = Replace(sMobileNum, ".", "")
        sMobileNum = Replace(sMobileNum, "(", "")
        sMobileNum = Replace(sMobileNum, ")", "")
        sMobileNum = Replace(sMobileNum, " ", "")
        sMobileNum = Trim(sMobileNum)
    End If

    lProvider = Request.Form.item("provider")

	If Not IsDate(PartReg(6)) Then
		sErrMsg="The date of birth supplied is not a valid date.  Please correct it and resubmit this information."
	Else
		'update participant table
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT FirstName, LastName, City, St, Phone, Email, DOB, Gender, FBook, Twitter FROM Participant "
		sql = sql & "WHERE ParticipantID = " & lPartID
		rs.Open sql, conn, 1, 2
		
		If PartReg(0) = vbNullString Then
			rs(0).Value = rs(0).OriginalValue
		Else
			rs(0).Value = PartReg(0)
		End If
		
		If PartReg(1) = vbNullString Then
			rs(1).Value = rs(1).OriginalValue
		Else
			rs(1).Value = PartReg(1)
		End If
		
		For i = 2 to 9
			rs(i).Value = PartReg(i)
		Next
		
		rs.Update
		rs.Close
		Set rs = Nothing
	End If

    'update partrace table
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Age FROM PartRace WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    rs(0).Value = PartReg(10)
    rs.Update
    rs.Close
    Set rs = Nothing

    'adjust age groups
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT p.Gender, pr.Age, pr.AgeGrp FROM PartRace pr INNER JOIN Participant p ON "
    sql = sql & "pr.ParticipantID = p.ParticipantID WHERE p.ParticipantID = " & lPartID & " AND pr.RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    rs(2).Value = GetAgeGrp(rs(0).Value, rs(1).Value, lRaceID)
    rs.Update
    rs.Close
    Set rs = Nothing

    'edit sms
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MobileNumber, CellProvider FROM MobileSettings WHERE EventID = " & lEventID & " AND PartID = " & lPartID
    rs.Open sql, conn, 1, 2
    If rs.REcordCount > 0 Then
        rs(0).Value = sMobileNum
        rs(1).Value = lProvider
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
End If

If CStr(lPartID) = vbNullString Then lPartID = 0
If CStr(lRaceID) = vbNullString Then lRaceID = 0

i = 0
ReDim Preserve PartArray(1, 0)			
For k = 0 to UBound(RaceArray, 2) - 1
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="SELECT p.ParticipantID, p.FirstName, p.LastName, rg.RaceID FROM "
	sql = sql & "Participant p INNER JOIN PartReg rg ON p.ParticipantID = rg.ParticipantID JOIN PartRace rc "
	sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE (rc.RaceID = " & RaceArray(0, k) & " AND rg.RaceID = " 
	sql = sql & RaceArray(0, k) & ") ORDER BY p.LastName, p.FirstName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		PartArray(0, i) = rs(0).value
		PartArray(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & RaceName(rs(3).Value) & ")"
		i = i + 1
		ReDim Preserve PartArray(1, i)
		rs.MoveNext
	Loop
	rs.Close
	Set rs=Nothing
Next
	
'sort the array
For i = 0 to UBound(PartArray, 2) - 2
	For j = i + 1 to UBound(PartArray, 2) - 1
		If CStr(UCase(PartArray(1, i))) > CStr(UCase(PartArray(1, j))) Then
			For k = 0 to 1
				TempArray(k) = PartArray(k, i)
				PartArray(k, i) = PartArray(k, j)
				PartArray(k, j) = TempArray(k)
			Next
		End If
	Next
Next

i = 0
'get participant data
If Not CLng(lPartID) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, City, St, Phone, Email, DOB, Gender, Twitter, FBook FROM Participant WHERE ParticipantID = " & lPartID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then 
		For i = 0 to 9
			If Not rs(i).Value & "" = "" Then PartReg(i) = rs(i).Value
		Next
	End If
	rs.Close
	Set rs = Nothing

    Call GetMyMobile(lPartID)
End If

'get part reg data
If Not (CLng(lPartID) = 0 Or CLng(lRaceID)) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Age FROM PartRace WHERE ParticipantID = " & lPartID & " AND RaceID = " & lRaceID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then 
		PartReg(10) = rs(0).Value
	End If
	rs.Close
	Set rs = Nothing
End If

Private Function GetAgeGrp(sMF, iAge, lThisRace)
    Dim sql_agegrp, rs_agegrp
    Dim iBegAge, iEndAge
    
    iBegAge = 0
    
    sql_agegrp = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & LCase(sMF) & "' AND RaceID = " & lThisRace & " ORDER BY EndAge DESC"
    Set rs_agegrp = conn.Execute(sql_agegrp)
    Do While Not rs_agegrp.EOF
        If CInt(iAge) <= CInt(rs_agegrp(0).Value) Then
            iEndAge = rs_agegrp(0).Value
        Else
            iBegAge = CInt(rs_agegrp(0).Value) + 1
            Exit Do
        End If
        rs_agegrp.MoveNext
    Loop
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

Private Function RaceName(lThisRace)
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RaceName FROM RaceData WHERE RaceID = " & lThisRace
	rs2.Open sql2, conn, 1, 2
	RaceName = Replace(rs2(0).Value, "''", "'")
	rs2.Close
	Set rs2 = Nothing
End Function

Private Sub GetMyMobile(lMyID)
    lProvider = 0
    sMobileNum = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT MobileNumber, CellProvider FROM MobileSettings WHERE EventID = " & lEventID & " AND PartID = " & lMyID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then 
        sMobileNum = rs2(0).Value
        lProvider = rs2(1).Value
    End If
    rs2.Close
    Set rs2 = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Participant Data</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
    <h4 class="h4">Edit Participants For <%=sEventName%> on <%=dEventDate%></h4>

    <form class="form-inline" name="get_part" method="post" action="edit_part.asp?event_id=<%=lEventID%>&amp;part_id=<%=lPartID%>&amp;race_id=<%=lRaceID%>"
          onsubmit="javascript:window.opener.location.reload();">
	<label for="participants">Select Participant To Edit:</label>
	<select class="form-control" name="participants" id="participants" onchange="this.form.get_part.click()">
		<option value="0">&nbsp;</option>
		<%For i = 0 to UBound(PartArray, 2) - 1%>
			<%If CLng(lPartID) = CLng(PartArray(0, i)) Then%>
				<option value="<%=PartArray(0, i)%>" selected><%=PartArray(1, i)%></option>
			<%Else%>
				<option value="<%=PartArray(0, i)%>"><%=PartArray(1, i)%></option>
			<%End If%>
		<%Next%>
	</select>
	<input type="hidden" class="form-control" name="submit_part" id="submit_part" value="submit_part">
	<input type="submit" class="form-control" name="get_part" id="get_part" value="Edit This Participant">
    </form>

    <hr>

    <%If Not sErrMsg = vbNullString Then%>
	    <p class="bg-danger text-danger"><%=sErrMsg%></p>
    <%End If%>

    <%If Not (CLng(lPartID) = 0 Or CLng(lRaceID)) = 0 Then%>
	    <form class="form" name="edit_part" method="post" action="edit_part.asp?race_id=<%=lRaceID%>&amp;part_id=<%=lPartID%>&amp;event_id=<%=lEventID%>">
		<div class="form-group">
			<label for="first_name" class="control-label col-xs-2">First:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="first_name" id="first_name"  maxlength="25" value="<%=PartReg(0)%>">
            </div>
			<label for="last_name" class="control-label col-xs-2">Last:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="last_name" id="last_name" maxlength="25" value="<%=PartReg(1)%>">
            </div>
		</div>
		<div class="form-group">
			<label for="gender" class="control-label col-xs-2">Gender:</label>
			<div class="col-xs-4">
				<select class="form-control" name="gender" id="gender">
					<%If PartReg(7) = "M" Then%>
						<option value="M" selected>M</option>
					<%Else%>
						<option value="M">M</option>
					<%End If%>
					<%If PartReg(7) = "F" Then%>
						<option value="F" selected>F</option>
					<%Else%>
						<option value="F">F</option>
					<%End If%>
					<%If PartReg(7) = "X" Then%>
						<option value="X" selected>X</option>
					<%Else%>
						<option value="X">X</option>
					<%End If%>
				</select>
			</div>
			<label for="phone" class="control-label col-xs-2">Phone:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="phone" id="phone" maxlength="12" value="<%=PartReg(4)%>">
            </div>
		</div>
		<div class="form-group">
			<label for="city" class="control-label col-xs-2">City:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="city" id="city" maxlength="25" value="<%=PartReg(2)%>">
            </div>
			<label for="state" class="control-label col-xs-2">State:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="state" id="state" maxlength="3" value="<%=PartReg(3)%>">
            </div>
        </div>
        <div class="form-group">
			<label for="age" class="control-label col-xs-2">Age:</label>
			<div class="col-xs-4">
				<select class="form-control" name="age" id="age">
					<%For i = 0 To 99%>
						<%If Int(PartReg(10)) = i Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
			</div>
			<label for="email" class="control-label col-xs-2">Email:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="email" id="email" maxlength="50" value="<%=PartReg(5)%>">
            </div>
		</div>
		<div class="form-group">
			<label for="mobile_num" class="control-label col-xs-2">Cell:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="mobile_num" id="mobile_num" value="<%=sMobileNum%>">
            </div>
			<label for="provider" class="control-label col-xs-2">Provider:</label>
			<div class="col-xs-4">
				<select class="form-control" name="provider" id="provider">
                    <option value="0">&nbsp;</option>
					<%For i = 0 To UBound(CellProviders, 2)%>
						<%If CLng(CellProviders(0, i)) = CLng(lProvider) Then%>
							<option value="<%=CellProviders(0, i)%>" selected><%=CellProviders(1, i)%></option>
						<%Else%>
							<option value="<%=CellProviders(0, i)%>"><%=CellProviders(1, i)%></option>
						<%End If%>
					<%Next%>
				</select>
            </div>
		</div>
		<div class="form-group">
			<label for="fbook" class="control-label col-xs-2">FBook:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="fbook" id="fbook" value="<%=PartReg(8)%>">
            </div>
			<label for="twitter" class="control-label col-xs-2">Twitter:</label>
			<div class="col-xs-4">
                <input type="text" class="form-control" name="twitter" id="twitter" value="<%=PartReg(9)%>">
            </div>
		</div>
		<div class="form-group">
			<label for="dob_month" class="control-label col-xs-2">DOB:</label>
			<div class="col-xs-2">
				<select class="form-control" name="dob_month" id="dob_month">
					<%For i = 1 to 12%>
						<%If CInt(Month(CDate(PartReg(6)))) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
            </div>
            <label for="dob_day" class="control-label col-xs-1">/</label>
            <div class="col-xs-2">
				<select class="form-control" name="dob_day" id="dob_day">
					<%For i = 1 to 31%>
						<%If CInt(Day(CDate(PartReg(6)))) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
            </div>
            <label for="dob_year" class="control-label col-xs-1">/</label>
            <div class="col-xs-2">
				<select class="form-control" name="dob_year" id="dob_year">
					<%For i = 2005 to 1900 Step -1%>
						<%If CInt(Year(CDate(PartReg(6)))) = CInt(i) Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
			</div>
            <div class="col-xs-2">
				&nbsp;
			</div>
		</div>
        <div class="form-group">
    		<input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
			<input type="submit" class="form-control" name="submit1" id="submit1" value="Save Changes">
		</div>
	    </form>
	
        <hr>

	    <%If UBound(RaceArray, 2) > 1 Then%>
		    <form class="form-inline" name="change_race" method="Post" action="edit_part.asp?race_id=<%=lRaceID%>&amp;part_id=<%=lPartID%>&amp;event_id=<%=lEventID%>">
		    <label for="races">Switch Races</label>
		    <select class="form-control" name="races" id="races">
			    <%For i = 0 to UBound(RaceArray, 2) - 1%>
				    <%If CLng(lRaceID) = CLng(RaceArray(0, i)) Then%>
					    <option value="<%=RaceArray(0, i)%>" selected><%=RaceArray(1, i)%></option>
				    <%Else%>
					    <option value="<%=RaceArray(0, i)%>"><%=RaceArray(1, i)%></option>
				    <%End If%>
			    <%Next%>
		    </select>
		    <input type="hidden" name="submit_race_change" id="submit_race_change" value="submit_race_change">
		    <input type="submit" class="form-control" name="get_race" id="get_race" value="Switch To This Race">
		    </form>
	    <%End If%>
	
	    <div class="bg-danger text-danger">
            You may use the following button to delete a participant from this race.  It will not delete them from other races in this event but it will
	        delete all of their data in this event, including results if the event has already happened.  There is no undo for this action.
	        <form class="form-inline" name="delete_part" method="Post" action="edit_part.asp?race_id=<%=lRaceID%>&amp;part_id=<%=lPartID%>&amp;event_id=<%=lEventID%>">
		    <input type="hidden" name="delete_part" id="delete_part" value="delete_part">
		    <input type="submit" class="form-control" name="submit3" id="submit3" value="Delete This Participant">
	        </form>
        </div>
    <%End If%>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>