<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lRaceID, lPartID, lEventID
Dim sEventname, sRace, sFirstName, sLastName, sCity, sState, sPhone, sEmail, sGender, sBib, sTwitter, sFbook, sShrtSize
Dim iLeapYrs, iAgeDays, iAge
Dim dEventDate, dDOB
Dim RaceArray()

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_part") = "submit_part" Then
	lRaceID = Request.Form.Item("races")

	'get event and race info
	sql = "SELECT ed.EventName, ed.EventDate, rd.Dist FROM Events ed INNER JOIN RaceData rd "
	sql = sql & "ON ed.EventID = rd.EventID WHERE ed.EventID = " & lEventID & " AND rd.RaceID = " & lRaceID
	Set rs = conn.Execute(sql)
	sEventName = rs(0).Value
	dEventDate = rs(1).Value
	sRace = rs(2).Value
	Set rs = Nothing
	
	sFirstName = Replace(Request.Form.Item("first_name"),"'","''")
	sLastName = Replace(Request.Form.Item("last_name"),"'","''")
	sGender = Request.Form.Item("gender")
	sCity = Replace(Request.Form.Item("city"),"'","''")
	sState = Trim(Request.Form.Item("state"))
	sPhone = Request.Form.Item("phone1") & "-" & Request.Form.Item("phone2") & "-"  & Request.Form.Item("phone3")
	sEmail = Request.Form.Item("email")
	dDOB = Request.Form.Item("dob_month") & "/" & Request.Form.Item("dob_day") & "/" & Request.Form.Item("dob_year")
	sBib = Request.Form.Item("bib")
	sTwitter = Request.Form.Item("twitter")
	sFbook = Request.Form.Item("fbook")
    sShrtSize = Request.Form.Item("shrt_size")
	If dDOB = "//" Then
		dDOB = "1/1/1900"
	End If
	
	If Request.Form.Item("age") = vbNullString Then
		iAge = RaceDayAge()
	Else
		iAge = Request.Form.Item("age")
	End If
    
    'check for a data match
    lPartID = 0
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM Participant WHERE FirstName = '" & sFirstName & "' AND LastName = '" & sLastName
    sql = sql & "' AND City = '" & sCity & "' AND Email = '" & sEmail & "' AND Gender = '" & sGender & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then lPartID = rs(0).Value
    rs.Close
    Set rs = Nothing
    
    'insert into the participant table
    If CLng(lPartID) = 0 Then
        sql = "INSERT INTO Participant (FirstName, LastName, City, St, Phone, Email, DOB, Gender, Twitter, Fbook)VALUES ('"
        sql = sql & sFirstName & "', '" & sLastName & "', '" & sCity & "', '" & sState & "', '" & sPhone & "', '" & sEmail
        sql = sql & "', '" & dDOB & "', '" & sGender & "', '" & sTwitter & "', '" & sFbook & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
        
        'get ParticipantID
        sql = "SELECT ParticipantID FROM Participant WHERE (FirstName = '" & sFirstName & "' AND LastName = '"
        sql = sql & sLastName & "' AND City = '" & sCity & "' AND St = '" & sState & "' AND Phone = '" & sPhone & "' AND Email = '"
        sql = sql & sEmail & "' AND Gender = '" & sGender & "')"
        Set rs = conn.Execute(sql)
        lPartID = rs(0).Value
        Set rs = Nothing
    End If
	
	'insert into part reg table
	sql = "INSERT INTO PartReg (ParticipantID, WhereReg, DateReg, RaceID, ShrtSize) VALUES (" & lPartID & ", 'Mail In', '" & Date & "', " & lRaceID 
    sql = sql & ", '" & sShrtSize & "')"
	Set rs=conn.Execute(sql)
	Set rs=Nothing

	'insert into part race table
	sql = "INSERT INTO PartRace (ParticipantID, Age, RaceID, Bib, AgeGrp) VALUES (" & lPartID & ", " 
	sql = sql & CInt(iAge) & ", " & lRaceID & ", '" & sBib & "', '" & GetAgeGrp(sGender, iAge, lRaceID) & "')"
	Set rs=conn.Execute(sql)
	Set rs=Nothing

	'get event and race info
	sql = "SELECT ed.EventName, ed.EventDate, rd.Dist FROM Events ed INNER JOIN RaceData rd "
	sql = sql & "ON ed.EventID = rd.EventID WHERE ed.EventID = " & lEventID & " AND rd.RaceID = " & lRaceID

	Set rs = conn.Execute(sql)
	sEventName = rs(0).Value
	dEventDate = rs(1).Value
	sRace = rs(2).Value
	Set rs = Nothing
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
	RaceArray(0, i) = rs(0).value
	RaceArray(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve RaceArray(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If UBound(RaceArray, 2) = 1 Then lRaceID = RaceArray(0, 0)

'get participants age on race day
Function RaceDayAge()
    'first get leap years
    iLeapYrs = 0
    For i = Year(CDate(dEventDate)) To Year(CDate(dDOB)) Step -1
        If i / 4 = i \ 4 Then iLeapYrs = iLeapYrs + 1
    Next
        
    iAgeDays = DateDiff("d", CDate(dDOB), CDate(dEventDate))
    iAgeDays = iAgeDays - iLeapYrs
    RaceDayAge = iAgeDays \ 365
End Function

Public Function GetAgeGrp(sMF, iAge, lThisRaceID)
    Dim iBegAge, iEndAge
    
    iBegAge = 0
    
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sMF & "' AND RaceID = " & lThisRaceID & " ORDER BY EndAge DESC"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        If CInt(iAge) <= CInt(rs2(0).Value) Then
            iEndAge = rs2(0).Value
        Else
            iBegAge = CInt(rs2(0).Value) + 1
            Exit Do
        End If
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

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
<title>Enter <%=sEventName%> Participants</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
	th{
		text-align:right;
		white-space:nowrap;
		padding:5px 0 0 5px;
	}
	
	td{
		padding-top:5px;
	}
	
	textarea{
		font-size:1.0em;
	}
	
	table{
		padding-top:10px;
	}
</style>

<script>
function checkFields() {
 	if (document.part_reg.first_name.value == '' || 
 	    document.part_reg.last_name.value == '' ||
	 	document.part_reg.gender.value == ''|| 
        document.part_reg.races.value == ''||
	 	(document.part_reg.age.value == ''&&
	 	(document.part_reg.dob_month.value == '' || 
	 	document.part_reg.dob_day.value == '' || 
	 	document.part_reg.dob_year.value == '')))
		{
  		alert('Please fill in all required fields-they are denoted with a red asterisk!');
  		return false
  		}
 	else
		if (isNaN(document.part_reg.bib.value))
    		{
			alert('The bib number field can not contain non-numeric values');
			return false
			}
 	else
		if (isNaN(document.part_reg.dob_month.value) ||
		   isNaN(document.part_reg.dob_day.value) ||
		   isNaN(document.part_reg.dob_year.value) ||
		   isNaN(document.part_reg.age.value))
    		{
			alert('All date of birth/age fields must be numeric values');
			return false
			} 	
	else
		if (document.part_reg.phone1.value != '')
			{
			if(isNaN(document.part_reg.phone1.value))
				{
				alert('Phone number fields must be numeric values');
				return false
				}
			else
				if (document.part_reg.phone1.value.length != 3)
				{
				alert('The first two phone number fields must contain 3 numbers.');
				return false
				}	
			}
	else
		if (document.part_reg.phone2.value != '')
			{
			if(isNaN(document.part_reg.phone2.value))
				{
				alert('Phone number fields must be numeric values');
				return false
				}
			else
				if (document.part_reg.phone2.value.length != 3)
				{
				alert('The first two phone number fields must contain 3 numbers.');
				return false
				}	
			}
	else
		if (document.part_reg.phone3.value != '')
			{
			if(isNaN(document.part_reg.phone3.value))
				{
				alert('Phone number fields must be numeric values');
				return false
				}
			else
				if (document.part_reg.phone3.value.length != 4)
				{
				alert('The last phone number fields must contain 4 numbers.');
				return false
				}	
			}
	else
		if (isNaN(document.part_reg.bib.value))
    		{
			alert('The bib must be numeric!');
			return false
			} 
	else
		if (document.part_reg.state.value.length > 2)
		{
		alert('Please use 2 characters for the state.');
		return false
		}	
	else
		if (document.part_reg.dob_year.value != '' &&
		   document.part_reg.dob_year.value.length < 4)
		{
		alert('Please use 4 numbers for the year field in date of birth.');
		return false
		}	
	else
   		return true
}
</script>
</head>

<body onload="javascript:part_reg.first_name.focus()">
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    

			<h4 class="h4">Enter <%=sEventName%> Participants</h4>
			
			<!--#include file = "../../includes/event_nav.asp" -->

			<div style="margin: 0;padding: 0;font-size: 0.85em;">
				<a href="/admin/participants/part_data.asp?event_id=<%=lEventID%>" rel="nofollow">Participant Data</a>
            </div>
				
			<div style="text-align:right;font-size:0.85em;margin-right:10px;">
				<a href="javascript:pop('batch_upload/upload_template.xls',1024,750)">Upload Template</a>
				&nbsp;|&nbsp;
				<a href="javascript:pop('batch_upload/batch_upload.asp?event_id=<%=lEventID%>',1000,750)">Batch Upload Participants</a>
			</div>
				
			<form name="part_reg" method="post" action="enter_parts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" onSubmit="return checkFields();">
			  <table style="margin:10px;">
                <tr>
 					<th><span style="color:#d62002">*</span>Select Race:</th>
                    <td colspan="3">
					    <select name="races" id="races">
						    <option value="">&nbsp;</option>
						    <%For i = 0 to UBound(RaceArray, 2) - 1%>
							    <%If CLng(lRaceID) = CLng(RaceArray(0, i)) Then%>
								    <option value="<%=RaceArray(0, i)%>" selected><%=RaceArray(1, i)%></option>
							    <%Else%>
								    <option value="<%=RaceArray(0, i)%>"><%=RaceArray(1, i)%></option>
							    <%End If%>
						    <%Next%>
					    </select>
                    </td>
               </tr>
			    <tr>
					<th><span style="color:#d62002">*</span>First Name:</th>
					<td><input name="first_name" id="first_name" size="9" maxLength="30"></td>
					<th><span style="color:#d62002">*</span>Last Name:</th>
					<td><input name="last_name" id="last_name" maxLength="30" size="15"></td>
				</tr>
				<tr>
					<th><span style="color:#d62002">*</span>DOB:</th>
					<td>
						<input name="dob_month" id="dob_month" maxLength="2" size="1">&nbsp;- 
						<input name="dob_day" id="dob_day" maxLength="2" size="1">&nbsp;- 
						<input name="dob_year" id="dob_year" maxLength="4" size="2">
						<span style="color:#d62002;margin-left:35px;">Or *</span>
					</td>
					<th>Age:</th>
					<td><input name="age" id="age" style="WIDTH: 25px; Height:22px" maxlength="2"></td>
				</tr>
				<tr>
					<th><span style="color:#d62002">*</span>Gender:	</th>
					<td>
						<select name="gender" id="gender"> 
							<option value="">&nbsp</option> 
							<option value="M">M</option> 
							<option value="F">F</option>
						</select>
					</td>
					<th>Phone:</th>
					<td>
						<input name="phone1" id="phone1" maxLength="5" size="2">&nbsp;-
						<input name="phone2" id="phone2" maxLength="5" size="2">&nbsp;-
						<input name="phone3" id="phone3" maxLength="5" size="3">
					</td>
				</tr>
				<tr>
					<th><span style="color:#d62002">*</span>City:</th>
					<td><input name="city" id="city" maxLength="30"></td>
					<th><span style="color:#d62002;">*</span>State:</th>
                    <td><input name="state" id="state" maxLength="3" size="3"></td>
				</tr>
				<tr>
					<th>Email:</th>
					<td><input name="email" id="email" maxLength="50" size="25" onKeyUp="chkStr(this)"></td>
					<th>Bib:</th>
					<td><input name="bib" id="bib" size="5" maxlength="4"></td>
				</tr>
			    <tr>
					<th>Shrt Size:</th>
					<td colspan="3"><input name="shrt_size" id="shrt_size" size="3" maxLength="4"></td>
				</tr>
			    <tr>
					<th>Twitter:</th>
					<td><input name="twitter" id="twitter" size="15" maxLength="30"></td>
					<th>FBook:</th>
					<td><input name="fbook" id="fbook" maxLength="30" size="15"></td>
				</tr>
				<tr>
					<td colspan="4">
						<input type="hidden" name="submit_part" id="submit_part" value="submit_part">
						<input type="submit" name="submit1" id="submit1" value="Enter Participant"> 
					</td>
			    </tr>
			</table>
			</form>
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