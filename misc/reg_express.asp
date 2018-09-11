<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lRaceID, lPartID, lEventID
Dim sEventname, sRace, sFirstName, sLastName, sAddress, sCity, sState, sZip, sPhone, sEmail, sGender, sBib, sWaiver
Dim iLeapYrs, iAgeDays, iAge
Dim dEventDate
Dim RaceArray()
Dim bDone

bDone = False

lEventID = Request.QueryString("event_id")
lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_part") = "submit_part" Then
	sFirstName = Replace(Request.Form.Item("first_name"),"'","''")
	sLastName = Replace(Request.Form.Item("last_name"),"'","''")
	sGender = Request.Form.Item("gender")
	sAddress = Replace(Request.Form.Item("address"),"'","''")
	sCity = Replace(Request.Form.Item("city"),"'","''")
	sState = Trim(Request.Form.Item("state"))
	sZip = Request.Form.Item("zip")
	sPhone = Request.Form.Item("phone")
	sEmail = Request.Form.Item("email")
	iAge = Request.Form.Item("age")
	
	'first see if they exist in the db and if so update their data
	sql = "INSERT INTO Participant (FirstName, LastName, Gender, Address, City, St, Zip, Phone, Email)"
	sql = sql & " VALUES ('" & sFirstName & "', '"  & sLastName & "', '" & sGender & "', '" & sAddress & "', '" & sCity 
	sql = sql & "', '" & sState & "', '" & sZip & "', '" & sPhone & "', '" & sEmail & "')"
	Set rs=conn.Execute(sql)
	Set rs=Nothing

	'get participant id
	sql = "SELECT ParticipantID FROM Participant WHERE FirstName='" & sFirstName & "' AND LastName='" & sLastName  
	sql = sql & "' AND Address = '" & sAddress & "' AND Gender = '" & sGender & "' AND City = '" & sCity & "' AND St = '" 
	sql = sql & sState & "' AND Zip = '" & sZip & "' AND Phone = '" & sPhone & "' AND Email = '" & sEmail & "' ORDER BY ParticipantID DESC"
	Set rs = conn.Execute(sql)
	lPartID = rs(0).Value
	Set rs=Nothing
	
	'insert into part reg table
	sql = "INSERT INTO PartReg (ParticipantID, WhereReg, DateReg, RaceID)"
	sql = sql & " VALUES (" & lPartID & ", 'Express', '" & Date & "', " & lRaceID & ")"
	Set rs=conn.Execute(sql)
	Set rs=Nothing

    'assign bib here

	'insert into part race table
	sql = "INSERT INTO PartRace (ParticipantID, Age, RaceID, Bib) VALUES (" & lPartID & ", " & lRaceID & ", '" & sBib & "')"
	Set rs=conn.Execute(sql)
	Set rs=Nothing

    bDone = True
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
    
'get waiver
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Waiver FROM Waiver WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sWaiver = rs(0).Value
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
%>
<!DOCTYPE html>
<html>
<head>

<title>GSE Express Registration</title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">
<meta name="description" content="GSE Express Registration.">

<link rel="icon" href="favicon.ico" type="image/x-icon"> 
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon"> 


 




<style type="text/css">
	th{
		text-align:right;
		white-space:nowrap;
		padding:5px 0 0 5px;
	}
	
	td{
		padding-top:5px;
	}
</style>

<script>
function checkFields() {
 	if (document.part_reg.first_name.value == '' || 
 	    document.part_reg.last_name.value == '' ||
	 	document.part_reg.gender.value == ''|| 
	 	document.part_reg.age.value == '')
		{
  		alert('Please fill in all required fields-they are denoted with a red asterisk!');
  		return false
  		}
	else
   		return true
}
</script>
</head>

<body onload="document.part_reg.first_name.focus();">
<%If bDone = True Then%>
    <p style="font-size: 1.5em;">Congratulations!  Your registration process is completed.  PLEASE SHOW THE REGISTRATION OFFICIAL THIS SCREEN!  The will give
    you the bib number listed below, collect your fee and make sure that you receive any other materials needed.  Have a GREAT RACE!</p>

    <div style="text-align: center;font-weight: bold;font-size: 3.0em;"><%=iBib%></div>
<%Else%>
    <table style="font-size:1.0em;">
        <tr>
            <th><h3><%=sEventName%> Express Race Day Registration</h3></th>
        </tr>
	    <tr>
            <td>
	            <%If UBound(RaceArray, 2) > 1 Then%>
		            <form name="get_race" method="post" action="reg_express.asp?event_id=<%=lEventID%>">
			        <span style="font-weight:bold;">Select Race:</span>
			        <select name="races" id="races" onchange="this.form.get_race.click()">
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
			        <input type="submit" name="get_race" id="get_race" value="Select Race">
		            </form>
	            <%End If%>
            </td>
        </tr>
	    <tr>
            <td>
                <%If Not CLng(lRaceID) = 0 Then%>
			        <form name="part_reg" method="post" action="reg_express.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>"
                    onSubmit="return checkFields();">
			        <table>
			            <tr>
					        <th><span style="color:#d62002">*</span>First Name:</th>
					        <td><input name="first_name" id="first_name" size="9" maxLength="30"></td>
					        <th<span style="color:#d62002">*</span>Last Name:</th>
					        <td><input name="last_name" id="last_name" maxLength="30" size="15"></td>
				        </tr>
				        <tr>
					        <th><span style="color:#d62002">*</span>Gender:</th>
					        <td>
						        <select name="gender" id="gender"> 
							        <option value="">&nbsp</option> 
							        <option value="M">Male</option> 
							        <option value="F">Female</option>
						        </select>
					        </td>
					        <th>Age:</th>
					        <td><input name="age" id="age" size="3" maxlength="2"></td>
				        </tr>
				        <tr>
					        <th>Phone:</th>
					        <td><input name="phone" id="phone" size="12" maxlength="12"></td>
					        <th>Email:</th>
					        <td><input name="email" id="email" maxLength="50" size="30"></td>				
                        </tr>
				        <tr>
					        <th>Address:</th>
					        <td><input name="address" id="address" maxLength="50" size="30"></td>
					        <th>City:</th>
					        <td><input name="city" id="city" maxLength="30" size="15"></td>
				        </tr>
				        <tr>
					        <th><span style="color:#d62002">*</span>St:</th>
					        <td><input name="state" id="state" maxLength="2" size="2"></td>
					        <th>Postal:</th>
					        <td<input name="zip" id="zip" maxLength="5" size="7"></td>
				        </tr>
				        <tr>
					        <td colspan="4">
                                <p><span style="font-weight: bold;">NOTE:  By submitting this form you are agreeing to the terms
                                of the waiver below!</span>  Local event administration may ask you to sign a hard copy waiver as well.</p>

                                <p><%=sWaiver%></p>
                            </td>
				        </tr>
				        <tr>
					        <td colspan="4">
						        <input type="hidden" name="submit_part" id="submit_part" value="submit_part">
						        <input type="submit" name="submit1" id="submit1" value="Enter Participant"> 
					        </td>
			            </tr>
			        </table>
			        </form>
                <%End If%>
            </td>
        </tr>	
    </table>
<%End If%>	
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>