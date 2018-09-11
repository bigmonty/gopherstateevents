<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID, lRaceID
Dim iAge
Dim sEventName, sRaceName, sWaiver, sSendInfo, sFirstName, sLastName, sGender, sAddress, sCity, sState, sZip, sPhone, sEmail, sComments
Dim dEventDate
Dim Races()
Dim bRegCompl

bRegCompl = False

lEventID = Request.QueryString("event_id")
If Not IsNumeric(lEventID) Then Response.Redirect "htttp://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

ReDim Races(1, 0)

If Not CStr(lEventID) & "" = "" Then
	'get event information
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	sEventName = rs(0).Value
	dEventDate = rs(1).Value
	rs.Close
	Set rs = Nothing
	
	i = 0
	sql = "SELECT RaceID, Dist FROM RaceData WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		Races(0, i) = rs(0).Value
		Races(1, i) = rs(1).Value
		i = i + 1
		ReDim Preserve Races(1, i)
		rs.MoveNext
	Loop
	Set rs = Nothing

	If UBound(Races, 2) = 1 Then lRaceID = Races(0, 0)
End If
	
If Request.Form.Item("submit_reg") = "submit_reg" Then
	sFirstName = Trim(Replace(Request.Form.Item("first_name"), "'", "''"))
	sLastName = Trim(Replace(Request.Form.Item("last_name"), "'", "''"))
	sGender = Request.Form.Item("gender")
	sAddress = Trim(Replace(Request.Form.Item("address"), "'", "''"))
	sCity = Trim(Replace(Request.Form.Item("city"), "'", "''"))
	sState = Trim(Request.Form.Item("state"))
	sZip = Request.Form.Item("zip")
	sPhone = Request.Form.Item("phone1") & "-" & Request.Form.Item("phone2") & "-" & Request.Form.Item("phone3")
	sEmail = Request.Form.Item("email")
	sComments = Replace(Request.Form.Item("comments"), "'", "''")
	iAge = Request.Form.Item("age")
	
	'insert into particpant table
	sql = "INSERT INTO Participant (FirstName, LastName, Gender, Address, City, St, Zip, Phone, Email, Comments) VALUES ('" & sFirstName & "', '" 
	sql = sql & sLastName & "', '" & sGender & "', '" & sAddress & "', '" & sCity & "', '" & sState & "', '" & sZip & "', '" & sPhone & "', '" 
	sql = sql & sEmail & "', '" & sComments & "')"
	Set rs=conn.Execute(sql)
	Set rs=Nothing

	'get participant id
	sql = "SELECT ParticipantID FROM Participant WHERE FirstName = '" & Replace(sFirstName,"'", "''") 
	sql = sql & "' AND LastName = '" & Replace(sLastName,"'", "''") & "' AND DOB = '" & CDate(dDOB) & "' ORDER BY ParticipantID DESC"
	Set rs = conn.Execute(sql)
	lThisPart = rs(0).Value
	Set rs=Nothing
	
	'insert into part reg table
	sql = "INSERT INTO PartReg (ParticipantID, WhereReg, DateReg, RaceID)"
	sql = sql & " VALUES (" & lThisPart & ", 'Race Day', '" & Date & "', " & lThisRace & ")"
	Set rs=conn.Execute(sql)
	Set rs=Nothing
	
	'insert into part race table
	sql = "INSERT INTO PartRace (ParticipantID, Age, RaceID) VALUES (" & lThisPart & ", " & iAge & ", "  & lThisRace & ")"
	Set rs=conn.Execute(sql)
	Set rs=Nothing

	'get waiver data
	sql = "SELECT Waiver FROM Waiver WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sWaiver = Replace(rs(0).Value, "''", "'")
	Set rs = Nothing
	
	'assign bib
	
	bRegCompl = True

	sFirstName = Replace(sFirstName, "''", "'")
	sLastName = Replace(sLastName, "''", "'")
	sAddress = Replace(sAddress, "''", "'")
	sCity = Replace(sCity, "''", "'")
	sComments = Replace(sComments, "''", "'")
End If
%>
<!DOCTYPE html>
<html>
<head>

<title>Gopher State Events Mobile Registration</title>
<!--#include file = "../includes/meta2.asp" -->
<meta name="description" content="GSE Mobile Registration form.">
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
	<img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
	<h3 class="h3"><%=sEventName%> Mobile Race Day Registration</h3>
	
	<div class="row form-group bg-warning" style="color:#fff;">&nbsp;= Required Fields</div>

	<%If bRegCompl = True Then%>
		<div class="bg-success">Your registration is not complete until you SHOW THIS SCREEN TO THE RACE DAY 
		REGISTRATION STAFF so that they can process your entry fee and give you your bib number!</div>
	
		<br>

		<ul>
			<li>Participant: <%=sFirstName%> <%=sLastName%></li>
			<li>Event: <%=sEventName%></li> 
			<li>Race: <%=sRaceName%></li>
			<li>Gender: <%=sGender%></li>
			<li>Age: <%=iAge%></li>
			<li>Address: <%=sAddress%></li>
			<li>City: <%=sCity%></li>
			<li>State: <%=sState%></li>
			<li>Zip: <%=sZip%></li>
			<li>Phone: <%=sPhone%></li>
			<li>Email: <%=sEmail%></li>
		</ul>
		
		<p><%=sWaiver%></p>
	<%Else%>
		<form class="form-horizontal" name="race_day_reg" method="post" action="mobile_reg.asp?event_id=<%=lEventID%>">
		<%If UBound(Races, 2) = 1 Then%>
			<div class="row bg-warning" style="color:#fff;"><%=sRaceName%></div>
		<%Else%>
			<div class="row form-group bg-warning" style="color:#fff;">
				<label class="control-label col-sm-4" for="races">Select Race:</label>
				<div class="col-sm-8">
					<select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
						<option value="">&nbsp;</option>
						<%For i = 0 to UBound(Races, 2) - 1%>
							<%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
								<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
							<%Else%>
								<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
							<%End If%>
						<%Next%>
					</select>
				</div>
			</div>
		<%End If%>

		<div class="row form-group bg-warning" style="color:#fff;">
			<label class="control-label col-sm-4" for="first_name">First Name:</label>
			<div class="col-sm-8">
				<input type="text" class="form-control form-control-sm" name="first_name" id="first_name">
			</div>
		</div>
		<div class="row form-group bg-warning" style="color:#fff;">
			<label class="control-label col-sm-4" for="last_name">Last Name:</label>
			<div class="col-sm-8">
				<input type="text" class="form-control form-control-sm" name="last_name" id="last_name">
			</div>
		</div>
		<div class="row form-group bg-warning" style="color:#fff;">
			<label class="control-label col-sm-4" for="gender">Gender:</label>
			<div class="col-sm-8">
				<select class="form-control form-control-sm" name="gender" id="gender">
					<option value=""></option>
					<option value="M">Male</option>
					<option value="F">Female</option>
				</select>
			</div>
		</div>
		<div class="row form-group bg-warning" style="color:#fff;">
			<label class="control-label col-sm-4" for="age">Age:</label>
			<div class="col-sm-8">
				<input type="text" class="form-control form-control-sm" name="age" id="age">
			</div>
		</div>
		<div class="row form-group">
			<label class="control-label col-sm-4" for="email">Email:</label>
			<div class="col-sm-8">
				<input type="text" class="form-control form-control-sm" name="email" id="email">
			</div>
		</div>
		<div class="row form-group">
			<label class="control-label col-sm-4" for="mobile">Mobile Phone:</label>
			<div class="col-sm-8">
				<input type="text" class="form-control form-control-sm" name="mobile" id="mobile">
			</div>
		</div>
		<div class="row form-group">
			<label class="control-label col-sm-4" for="provider">Cell Provider:</label>
			<div class="col-sm-8">
				<select class="form-control form-control-sm" name="provider" id="provider">
					<option value=""></option>
					<option value="M">Male</option>
					<option value="F">Female</option>
				</select>
			</div>
		</div>
		<div class="row form-group">
			<label class="control-label col-sm-4" for="city">City:</label>
			<div class="col-sm-8">
				<input type="text" class="form-control form-control-sm" name="city" id="city">
			</div>
		</div>
		<div class="row form-group">
			<label class="control-label col-sm-4" for="st">St/Prov:</label>
			<div class="col-sm-8">
				<input type="text" class="form-control form-control-sm" name="st" id="st">
			</div>
		</div>

		<input class="form-control" type="hidden" name="submit_reg" id="submit_reg" value="submit_reg">
		<input class="form-control" type="submit" name="submit1" id="submit1" value="Submit Registration">
		</form>
	<%End If%>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>