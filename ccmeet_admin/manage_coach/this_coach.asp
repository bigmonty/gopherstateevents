<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lCoachID
Dim DataArr(10)
Dim sFirstName, sLastName, sEmail, sPhone, sUserID, sPassword, sComments

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lCoachID = Request.QueryString("coach_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	sFirstName = Replace(Request.Form.Item("first_name"), "'", "''")
	sLastName =  Replace(Request.Form.Item("last_name"), "'", "''")
	sEmail =  Request.Form.Item("email")
	sPhone =  Request.Form.Item("phone")
	sUserID =  Request.Form.Item("user_id")
	sPassword =  Request.Form.Item("password")
	If Not Request.Form.Item("comments") & "" = "" Then sComments =  Replace(Request.Form.Item("comments"), "'", "''")
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Email, Phone, UserID,  Password, Comments FROM Coaches WHERE CoachesID = " & lCoachID
	rs.Open sql, conn, 1, 2
	rs(0).Value = sFirstName
	rs(1).Value = sLastName
	rs(2).Value = sEmail
	rs(3).Value = sPhone
	rs(4).Value = sUserID
	rs(5).Value = sPassword
	rs(6).Value = sComments
	rs.Update
	rs.Close
	Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Email, Phone, UserID,  Password, Comments FROM Coaches WHERE CoachesID = " & lCoachID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
	sFirstName = Replace(rs(0).Value, "''", "'")
	sLastName = Replace(rs(1).Value, "''", "'")
	sEmail = rs(2).Value
	sPhone = rs(3).Value
	sUserID = rs(4).Value
	sPassword = rs(5).Value
	If Not rs(6).Value & "" = "" Then sComments = Replace(rs(6).Value, "''", "'")
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>This CC Coach</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
    <h3 class="h3">Cross-Country Coach Data</h3>
	<form role="form" class="form-horizontal" name="coach_data" method="post" action="this_coach.asp?coach_id=<%=lCoachID%>">
    <div class="form-group">
		<label for="first_name" class="control-label col-xs-2">First Name:</label>
		<div class="col-xs-4">
			<input type="text" class="form-control" name="first_name" id="first_name" maxlength="10" value="<%=sFirstName%>">
		</div>
		<label for="last_name" class="control-label col-xs-2">Last Name:</label>
		<div class="col-xs-4">
			<input type="text" class="form-control" name="last_name" id="last_name" maxlength="15"  value="<%=sLastName%>">
		</div>
	</div>
	<div class="form-group">
		<label for="phone" class="control-label col-xs-2">Phone:</label>
		<div class="col-xs-4">
			<input type="text" class="form-control" name="phone" id="phone" maxlength="12"  value="<%=sPhone%>">
		</div>
		<label for="email" class="control-label col-xs-2">Email:</label>
		<div class="col-xs-4">
			<input type="text" class="form-control" name="email" id="email" maxlength="50"  value="<%=sEmail%>">
		</div>
	</div>
	<div class="form-group">
		<label for="user_id" class="control-label col-xs-2">User ID:</label>
		<div class="col-xs-4">
			<input type="text" class="form-control" name="user_id" id="user_id" maxlength="10"  value="<%=sUserID%>">
		</div>
		<label for="password" class="control-label col-xs-2">Password:</label>
		<div class="col-xs-4">
			<input type="text" class="form-control" name="password" id="password" maxlength="10"  value="<%=sPassword%>">
		</div>
	</div>
	<div class="form-group">
		<label for="comments" class="control-label col-xs-2">Comments:</label>
		<div class="col-xs-10">
			<textarea class="form-control" name="comments" id="comments" rows="2"><%=sComments%></textarea>
		</div>
	</div>
	<div class="form-group">
		<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
		<input type="submit" class="form-control" name="submit" id="submit" value="Save Changes">
	</div>
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
