<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lMeetDirID
Dim DataArr(10)
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sEmail, sPhone, sUserID, sPassword, sComments

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lMeetDirID = Request.QueryString("meet_dir_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	sFirstName = Replace(Request.Form.Item("first_name"), "'", "''")
	sLastName =  Replace(Request.Form.Item("last_name"), "'", "''")
	sAddress = Replace(Request.Form.Item("address") , "'", "''")
	sCity = Replace(Request.Form.Item("city"), "'", "''")
	sState = Request.Form.Item("state") 
	sZip = Request.Form.Item("zip")
	sEmail =  Request.Form.Item("email")
	sPhone =  Request.Form.Item("phone")
	sUserID =  Request.Form.Item("user_id")
	sPassword =  Request.Form.Item("password")
	sComments =  Replace(Request.Form.Item("comments"), "'", "''")
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, UserID,  Password, Comments "
	sql = sql & "FROM MeetDir WHERE MeetDirID = " & lMeetDirID
	rs.Open sql, conn, 1, 2
	rs(0).Value = sFirstName
	rs(1).Value = sLastName
	rs(2).Value = sAddress
	rs(3).Value = sCity
	rs(4).Value = sState
	rs(5).Value = sZip
	rs(6).Value = sEmail
	rs(7).Value = sPhone
	rs(8).Value = sUserID
	rs(9).Value = sPassword
	rs(10).Value = sComments
	rs.Update
	rs.Close
	Set rs = Nothing
End If

sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, UserID,  Password, Comments "
sql = sql & "FROM MeetDir WHERE MeetDirID = " & lMeetDirID

Set rs = conn.Execute(sql)
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName = Replace(rs(1).Value, "''", "'")
sAddress = Replace(rs(2).Value, "''", "'")
sCity = Replace(rs(3).Value, "''", "'")
sState = rs(4).Value
sZip = rs(5).Value
sEmail = rs(6).Value
sPhone = rs(7).Value
sUserID = rs(8).Value
sPassword = rs(9).Value
sComments = Replace(rs(10).Value, "''", "'")
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CC Meet Director Data</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<h4 class="h4">This Meet Director Data</h4>
	<form name="meet_dir_data" method="post" action="this_meet_dir.asp?meet_dir_id=<%=lMeetDirID%>">
	<table>
		<tr>
			<td style="text-align:right">
				First Name:
			</td>
			<td style="text-align:left">
				<input name="first_name" id="first_name" size="8" maxlength="10" value="<%=sFirstName%>">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				Last Name:
			</td>
			<td style="text-align:left">
				<input name="last_name" id="last_name" size="10" maxlength="15"  value="<%=sLastName%>">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				Address:
			</td>
			<td style="text-align:left">
				<input name="address" id="address" size="15" maxlength="50"  value="<%=sAddress%>">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				City:
			</td>
			<td style="text-align:left">
				<input name="city" id="city" size="8" maxlength="10"  value="<%=sCity%>">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				St:
			</td>
			<td style="text-align:left">
				<input name="state" id="state" size="2" maxlength="2"  value="<%=sState%>" onkeyup="chkStr(this)">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				Zip:
			</td>
			<td style="text-align:left">
				<input name="zip" id="zip" size="7" maxlength="7"  value="<%=sZip%>" onkeyup="chkStr(this)">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				Phone:
			</td>
			<td style="text-align:left">
				<input name="phone" id="phone" size="10" maxlength="12"  value="<%=sPhone%>" onkeyup="chkStr(this)">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				Email:
			</td>
			<td style="text-align:left">
				<input name="email" id="email" size="18" maxlength="50"  value="<%=sEmail%>" onkeyup="chkStr(this)">
			</td>
		</tr>
		<tr>
			<td style="text-align:right;white-space:nowrap;">
				User ID:
			</td>
			<td style="text-align:left">
				<input name="user_id" id="user_id" size="8" maxlength="10"  value="<%=sUserID%>" onkeyup="chkStr(this)">
			</td>
		</tr>
		<tr>
			<td style="text-align:right">
				Password:
			</td>
			<td style="text-align:left">
				<input name="password" id="password" size="8" maxlength="10"  value="<%=sPassword%>" onkeyup="chkStr(this)">
			</td>
		</tr>
		<tr>
			<td style="text-align:right" valign="top">
				Comments:
			</td>
			<td style="text-align:left">
				<textarea name="comments" id="comments"  rows="2" cols="25"><%=sComments%></textarea>
			</td>
		</tr>
		<tr>
			<td style="text-align:center" colspan="2">
				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input type="submit" name="submit" id="submit" value="Save Changes">
			</td>
		</tr>
	</table>
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
