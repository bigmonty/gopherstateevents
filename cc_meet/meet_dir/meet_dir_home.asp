<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sName, sSiteAccessStatus
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sEmail, sPhone, sComments, sUserID, sPassword, sMsg, sEmailMsg
Dim bChangeLogin, bLoginExists
Dim cdoMessage, cdoConfig

If Not Session("role") = "meet_dir" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT Status FROM SiteAccess ORDER BY SiteAccessID DESC"
Set rs = conn.Execute(sql)
sSiteAccessStatus = rs(0).Value
Set rs = Nothing

If Request.Form.Item("submit_info") = "submit_info" Then
	sFirstName = Replace(Request.Form.Item("first_name"), "''", "'")
	sLastName =  Replace(Request.Form.Item("last_name"), "''", "'")
	sAddress =  Replace(Request.Form.Item("address"), "''", "'")
	sCity =  Replace(Request.Form.Item("city"), "''", "'")
	sState = Request.Form.Item("state")
	sZip = Request.Form.Item("zip")
	sEmail = Request.Form.Item("email")
	sPhone = Request.Form.Item("phone")
	sComments =  Replace(Request.Form.Item("comments"), "''", "'")
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, Comments FROM MeetDir WHERE MeetDirID = " & Session("my_id")
	rs.Open sql, conn, 1, 2
	rs(0).Value = sFirstName
	rs(1).Value = sLastName
	rs(2).Value = sAddress
	rs(3).Value = sCity
	rs(4).Value = sState
	rs(5).Value = sZip
	rs(6).Value = sEmail
	rs(7).Value = sPhone
	rs(8).Value = sComments
	rs.Update
	rs.Close
	Set rs = Nothing
	
	sUserID = Request.Form.Item("user_id")
	sPassword = Request.Form.Item("password")
	bChangeLogin = True
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT UserID, Password FROM MeetDir WHERE MeetDirID = " & Session("my_id") & " AND UserID = '" & sUserID & "' AND Password = '" & sPassword & "'"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then bChangeLogin = False
	rs.Close
	Set rs = Nothing
	
	If bChangeLogin = True Then 
		'check for uniqueness of this login
		sql = "SELECT UserID, Password FROM MeetDir"
		Set rs = conn.Execute(sql)
		Do While Not rs.EOF
			If CStr(sUserID) = cStr(rs(0).Value) Then
				If CStr(sPassword) = CStr(rs(1).Value) Then
					bLoginExists = True
					Exit Do
				End If
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing

		If bLoginExists = True Then
			sMsg = "This user id and password combination is not available.  Please change both items and resubmit."
		Else
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT UserID, Password, Email FROM MeetDir WHERE MeetDirID = " & Session("my_id")
			rs.Open sql, conn, 1, 2
			rs(0).Value = sUserID
			rs(1).Value = sPassword
			sEmail = rs(2).Value
			rs.Update
			rs.Close
			Set rs = Nothing
			
			sEmailMsg = vbCrLf
			sEmailMsg = sEmailMsg & "You are receiving this email because a request for a change in login information for your "
			sEmailMsg = sEmailMsg & "Gopher State Events (www.gopherstateevents.com) account was requested.  If you did not make this "
			sEmailMsg = sEmailMsg & "request, please notify us immediately at 612.720.8427 or by sending an email to "
			sEmailMsg = sEmailMsg & "bob.schneider@gopherstateevents.com." & vbCrLf & vbCrLf
			
			sEmailMsg = sEmailMsg & "Here is your new login information: " & vbCrLf
			sEmailMsg = sEmailMsg & "Your UserID is: " & sUserID & vbCrLf
			sEmailMsg = sEmailMsg & "Your Password is: " & sPassword & vbCrLf & vbCrLf
			
			sEmailMsg = sEmailMsg & "Sincerely~" & vbCrLf
			sEmailMsg = sEmailMsg & "Bob Schneider" & vbCrLf
			sEmailMsg = sEmailMsg & "612.720.8427" & vbCrLf
			sEmailMsg = sEmailMsg & "Hangar51 Software/GSE/eTRaXC"
			
%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = sEmail
				.CC = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
				.Subject = "GSE New Login"
				.TextBody = sEmailMsg
				.Send
			End With
			Set cdoMessage = Nothing
			Set cdoConfig = Nothing
			
			sMsg = "Your login information has been changed.  An email has been sent to you confirming this change."
		End If
	End If
End If

'get information
sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, Comments, UserID, Password FROM MeetDir WHERE MeetDirID = " & Session("my_id")
Set rs = conn.Execute(sql)
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName =  Replace(rs(1).Value, "''", "'")
sAddress =  Replace(rs(2).Value, "''", "'")
sCity =  Replace(rs(3).Value, "''", "'")
sState = rs(4).Value
sZip = rs(5).Value
sEmail = rs(6).Value
sPhone = rs(7).Value
sComments =  Replace(rs(8).Value, "''", "'")
sUserID = rs(9).Value
sPassword = rs(10).Value
Set rs = Nothing

sql = "SELECT LastName FROM MeetDir WHERE MeetDirID = " & Session("my_id")
Set rs = conn.Execute(sql)
sName = rs(0).Value
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CCMeet Meet Director Home</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	<!--#include file = "../../includes/meet_dir_menu.asp" -->
		<h4 class="h4">CCMeet Director Home Page: Welcome, Meet Director <%=sName%></h4>
			
		<p>NOTE: The information included on this site is considered private and for event management purposes only.  Under no circumstances will it 
		ever be made available for sale or gift to any third party for any reason without the express permission of the person that it represents.
		Please fill this form out on your first visit to the site and attempt to keep the information here current and complete so that we can contact 
		you as needed for meet-pertinent information.</p>
			
		<div class="bg-info">
			<a href="javascript:pop('site_use.htm',600,700)">How to Use This Site</a>
		</div>
			
		<form class="form" name="meet_dir_info" method="post" action="meet_dir_home.asp">
		<h4 class="h4">My Profile</h4>
		<table class="table">
			<tr>
				<th>First Name:</th>
				<td><input type="text" class="form-control" name="first_name" id="first_name" maxlength = "10" value="<%=sFirstName%>"></td>
				<th>Last Name:</th>
				<td><input type="text" class="form-control" name="last_name" id="last_name" maxlength = "15" value="<%=sLastName%>"></td>
				<th>Phone:</th>
				<td><input type="text" class="form-control" name="phone" id="phone" maxlength = "20" value="<%=sPhone%>"></td>
			</tr>
			<tr>
				<th>Address:</th>
				<td><input type="text" class="form-control" name="address" id="address" maxlength = "50" value="<%=sAddress%>"></td>
				<th>City:</th>
				<td><input type="text" class="form-control" name="city" id="city" maxlength = "25" value="<%=sCity%>"></td>
				<th>State:</th>
				<td>
					<input type="text" class="form-control" name="state" id="state" maxlength = "2" value="<%=sState%>">&nbsp;&nbsp;
					Zip:&nbsp;
					<input type="text" class="form-control" name="zip" id="zip" maxlength = "8" value="<%=sZip%>">
				</td>
			</tr>
			<tr>
				<th>User ID:</th>
				<td><input type="text" class="form-control" name="user_id" id="user_id" value="<%=sUserID%>" maxlength="10"></td>
				<th>Password:</th>
				<td><input type="password" class="form-control" name="password" id="password" value="<%=sPassword%>" maxlength="10"></td>
				<th>Confirm Password:</th>
				<td><input type="password" class="form-control" name="confirm_pword" id="confirm_pword"></td>
			</tr>
			<tr>
				<th valign="top">Email:</th>
				<td valign="top"><input type="text" class="form-control" name="email" id="email" maxlength = "25" value="<%=sEmail%>"></td>
				<th valign="top">Comments:</th>
				<td colspan="3">
					<textarea class="form-control" name="comments" id="comments" rows="2"><%=sComments%></textarea>
				</td>
			</tr>
			<tr>
				<td style="text-align:center" colspan="6">
					<input type="hidden" name="submit_info" id="submit_info" value="submit_info">
					<input type="submit" class="form-control" name="submit" id="submit" value="Save Changes">
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
<%
conn.close
Set conn = Nothing
%>
</body>
</html>
