<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, conn, rs
Dim i
Dim sUserName, sPassword, sFirstName, sLastName, sEmail, sMobilePhone, sProvider

If Session("follower_id") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Email, MobilePhone, Provider, UserName, Password FROM Followers WHERE FollowersID = " & Session("follower_id")
rs.Open sql, conn, 1, 2
sFirstName = Replace(rs(0).Value, "''", "'")
sFirstName = Replace(rs(1).Value, "''", "'")
sEmail = rs(2).Value
sMobilePhone = rs(3).Value
sProvider = rs(4).Value
sUserName = rs(5).Value
sPassword = rs(6).Value
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE Followers Control Panel</title>
<!--#include file = "../includes/meta2.asp" -->



</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="content" style="width:1000px;">
		<div style="width:1000px;margin:0;color:#036;padding: 0;">
            <div style="margin: 0;padding: 0;text-align: right;font-size: 0.8em;background-color: #ececec;">
                <a href="profile.asp?follower_id=<%=lFollowerID%>">My Profile</a>
                &nbsp;|&nbsp;
                <a href="control_panel.asp?follower_id=<%=lFollowerID%>">Control Panel</a>
            </div>
 			<h1 style="margin:5px;padding:5px;font-size:1.1em;">Gopher State Events Followers Control Panel</h1>
			
			<form name="my_settings" method="Post" action="control_panel.asp">
			<table style="font-size: 1.0em;">
				<tr>
                    <th>First Name:</th>
					<td><input type="text" name="first_name" id="first_name" maxlength="25" value="<%=sFirstName%>"></td>
                    <th>Last Name:</th>
					<td><input type="text" name="last_name" id="last_name" maxlength="25" value="<%=sLastName%>"></td>
				</tr>
				<tr>
                    <th>Mobile Phone:</th>
					<td><input type="text" name="mobile_phone" id="mobile_phone" maxlength="12" size="12" value="<%=sMobilePhone%>"></td>
                    <th>Provider:</th>
					<td>
                        <select name="provider" id="provider">
                            <option value="">&nbsp;</option>
                        </select>
                    </td>
				</tr>
				<tr>
                    <th>Email:</th>
					<td><input type="text" name="email" id="email" maxlength="50" size="30" value="<%=sEmail%>"></td>
                    <th>User Name:</th>
					<td><input type="text" name="user_name" id="user_name" maxlength="12" value="<%=sUserName%>"></td>
				</tr>
				<tr>
                    <th>Password:</th>
					<td><input type="password" name="new_password" id="new_password" maxlength="12"></td>
                    <th>Password Again:</th>
					<td><input type="password" name="password2" id="password2" maxlength="12"></td>
				</tr>
				<tr>
					<td style="text-align:center;">
						<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
						<input type="submit" name="submit" id="submit" value="Save Settings">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
