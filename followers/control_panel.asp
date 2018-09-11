<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, conn, rs
Dim lTeamID, lRosterID, lPartID, lMeetID, lEventID
Dim i
Dim sFirstName, sLastName

If Session("follower_id") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

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
			<table>
				<tr>
					<td>
                    </td>
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
