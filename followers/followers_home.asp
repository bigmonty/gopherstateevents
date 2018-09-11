<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, conn, rs
Dim i
Dim sUserName, sPassword, sLogInErr, sFirstName, sLastName, sEmail, sPassword2, sCreateAccntErr
Dim sHackMsg, sMsgText, sMsg
Dim bFound
Dim cdoMessage, cdoConfig

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%
    
If Request.Form.Item("submit_login") = "submit_login" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'followers_home' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_followers_home") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_followers_home") = "y" Then	'if they are an authorized user allow them to proceed
		sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
		If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))

		If sHackMsg = vbNullString Then
            bFound = False
	        Set rs = Server.CreateObject("ADODB.Recordset")
	        sql = "SELECT FollowersID FROM Followers WHERE UserName = '" & sUserName & "' AND Password = '" & sPassword & "'"
	        rs.Open sql, conn, 1, 2
	        If rs.RecordCount > 0 Then
                Session("follower_id") = rs(0).Value
                bFound = True
            Else
                sLogInErr = "I'm sorry.  Those credentials were not found.  Please try again or create an account."
	        End If
	        rs.Close
	        Set rs = Nothing
 					
            If bFound = True Then 
	            sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'followers_home'"
	            Set rs = conn.Execute(sql)
	            Set rs = Nothing

	            Session.Contents.Remove("access_followers_home")
                Response.Redirect "control_panel.asp?follower_id=" & lFollowerID
            End If
        End If
    End If
ElseIf Request.Form.Item("submit_account") = "submit_account" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'followers_home' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_followers_home") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_followers_home") = "y" Then	'if they are an authorized user allow them to proceed
		sUserName = CleanInput(Trim(Request.Form.Item("new_user_name")))
		If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("new_password")))
        If sHackMsg = vbNullString Then sFirstName = CleanInput(Trim(Request.Form.Item("first_name")))
		If sHackMsg = vbNullString Then sLastName = CleanInput(Trim(Request.Form.Item("last_name")))
        If sHackMsg = vbNullString Then sEmail = CleanInput(Trim(Request.Form.Item("email")))
        If sHackMsg = vbNullString Then sPassword2 = CleanInput(Trim(Request.Form.Item("password2")))

		If sHackMsg = vbNullString Then
            If Not sPassword = sPassword2 Then
                sCreateAccntErr = "Your passwords do not match."
            ElseIf ValidEmail(sEmail) = False  Then
                sCreateAccntErr = "Oops.  Your email address does not look correctly formatted."
            ElseIf IsUnique("Email", sEmail) = "n"  Then
                sCreateAccntErr = "Oops.  That email address is already in the system.  All email addresses must be unique."
            ElseIf IsUnique("UserName", sUserName) = "n"  Then
                sCreateAccntErr = "Oops.  That user name is already in use."
            ElseIf IsUnique("Password", sPassword) = "n"  Then
                sCreateAccntErr = "Oops.  That password is already in use."
            Else
                sql = "INSERT INTO Followers (FirstName, LastName, Email, UserName, Password, WhenReg) VALUES ('" & Replace(sFirstName, "'", "''") & "', '"
                sql = sql & Replace(sLastName, "'", "''") & "', '" & Replace(sEmail, "'", "''") & "', '" & sUserName & "', '" & sPassword 
                sql = sql & "', '" & Now() & "')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing

                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT FollowersID FROM Followers WHERE FirstName = '" & Replace(sFirstName, "'", "''") & "' AND LastName = '" 
                sql = sql & Replace(sLastName, "'", "''") & "' AND Email = '" & Replace(sEmail, "'", "''") & "' AND UserName = '" & sUserName
                sql = sql & "' AND Password = '" & sPassword & "' ORDER BY FollowersID DESC"
                rs.Open sql, conn, 1, 2
                Session("follower_id") = rs(0).Value
                rs.Close
                Set rs = Nothing

                sMsg = "New GSE Follower " & vbCrLf
                sMsg = sMsg & "Name: " & sFirstName & " " & sLastName & " " & vbCrLf
                sMsg = sMsg & "Email: " & sEmail

                'send me email of new follower
	            Set cdoMessage = CreateObject("CDO.Message")
	            With cdoMessage
		            Set .Configuration = cdoConfig
		            .To = "bob.schneider@gopherstateevents.com"
		            .From = "bob.schneider@gopherstateevents.com"
		            .Subject = "New GSE Follower"
		            .TextBody = sMsg
		            .Send
	            End With
	            Set cdoMessage = Nothing
	            Set cdoConfig = Nothing

	            sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'followers_home'"
	            Set rs = conn.Execute(sql)
	            Set rs = Nothing

	            Session.Contents.Remove("access_followers_home")
                Response.Redirect "control_panel.asp
            End If
        End If
    End If
End If

Function IsUnique(sField, sValue)
    IsUnique = "y"

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FollowersID FROM Followers WHERE " & sField & " = '" & sValue & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then IsUnique = "n"
    rs.Close
    Set rs = Nothing
End Function

%>
<!--#include file = "../includes/valid_email.asp" -->
<%

'log this user if they are just entering the site
If Session("access_followers_home") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'followers_home')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

%>
<!--#include file = "../includes/clean_input.asp" -->
<%
Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/js.asp" -->
<title>GSE Followers Utility</title>
<!--#include file = "../includes/meta2.asp" -->

<script>
function chkFlds() {
if (document.create_account.first_name.value == '' || 
    document.create_account.last_name.value == '' ||
    document.create_account.email.value == '' ||
    document.create_account.user_name.value == '' || 
    document.create_account.password.value == '' ||
    document.create_account.password2.value == '') 
{
 	alert('All fields are required when creating an account!');
 	return false
 	}
else
 	return true;
}

function chkFlds2() {
if (document.sign_in.user_name.value == '' || 
    document.sign_in.password.value == '') 
{
 	alert('All fields are required when signing in!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="content" style="width:1000px;">
		<div style="width:1000px;margin:0;color:#036;padding: 0;font-size: 0.9em;">
 			<h1 style="margin:5px;padding:5px;font-size:1.1em;">Gopher State Events Followers Utility</h1>
			
            <p>
                The GSE Follower Utility is designed to allow non-participants (friends, family, coaches, managers, team members, etc) get updates
                via email and/or text message on a specific participant, team, or event/meet as soon as it is available.  This would typically involve 
                race results but could also include being notified when finish line pictures go online, etc.
            </p>

            <p>
                <span style="font-weight: bold;">GSE PRIVACY PROMISE:</span>  Under no circumstances will we use your GSE Follower information for any 
                purpose other than that which it was submitted for.  We will not use it to promote our own events and services and we certainly would 
                never make this data available to another third party.  Ever!  Promise!  Further, you can opt out of receiving notifications, delete your
                account, or change your notification settings at any time.
            </p>

            <div style="float: left;width: 600px;font-size: 0.9em;margin-top: 0;padding-top: 0;">
                <h4 style="background: none;border: none;">Create Follower Account</h4>

                <%If Not sCreateAccntErr = vbNullString Then%>
                    <p><%=sCreateAccntErr%></p>
                <%End If%>

			    <form name="create_account" method="Post" action="followers_home.asp" onsubmit="return chkFlds();">
			    <table style="font-size: 1.0em;">
				    <tr>
                        <th>First Name:</th>
					    <td><input type="text" name="first_name" id="first_name" maxlength="25" value="<%=sFirstName%>"></td>
                        <th>Last Name:</th>
					    <td><input type="text" name="last_name" id="last_name" maxlength="25" value="<%=sLastName%>"></td>
				    </tr>
				    <tr>
                        <th>Email:</th>
					    <td><input type="text" name="email" id="email" maxlength="50" size="30" value="<%=sEmail%>"></td>
                        <th>User Name:</th>
					    <td><input type="text" name="new_user_name" id="new_user_name" maxlength="12" value="<%=sUserName%>"></td>
				    </tr>
				    <tr>
                        <th>Password:</th>
					    <td><input type="password" name="new_password" id="new_password" maxlength="12"></td>
                        <th>Password Again:</th>
					    <td><input type="password" name="password2" id="password2" maxlength="12"></td>
				    </tr>
				    <tr>
					    <td style="text-align:center;" colspan="4">
						    <input type="hidden" name="submit_account" id="submit_account" value="submit_account">
						    <input type="submit" name="submit1" id="submit1" value="Create Account">
					    </td>
				    </tr>
			    </table>
			    </form>
            </div>
            <div style="margin-left: 625px;font-size: 0.9em;margin-top: 10px;padding: 5px;background-color: #ececd8;">
                <h4 style="background: none;border: none;">Sign In</h4>

                <%If Not sLoginErr = vbNullString Then%>
                    <p><%=sLogInErr%></p>
                <%End If%>

			    <form name="sign_in" method="Post" action="followers_home.asp" onsubmit="return chkFlds2();">
			    <table style="font-size: 1.0em;">
				    <tr>
                        <th>User Name:</th>
					    <td><input type="text" name="user_name" id="user_name" size="10" maxlength ="12"></td>
				    </tr>
				    <tr>
                        <th>Password:</th>
					    <td><input type="password" name="password" id="password" size="10" maxlength="12"></td>
				    </tr>
				    <tr>
					    <td style="text-align: center;" colspan="2">
						    <input type="hidden" name="submit_login" id="submit_login" value="submit_login">
						    <input type="submit" name="submit2" id="submit2" value="Sign In">
					    </td>
				    </tr>
			    </table>
			    </form>
            </div>
		</div>
	</div>
</div>
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
