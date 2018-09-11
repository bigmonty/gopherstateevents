<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sPhone, sEmail, sUserID, sPassword, sComments, sRole, sTech, sSupport

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_staff") = "submit_staff" Then
	sFirstName = Replace(Request.Form.Item("first_name"), "''", "'")
	sLastName = Replace(Request.Form.Item("last_name"), "''", "'")
	If Not Request.Form.Item("address") & "" = "" Then sAddress =  Replace(Request.Form.Item("address"), "''", "'")
	If Not Request.Form.Item("city") & "" = "" Then sCity =  Replace(Request.Form.Item("city"), "''", "'")
	If Not Request.Form.Item("state") & "" = "" Then sState =  Replace(Request.Form.Item("state"), "''", "'")
	If Not Request.Form.Item("zip") & "" = "" Then sZip =  Replace(Request.Form.Item("zip"), "''", "'")
	If Not Request.Form.Item("phone") & "" = "" Then sPhone =  Replace(Request.Form.Item("phone"), "''", "'")
	sEmail =  Replace(Request.Form.Item("email"), "''", "'")
	sUserID =  Replace(Request.Form.Item("user_id"), "''", "'")
	sPassword =  Replace(Request.Form.Item("password"), "''", "'")
	If Not Request.Form.Item("comments") & "" = "" Then sComments =  Replace(Request.Form.Item("comments"), "''", "'")

	sRole = Request.Form.Item("role")
    sTech = "n"
    sSupport = "n"

    Select Case sRole 
        Case "support"
            sSupport = "y"
        Case "tech"
            sTech = "y"
        Case Else
            sTech = "y"
            sSupport = "y"
    End Select

    sql = "INSERT INTO Staff (FirstName, LastName, Address, City, State, Zip, Phone, Email, UserID, Password, Comments, Tech, Support) VALUES ('" 
    sql = sql & sFirstName & "', '" & sLastName & "', '" & sAddress & "', '" & sCity & "', '" & sState & "', '" & sZip & "', '" & sPhone & "', '" & sEmail 
    sql = sql & "', '" & sUserID  & "', '" & sPassword & "', '" & sComments & "', '" & sTech & "', '" & sSupport & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    If Request.Form.Item("notify") = "y" Then
        Dim cdoMessage, cdoConfig
        Dim sMsg

        sMsg = "Dear " & sFirstName & " " & sLastName & ": " & vbCrLf & vbCrLf

        sMsg = sMsg & "You have been added as a staff member for Gopher State Events.  Welcome!  You will find your login credentials below.  Please "
        sMsg = sMsg & "log in at http://www.gopherstateevents.com/default.asp?sign_out=y at your earliest convenience.  There you will see the terms of this "
        sMsg = sMsg & "opportunity, be able to update your personal information, manage your login credentials, indicate a preference for events "
        sMsg = sMsg & "that you would like to work, view your work assignments, and more. " & vbCrLf & vbCrLf

		sMsg = sMsg & "Login Information:" & vbCrLf
		sMsg = sMsg & "User Name: " & sUserID & vbCrLf
		sMsg = sMsg & "Password: " & sPassword & vbCrLf & vbCrLf

        sMsg = sMsg & "Please contact bob.schneider@gopherstateevents.com with any questions. "& vbCrLf & vbCrLf

        sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
        sMsg = sMsg & "Bob Schneider " & vbCrLf
        sMsg = sMsg & "Owner " & vbCrLf
        sMsg = sMsg & "Gopher State Events, LLC " & vbCrLf
        sMsg = sMsg & "612.720.8427 " & vbCrLf
        sMsg = sMsg & "bob.schneider@gopherstateevents.com "

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	    Set cdoMessage = CreateObject("CDO.Message")
	    With cdoMessage
		    Set .Configuration = cdoConfig
            .To = sEmail
		    .BCC = "bob.schneider@gopherstateevents.com"
		    .From = "support@gopherstateevents.com;"
	        .Subject = "GSE Staff Member Welcome"
		    .TextBody = sMsg
		    .Send
	    End With
	    Set cdoMessage = Nothing
	    Set cdoConfig = Nothing
    End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Add Staff</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
	th{
		text-align:right;
	}
</style>

<script>
function chkFlds(){
 	if (document.add_staff.first_name.value == '' || 
 	    document.add_staff.last_name.value == '' ||
 	    document.add_staff.email.value == '' ||
	 	document.add_staff.user_id.value == '' || 
	 	document.add_staff.password.value == '')
		{
  		alert('First Name, Last Name, Email, User ID and Password are required.');
  		return false
  		}
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">Add Staff</h4>
			
		    <form name="add_staff" method="Post" action="add_staff.asp" onsubmit="return chkFlds();">
		    <table style="margin:10px;">
			    <tr>
				    <th>First Name:</th>
				    <td><input type="text" name="first_name" id="first_name"></td>
				    <th>Last Name:</th>
				    <td><input type="text" name="last_name" id="last_name"></td>
			    </tr>
			    <tr>
				    <th>Address:</th>
				    <td><input type="text" name="address" id="address"></td>
				    <th>City:</th>
				    <td><input type="text" name="city" id="city"></td>
			    </tr>
			    <tr>
				    <th>State:</th>
				    <td><input type="text" name="state" id="state" size="2" maxlength="2"></td>
				    <th>Zip:</th>
				    <td><input type="text" name="zip" id="zip" size="7"></td>
			    </tr>
			    <tr>
				    <th>Phone:</th>
				    <td><input type="text" name="phone" id="phone"></td>
				    <th>Email:</th>
				    <td><input type="text" name="email" id="email"></td>
			    </tr>
			    <tr>
				    <th>Role:</th>
				    <td>
                        <select name="role" id="role">
                            <option value="support">Support</option>
                            <option value="tech">Technician</option>
                            <option value="both">Both</option>
                        </select>
                    </td>
				    <th>Notify:</th>
				    <td>
                        <select name="notify" id="notify">
                            <option value="n">No</option>
                            <option value="y" selected>Yes</option>
                        </select>
                    </td>
			    </tr>
			    <tr>
				    <th>User Name:</th>
				    <td><input type="text" name="user_id" id="user_id" maxlength="12"></td>
				    <th>Password:</th>
				    <td><input type="text" name="password" id="password" maxlength="12"></td>
			    </tr>
			    <tr>
				    <th valign="top">Comments:</th>
				    <td colspan="3"><textarea name="comments" id="comments" cols="60" rows="3" style="font-size: 1.0em;"><%=sComments%></textarea></td>
			    </tr>
			    <tr>
				    <td style="background-color:#ececd8;text-align:center;" colspan="4">
					    <input type="hidden" name="submit_staff" id="submit_staff" value="submit_staff">
					    <input type="submit" name="submit1" id="submit1" value="Submit Staff Member">
				    </td>
			    </tr>
		    </table>
		    </form>
        </div>
	</div>
</div>
<%	
conn.Close
Set conn = Nothing
%>
</body>
</html>
