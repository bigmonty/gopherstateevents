<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sPhone, sEmail, sUserID, sPassword, sComments

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_meet_dir") = "submit_meet_dir" Then
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
	
    sql = "INSERT INTO MeetDir (FirstName, LastName, Address, City, State, Zip, Phone, Email, UserID, Password, Comments) VALUES ('" & sFirstName & "', '" 
	sql = sql & sLastName & "', '" & sAddress & "', '" & sCity & "', '" & sState & "', '" & sZip & "', '" & sPhone & "', '" & sEmail & "', '" & sUserID 
	sql = sql & "', '" & sPassword & "', '" & sComments & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE Add CCMeet Director</title>
<script>
function chkFlds(){
 	if (document.new_meet_dir.first_name.value == '' || 
 	    document.new_meet_dir.last_name.value == '' ||
 	    document.new_meet_dir.email.value == '' ||
	 	document.new_meet_dir.user_id.value == '' || 
	 	document.new_meet_dir.password.value == '')
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
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		
		<div class="col-md-10">
			<h4 class="h4">Add GSE Meet Director</h4>
			
			<form name="new_meet_dir" method="Post" action="add_meet_dir.asp" onsubmit="return chkFlds();">
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
					<td><input type="text" name="state" id="state" size="2"></td>
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
					<th>User Name:</th>
					<td><input type="text" name="user_id" id="user_id" maxlength="10"></td>
					<th>Password:</th>
					<td><input type="text" name="password" id="password" maxlength="10"></td>
				</tr>
				<tr>
					<th valign="top">Comments:</th>
					<td colspan="3"><textarea name="comments" id="comments" cols="60" rows="3"><%=sComments%></textarea></td>
				</tr>
				<tr>
					<td style="background-color:#ececd8;text-align:center;" colspan="4">
						<input type="hidden" name="submit_meet_dir" id="submit_meet_dir" value="submit_meet_dir">
						<input type="submit" name="submit1" id="submit1" value="Submit Event Director">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
<%	
conn.Close
Set conn = Nothing
%>
</body>
</html>
