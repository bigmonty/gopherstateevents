<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sEmail, sPhone, sComments
Dim Staff(10)

If Not Session("role") = "staff" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_info") = "submit_info" Then
	If Not Request.Form.Item("first_name") & "" = "" Then sFirstName = Replace(Request.Form.Item("first_name"), "''", "'")
	If Not Request.Form.Item("last_name") & "" = "" Then sLastName =  Replace(Request.Form.Item("last_name"), "''", "'")
	If Not Request.Form.Item("address") & "" = "" Then sAddress =  Replace(Request.Form.Item("address"), "''", "'")
	If Not Request.Form.Item("city") & "" = "" Then sCity =  Replace(Request.Form.Item("city"), "''", "'")
	If Not Request.Form.Item("state") & "" = "" Then sState = Replace(Request.Form.Item("state"), "''", "'")
	If Not Request.Form.Item("zip") & "" = "" Then sZip = Replace(Request.Form.Item("zip"), "''", "'")
	If Not Request.Form.Item("email") & "" = "" Then sEmail = Replace(Request.Form.Item("email"), "''", "'")
	If Not Request.Form.Item("phone") & "" = "" Then sPhone = Replace(Request.Form.Item("phone"), "''", "'")
	If Not Request.Form.Item("comments") & "" = "" Then sComments =  Replace(Request.Form.Item("comments"), "''", "'")
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, Comments FROM Staff WHERE StaffID = " & Session("staff_id")
	rs.Open sql, conn, 1, 2
	
	If sFirstName & "" = "" Then
		s(0).Value = rs(0).OriginalValue
	Else
		rs(0).Value = sFirstName
	End if
	
	If sLastName & "" = "" Then
		s(1).Value = rs(1).OriginalValue
	Else
		rs(1).Value = sLastName
	End if
	
	rs(2).Value = sAddress
	rs(3).Value = sCity
	rs(4).Value = sState
	rs(5).Value = sZip
	
	If sEmail & "" = "" Then
		s(6).Value = rs(6).OriginalValue
	Else
		rs(6).Value = sEmail
	End if
	
	rs(7).Value = sPhone
	rs(8).Value = sComments
	rs.Update
	rs.Close
	Set rs = Nothing
End If

'get information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Phone, Address, City, State, Zip, Email, Comments, Tech, Support FROM Staff WHERE StaffID = " & Session("staff_id")
rs.Open sql, conn, 1, 2
For i = 0 to 10
	If not rs(i).Value & "" = "" Then Staff(i) =  Replace(rs(i).Value, "''", "'")
Next
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Staff Profile</title>
<meta name="description" content="Gopher State Events staff profile page.">
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
  	<div class="row">
		<!--#include file = "staff_menu.asp" -->
		<div class="col-md-10">
			<h3 class="admin_hdr">GSE Staff Profile</h3>
			
			<p>This page allows you to edit your personal information as it pertains to your role as a Gopher State Events staff member.  Please keep
            it as current as possible.</p>

			<form name="staff_info" method="post" action="profile.asp" onsubmit="return chkFields()">
			<table style="margin-top:10px;">
				<tr>
					<th>First Name:</th>
					<td><input name="first_name" id="first_name" maxlength = "10" size="10" value="<%=Staff(0)%>"></td>
					<th>Last Name:</th>
					<td><input name="last_name" id="last_name" maxlength = "15" size="15" value="<%=Staff(1)%>"></td>
					<th>Phone:</th>
					<td><input name="phone" id="phone" maxlength = "20" size="15" value="<%=Staff(2)%>"></td>
				</tr>
				<tr>
					<th>Address:</th>
					<td><input name="address" id="address" maxlength = "50" size="25" value="<%=Staff(3)%>"></td>
					<th>City:</th>
					<td><input name="city" id="city" maxlength = "25" size="15" value="<%=Staff(4)%>"></td>
					<th>State:</th>
					<td>
						<input name="state" id="state" maxlength = "2"  size="2" value="<%=Staff(5)%>">&nbsp;&nbsp;&nbsp;
						<span style="font-weight:bold;">Zip:</span>
						<input name="zip" id="zip" maxlength = "8"  size="8" value="<%=Staff(6)%>">
					</td>
				</tr>
				<tr>
					<th>Email:</th>
					<td><input name="email" id="email" maxlength = "25" size="30" value="<%=Staff(7)%>"></td>
					<th>Tech:</th>
					<td>&nbsp;<%=UCASE(Staff(9))%></td>
					<th>Support:</th>
					<td>&nbsp;<%=UCASE(Staff(10))%></td>
				</tr>
				<tr>
					<th valign="top">Comments:</th>
					<td colspan="5"><textarea name="comments" id="comments" rows="3" cols="60" style="font-size: 1.1em;"><%=Staff(8)%></textarea></td>
				</tr>
				<tr>
					<td colspan="6">
						<input type="hidden" name="submit_info" id="submit_info" value="submit_info">
						<input type="submit" name="submit" id="submit" value="Save Changes">
					</td>
				</tr>
			</table>
			</form>

            <p style="text-align: left;">IMPORTANT NOTE:  Your involvement with GSE is as an INDEPENDENT CONTACTOR for tax purposes.  You 
            will be paid by check on an event-by-event basis and no withholdings will be deducted.  It is your responsibility to claim your earnings
            on your taxes.</p>

           <h3>Google Nuggets</h3>
           <ul style="padding: 5px 5px 5px 20px;">
                <li>Doing things better or cheaper is not enough.  You have to do things differently.</li>
                <li>Do one thing really, really well.</li>
                <li>We are working much, much harder than we would in a normal job.</li>
                <li>You don't want to be looking at your competitors.  You want to shoot higher.  You want to be looking at what's possible.</li>
                <li>We see being great at something as a starting point, not an endpoint.</li>
                <li>Incremental improvement is guaranteed to be obsolete over time.</li>
                <li>If we can't win on quality, then we shouldn't win at all.</li>
                <li>You can make money without doing evil.</li>
                <li>Our goal is to develop services that improve the lives of as many people as possible.</li>
                <li>Focus on the user and all else will follow.</li>
                <li>The user is never wrong.</li>
           </ul>
		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" --> 
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>