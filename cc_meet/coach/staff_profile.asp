<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lProvider
Dim sFirstName, sLastName, sEmail, sMobilePhone, sUserName, sPassword, sEmailMsg, sMsg
Dim bChangeLogin, bLoginExists
Dim cdoMessage, cdoConfig
Dim CellProviders

If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_info") = "submit_info" Then
	sFirstName = Replace(Request.Form.Item("first_name"), "''", "'")
	sLastName =  Replace(Request.Form.Item("last_name"), "''", "'")
	sEmail = Request.Form.Item("email")
	sMobilePhone = Request.Form.Item("mobile_phone")
    lProvider = Request.Form.Item("provider")

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Email, MobilePhone, Provider FROM TeamStaff WHERE TeamStaffID = " & Session("my_id")
	rs.Open sql, conn, 1, 2
	rs(0).Value = sFirstName
	rs(1).Value = sLastName
	rs(2).Value = sEmail
	rs(3).Value = sMobilePhone
    rs(4).Value = lProvider
	rs.Update
	rs.Close
	Set rs = Nothing
	
	sUserName = Request.Form.Item("user_name")
	sPassword = Request.Form.Item("password")
	
	sql = "SELECT UserName, Password FROM TeamStaff WHERE TeamStaffID = " & Session("my_id") & " AND UserName = '" & sUserName & "' AND Password = '" 
	sql = sql & sPassword & "'"
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
        bChangeLogin = True
    Else
		bChangeLogin = False
    End If
	Set rs = Nothing
	
	If bChangeLogin = True Then 
		'check for uniqueness of this login
		sql = "SELECT UserName, Password FROM TeamStaff"
		Set rs = conn.Execute(sql)
		Do While Not rs.EOF
			If CStr(sUserName) = cStr(rs(0).Value) Then
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
			sql = "SELECT UserName, Password, Email FROM TeamStaff WHERE TeamStaffID = " & Session("my_id")
			rs.Open sql, conn, 1, 2
			rs(0).Value = sUserName
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
			sEmailMsg = sEmailMsg & "Your User Name is: " & sUserName & vbCrLf
			sEmailMsg = sEmailMsg & "Your Password is: " & sPassword & vbCrLf & vbCrLf
			
			sEmailMsg = sEmailMsg & "Sincerely~" & vbCrLf
			sEmailMsg = sEmailMsg & "Bob Schneider" & vbCrLf
			sEmailMsg = sEmailMsg & "612.720.8427" & vbCrLf
			sEmailMsg = sEmailMsg & "Gopher State Events/eTRaXC"
			
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
sql = "SELECT FirstName, LastName, Email, MobilePhone, UserName, Password, Provider FROM TeamStaff WHERE TeamStaffID = " & Session("my_id")
Set rs = conn.Execute(sql)
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName =  Replace(rs(1).Value, "''", "'")
sEmail = rs(2).Value
sMobilePhone = rs(3).Value
sUserName = rs(4).Value
sPassword = rs(5).Value
lProvider = rs(6).Value
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Cross-Country Team Staff Profile</title>

<!--#include file = "../../includes/js.asp" --> 
     
<script>
function chkFields(){
	if (document.staff_info.first_name.value==''){
		alert('You must supply a first name!');
		return false;
	}
	else
		if (document.staff_info.last_name.value==''){
			alert('You must supply a last name!');
			return false;
		}
	else
		if (document.staff_info.email.value==''){
			alert('You must supply an email address!');
			return false;
		}
	else
		return true;
}
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
    <!--#include file = "../../includes/coach_menu.asp" -->

	<h4 class="h4">GSE Team Staff Profile Page:  Welcome <%=sFirstName%> <%=sLastName%>!</h4>
			
	<p>NOTE: The information included on this site is considered private and for event management purposes only.  Under no circumstances will it 
	ever be made available for sale or gift to any third party for any reason without the express permission of the person that it represents.
	Please fill this form out on your first visit to the site and attempt to keep the information here current and complete so that we can contact 
	you as needed for meet-pertinent information.</p>
			
	<form role="form" class="form-horizontal" name="staff_info" method="post" action="staff_profile.asp" onsubmit="return chkFields()">
	<div  class="form-group">
		<label for="first_name" class="control-label col-xs-2">First Name:</label>
		<div class="col-xs-4">
            <input class="form-control" name="first_name" id="first_name" maxlength = "10" value="<%=sFirstName%>">
        </div>
		<label for="last_name" class="control-label col-xs-2">Last Name:</label>
		<div class="col-xs-4">
            <input class="form-control" name="last_name" id="last_name" maxlength = "15" value="<%=sLastName%>">
        </div>
    </div>
    <div class="form-group">
		<label for="mobile_phone" class="control-label col-xs-2">Mobile Phone:</label>
		<div class="col-xs-4">
            <input class="form-control" name="mobile_phone" id="mobile_phone" maxlength = "12" value="<%=sMobilePhone%>">
        </div>
		<label for="provider" class="control-label col-xs-2">Mobile Provider:</label>
		<div class="col-xs-4">
            <select class="form-control" name="provider" id="provider"> 
                <option value="0">None</option>
				<%For j = 0 To UBound(CellProviders, 2)%>
                    <%If CLng(lProvider) = CLng(CellProviders(0, j)) Then%>
						<option value="<%=CellProviders(0, j)%>" selected><%=CellProviders(1, j)%></option>
					<%Else%>
						<option value="<%=CellProviders(0, j)%>"><%=CellProviders(1, j)%></option>
					<%End If%>
                <%Next%>
			</select>
        </div>
    </div>
    <div class="form-group">
		<label for="email" class="control-label col-xs-2">Email:</label>
		<div class="col-xs-10">
            <input class="form-control" name="email" id="email" maxlength = "25" value="<%=sEmail%>">
        </div>
    </div>
    <div class="form-group">
		<label for="user_name" class="control-label col-xs-2">User ID:</label>
		<div class="col-xs-2">
            <input class="form-control" type="text" name="user_name" id="user_name" value="<%=sUserName%>" maxlength="10">
        </div>
		<label for="password" class="control-label col-xs-2">Password:</label>
		<div class="col-xs-2">
            <input class="form-control" type="password" name="password" id="password" value="<%=sPassword%>" maxlength="10">
        </div>
		<label for="confirm_pword" class="control-label col-xs-2">Confirm:</label>
		<div class="col-xs-2">
            <input class="form-control" type="password" name="confirm_pword" id="confirm_pword">
        </div>
    </div>
    <div class="form-group">
		<input class="form-control" type="hidden" name="submit_info" id="submit_info" value="submit_info">
		<input class="form-control" type="submit" name="submit" id="submit" value="Save Changes">
    </div>
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
