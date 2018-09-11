<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lCellProvidersID
Dim i, j
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sEmail, sPhone, sComments, sUserID, sPassword, sMsg, sEmailMsg, sCellPhone
Dim bChangeLogin, bLoginExists
Dim cdoMessage, cdoConfig
Dim CellProviders

If Not Session("role") = "coach" Then Response.Redirect "/default.asp?sign_out=y"

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
	sCellPhone = Request.Form.Item("cell_phone")
    lCellProvidersID =  Request.Form.Item("cell_provider")

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Email, CellPhone, CellProvidersID FROM Coaches WHERE CoachesID = " & Session("my_id")
	rs.Open sql, conn, 1, 2
	rs(0).Value = sFirstName
	rs(1).Value = sLastName
	rs(2).Value = sEmail
	rs(3).Value = sCellPhone
	rs(4).Value = lCellProvidersID
	rs.Update
	rs.Close
	Set rs = Nothing
	
	sUserID = Request.Form.Item("user_id")
	sPassword = Request.Form.Item("password")
	bChangeLogin = True
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT UserID, Password FROM Coaches WHERE CoachesID = " & Session("my_id") & " AND UserID = '" & sUserID & "' AND Password = '" 
	sql = sql & sPassword & "'"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then bChangeLogin = False
	rs.Close
	Set rs = Nothing
	
	If bChangeLogin = True Then 
		'check for uniqueness of this login
		sql = "SELECT UserID, Password FROM Coaches"
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
			sql = "SELECT UserID, Password, Email FROM Coaches WHERE CoachesID = " & Session("my_id")
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
sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, Comments, UserID, Password, CellPhone, CellProvidersID FROM Coaches "
sql = sql & "WHERE CoachesID = " & Session("my_id")
Set rs = conn.Execute(sql)
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName =  Replace(rs(1).Value, "''", "'")
If Not rs(2).Value & "" = "" Then sAddress =  Replace(rs(2).Value, "''", "'")
If Not rs(3).Value & "" = "" Then sCity =  Replace(rs(3).Value, "''", "'")
sState = rs(4).Value
sZip = rs(5).Value
sEmail = rs(6).Value
sPhone = rs(7).Value
If Not rs(8).Value & "" = "" Then sComments =  Replace(rs(8).Value, "''", "'")
sUserID = rs(9).Value
sPassword = rs(10).Value
sCellPhone = rs(11).Value
lCellProvidersID = rs(12).Value
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Coach Home</title>

<script>
function chkFields(){
	if (document.coach_info.first_name.value==''){
		alert('You must supply a first name!');
		return false;
	}
	else
		if (document.coach_info.last_name.value==''){
			alert('You must supply a last name!');
			return false;
		}
	else
		if (document.coach_info.email.value==''){
			alert('You must supply an email address!');
			return false;
		}
	else
		return true;
}
</script>

<style type="text/css">
    label {
        text-align: right;
    }
</style>
</head>
<body>
<div class="container">
    <!--#include file = "../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<h4 class="h4">GSE Coach Home Page:  Welcome Coach <%=sLastName%>!</h4>
					
			<p>NOTE: The information included on this site is considered private and for event management purposes only.  Under no circumstances will it 
			ever be made available for sale or gift to any third party for any reason without the express permission of the person that it represents.
			Please fill this form out on your first visit to the site and attempt to keep the information here current and complete so that we can contact 
			you as needed for meet-pertinent information.</p>
					
			<form role="form" class="form-horizontal" name="coach_info" method="post" action="coach_home.asp" onsubmit="return chkFields()">
			<div  class="form-group row">
				<label for="first_name" class="control-label col-sm-2">First Name:</label>
				<div class="col-sm-4">
					<input class="form-control" name="first_name" id="first_name" maxlength = "10" value="<%=sFirstName%>">
				</div>
				<label for="last_name" class="control-label col-sm-2">Last Name:</label>
				<div class="col-sm-4">
					<input class="form-control" name="last_name" id="last_name" maxlength = "15" value="<%=sLastName%>">
				</div>
			</div>
			<div class="form-group row">
				<label for="cell_phone" class="control-label col-sm-2">Cell Phone:</label>
				<div class="col-sm-4">
					<input class="form-control" name="cell_phone" id="cell_phone" maxlength = "12" value="<%=sCellPhone%>">
				</div>
				<label for="cell_provider" class="control-label col-sm-2">Cell Provider:</label>
				<div class="col-sm-4">
					<select class="form-control" name="cell_provider" id="cell_provider"> 
						<option value="0">None</option>
						<%For j = 0 To UBound(CellProviders, 2)%>
							<%If CLng(lCellProvidersID) = CLng(CellProviders(0, j)) Then%>
								<option value="<%=CellProviders(0, j)%>" selected><%=CellProviders(1, j)%></option>
							<%Else%>
								<option value="<%=CellProviders(0, j)%>"><%=CellProviders(1, j)%></option>
							<%End If%>
						<%Next%>
					</select>
				</div>
			</div>
			<div class="form-group row">
				<label for="email" class="control-label col-sm-2">Email:</label>
				<div class="col-sm-10">
					<input class="form-control" name="email" id="email" maxlength = "25" value="<%=sEmail%>">
				</div>
			</div>
			<div class="form-group row">
				<label for="user_id" class="control-label col-sm-2">User ID:</label>
				<div class="col-sm-2">
					<input class="form-control" type="text" name="user_id" id="user_id" value="<%=sUserID%>" maxlength="10">
				</div>
				<label for="password" class="control-label col-sm-2">Password:</label>
				<div class="col-sm-2">
					<input class="form-control" type="password" name="password" id="password" value="<%=sPassword%>" maxlength="10">
				</div>
				<label for="confirm_pword" class="control-label col-sm-2">Confirm:</label>
				<div class="col-sm-2">
					<input class="form-control" type="password" name="confirm_pword" id="confirm_pword">
				</div>
			</div>
			<div class="form-group">
				<input class="form-control" type="hidden" name="submit_info" id="submit_info" value="submit_info">
				<input class="form-control" type="submit" name="submit" id="submit" value="Save Changes">
			</div>
			</form>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
