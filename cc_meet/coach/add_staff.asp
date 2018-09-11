<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lProvider
Dim i, j
Dim sFirstName, sLastName, sMobilePhone, sEmail, sUserName, sPassword, sRole, sSendTo, sAllowAccess
Dim CellProviders,  StaffRoles(2)

If Not Session("role") = "coach" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

StaffRoles(0) = "Asst Coach"
StaffRoles(1) = "Manager"
StaffRoles(2) = "Other"

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_this") = "submit_this" Then
	sFirstName = Replace(Request.Form.Item("first_name"), "''", "'")
	sLastName = Replace(Request.Form.Item("last_name"), "''", "'")
	sMobilePhone =  Request.Form.Item("mobile_phone")
    lProvider =  Request.Form.Item("provider")
	sEmail =  Request.Form.Item("email")
	sSendTo =  Request.Form.Item("send_to")
	sAllowAccess =  Request.Form.Item("allow_access")
	sRole = Request.Form.Item("role")

    'generate user name
    sUserName = sFirstName & "_" & sLastName
    sUserName = Left(sUserName, 12)

    'generate unique password
    sPassword = CreatePassword()

    sql = "INSERT INTO TeamStaff (CoachesID, FirstName, LastName, MobilePhone, Provider, Email, SendTo, AllowAccess, Role, UserName, Password, WhenReg) "
    sql = sql & "VALUES (" & Session("my_id") & ", '" & sFirstName & "', '" & sLastName & "', '" & sMobilePhone & "', " & lProvider  & ", '" & sEmail  
    sql = sql & "', '" & sSendTo  & "', '" & sAllowAccess & "', '" & sRole  & "', '" & sUserName  & "', '" & sPassword & "', '" &  Now() & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    If sAllowAccess = "y" Then
        If Request.Form.Item("notify") = "y" Then
            If ValidEmail(sEmail) = True Then
                Dim cdoMessage, cdoConfig
                Dim sMsg

                sMsg = "Dear " & sFirstName & " " & sLastName & ": " & vbCrLf & vbCrLf

                sMsg = sMsg & "You have been added as a staff member for all teams coached by " & Session("my_name") & ".  Welcome!  You will find your login credentials "
                sMsg = sMsg & "below.  Please log in at http://www.gopherstateevents.com/default.asp?sign_out=y at your earliest convenience.  By doing so "
                sMsg = sMsg & "you will be able to assist your coach in managing lineups, editing the roster, etc.  Once on the site login using the "
                sMsg = sMsg & "credentials below and select 'CC/Nordic Coach' as your role. " & vbCrLf & vbCrLf

		        sMsg = sMsg & "Login Information:" & vbCrLf
		        sMsg = sMsg & "User Name: " & sUserName & vbCrLf
		        sMsg = sMsg & "Password: " & sPassword & vbCrLf & vbCrLf

                sMsg = sMsg & "Please note that this utility is designed purely for assisting with Cross-Country Running and Nordic Ski team management.  For our "
                sMsg = sMsg & "part your contact information will NEVER be used by us or provided to anyone else by us for purposes other than stated above.  "
                sMsg = sMsg & "Please use the site only as intended. " & vbCrLf & vbCrLf

                sMsg = sMsg & "Please contact bob.schneider@gopherstateevents.com with any questions. "& vbCrLf & vbCrLf

                sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
                sMsg = sMsg & "Bob Schneider " & vbCrLf
                sMsg = sMsg & "Owner " & vbCrLf
                sMsg = sMsg & "Gopher State Events, LLC " & vbCrLf
                sMsg = sMsg & "612.720.84277 " & vbCrLf
                sMsg = sMsg & "bob.schneider@gopherstateevents.com "

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	            Set cdoMessage = CreateObject("CDO.Message")
	            With cdoMessage
		            Set .Configuration = cdoConfig
                    .To = sEmail
		            .BCC = "bob.schneider@gopherstateevents.com;" & Session("my_email")
		            .From = "support@gopherstateevents.com;"
	                .Subject = "GSE Team Staff Welcome"
		            .TextBody = sMsg
		            .Send
	            End With
	            Set cdoMessage = Nothing
	            Set cdoConfig = Nothing
            End If
        End If
    End If
End If

Function CreatePassword()
    Dim sDefaultChars
    Dim x
    Dim sMyPassword
    Dim iPickedChar
    Dim iDefaultCharactersLength

    sDefaultChars="abcdefghijklmnpqrstuvxyzABCDEFGHIJKLMNPQRSTUVXYZ123456789_!"

    Randomize

    For x = 1 To 12
        iPickedChar = Int((59 * Rnd) + 1) 
        sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
    Next 

    CreatePassword = sMyPassword
End Function

%>
<!--#include file = "../../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Add Staff</title>

<script>
function chkFlds(){
 	if (document.add_staff.first_name.value == '' || 
 	    document.add_staff.last_name.value == '' ||
 	    document.add_staff.email.value == '' ||
	 	document.add_staff.role.value == '')
		{
  		alert('First Name, Last Name, Email, and Role are required.');
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
     
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
            <h4 class="h4">Add Team Staff</h4>
                    
            <p>
                Team Staff are folks that have almost daily responsibilities with your team.  This utility is designed to make our site a more functional 
                site for you the coach.  It's immediate purpose is so that members of your staff can receive results emails and assist you in managing 
                rosters and line-ups. Team Staff Members are "attached" to coaches, not teams, so they have a connection to all teams managed by that coach.  They can be given 
                access to team data or not as the coach desires.  NOTE:  Only the team's head coach will have access to any staff functionality.
            </p>

            <div class="row">
                <div class="col-sm-6">
                    <p>(First Name, Last Name, Email, and Role are required fields.)</p>

                    <form role="form" class="form-horizontal" name="add_staff" method="post" action="add_staff.asp" onsubmit="return chkFlds();">
                    <div class="form-group row">
                        <label for="first_name" class="control-label col-sm-3">First:</label>
                        <div class="col-sm-9">
                            <input type="text" class="form-control" name="first_name" id="first_name">
                        </div>
                    </div>
                    <div class="form-group row">
                        <label for="last_name" class="control-label col-sm-3">Last:</label>
                        <div class="col-sm-9">
                            <input type="text" class="form-control" name="last_name" id="last_name">
                        </div>
                    </div>
                    <div class="form-group row">
                        <label for="email" class="control-label col-sm-3">Email:</label>
                        <div class="col-sm-9">
                            <input type="text" class="form-control" name="email" id="email">
                        </div>
                     </div>
                    <div class="form-group row">
                       <label for="send_to" class="control-label col-sm-3">Send To:</label>
                        <div class="col-sm-9">
                            <select class="form-control" name="send_to" id="send_to">
                                <option value="n">No</option>
                                <option value="y"selected>Yes</option>
                            </select>
                        </div>
                    </div>
                    <div class="form-group row">
                        <label for="mobile_phone" class="control-label col-sm-3">Mobile #:</label>
                        <div class="col-sm-9">
                            <input type="text" class="form-control" name="mobile_phone" id="mobile_phone">
                        </div>
                    </div>
                    <div class="form-group row">
                        <label for="provider" class="control-label col-sm-3">Provider:</label>
                        <div class="col-sm-9">
                            <select class="form-control" name="provider" id="provider"> 
                                <option value="0">None</option>
                                <%For j = 0 To UBound(CellProviders, 2)%>
                                    <option value="<%=CellProviders(0, j)%>"><%=CellProviders(1, j)%></option>
                                <%Next%>
                            </select>
                        </div>
                    </div>
                    <div class="form-group row">
                        <label for="allow_access" class="control-label col-sm-3">Allow Access:</label>
                        <div class="col-sm-9">
                            <select class="form-control" name="allow_access" id="allow_access">
                                <option value="n" selected>No</option>
                                <option value="y">Yes</option>
                            </select>
                        </div>
                     </div>
                    <div class="form-group row">
                       <label for="role" class="control-label col-sm-3">Role:</label>
                        <div class="col-sm-9">
                            <select class="form-control" name="role" id="role">
                                <option value="">&nbsp;</option>
                                <%For j = 0 To UBound(StaffRoles)%>
                                    <option value="<%=StaffRoles(j)%>"><%=StaffRoles(j)%></option>
                                <%Next%>
                            </select>
                        </div>
                    </div>
                    <div class="form-group">
                        <input type="hidden" name="submit_this" id="submit_this" value="submit_this">
                        <input type="submit" class="form-control" name="submit1" id="submit1" value="Add Staff">
                    </div>
                    </form>
                </div>
                <div class="col-sm-6">
                    <h5 class="h5">Legend:</h5>
                    <dl class="dl">
                        <dt>Allow Access</dt>
                        <dd>
                            Selecting "Yes" here allows this person full access to your team data (roster, line-ups, etc).  This is helpful if you 
                            trust them to assist with roster and lineup entry on the site.  Please be thoughtful
                            who you give this access to.  The default value is "No".
                        </dd>
                        <dt>Send To</dt>
                        <dd>
                            Selecting "Yes" here allows this person to receive results emails and notification when pictures are online.  The default 
                            value is "Yes".  This can be helpful on meet day to assist in identifying results errors in case you do not have
                            time to check your email during the meet.
                        </dd>
                        <dt>Notify</dt>
                        <dd>
                            Selecting "Yes" here allows this person to be notified that you have created this account for them and sends them their 
                            login credentials.  All Team Staff login using the "CC/Nordic Coach" role on the site's home page, regardless of their
                            role on the team.  NOTE:  If you have not set "Allow Access" to "Yes" they will not be notified of their login information.
                            If you change their status to allow them access to your 
                        </dd>
                    </dl>
                </div>
            </div>
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
