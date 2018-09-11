<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i
Dim lPartID
Dim sPartName, sMyGender, sErrMsg, sMsg, sUserName, sPassword
Dim cdoMessage, cdoConfig

Session.Contents.RemoveAll()

lPartID = Request.QueryString("part_id")
If CStr(lPartID) & "" = "" Then lPartID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"
	
Dim sRandPic
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, PixName FROM RacePix ORDER BY NEWID()"
rs.Open sql, conn, 1, 2
sRandPic = "/gallery/" & rs(0).Value & "/" & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_login") = "submit_login" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'my_hist_login' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_my_hist_login") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_my_hist_login") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg
		
		sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
		If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))
		
		If sHackMsg = vbNullString Then
			'check for event director login and redirect to their site if valid
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT MyHistID, ParticipantID FROM MyHist WHERE UserName = '" & sUserName & "' AND Password = '" & sPassword & "'"
			rs.Open sql, conn, 1, 2
			If rs.RecordCount > 0 Then 
                Session("my_hist_id") = rs(0).value
                Session("part_id") = rs(1).Value
            End If
			rs.Close
			Set rs = Nothing
					
			If CStr(Session("my_hist_id")) = vbNullString Then
				sErrMsg = "We are sorry but those login credentials were not found.  Please try again or <a href=mailto:bob.schneider@gopherstateevents.com>contact</a> H51Software, LLC for assistance."
			Else
			    'check for event director login and redirect to their site if valid
			    Set rs = Server.CreateObject("ADODB.Recordset")
			    sql = "SELECT PartID FROM PartData WHERE MyHistID = " & Session("my_hist_id")
			    rs.Open sql, conn2, 1, 2
                Session("etraxc_id") = rs(0).value
			    rs.Close
			    Set rs = Nothing

                sql = "INSERT INTO MyHistLogin (MyHistID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_hist_id") & ", '" & Now() 
                sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                Set rs=conn.Execute(sql)
                Set rs=Nothing

				Response.Redirect "profile.asp"
			End If
		End If
    End If
End If

'log this user if they are just entering the site
If Session("access_my_hist_login") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'my_hist_login')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'my_hist_login'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_my_hist_login")
End If

If Not CLng(lPartID) = 0 Then
	sql = "SELECT FirstName, LastName, Gender FROM Participant WHERE ParticipantID = " & lPartID
	Set rs = conn.Execute(sql)
	sPartName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
	sMyGender = rs(2).Value
	Set rs = Nothing
End If

%>
<!--#include file = "../includes/cdo_connect.asp" -->

<!--#include file = "../includes/clean_input.asp" -->
<%
Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>My GSE&copy; Login</title>
<meta name="description" content="My participant history login for a Gopher State Events (GSE) timed event.">
<!--#include file = "../includes/js.asp" --> 

<script>
function chkFlds() {
if (document.my_hist_login.user_name.value == '' || 
    document.my_hist_login.password.value == '') 
{
 	alert('All fields are required!');
 	return false
 	}
else
 	return true;
}
</script>
</head>
<body onload="document.my_hist_login.user_name.focus();">
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="My GSE History Portal">
    <h3 class="h3">My GSE History Login</h3>

    <div class="bg-info">
        With GSE's My History app you  can use the full force of the web to track your racing, manage your fitness and 
        training, plan future races, and otherwise live an active life full of energy and exhuberance.  This FREE service coordinates with
        <a href="http://www.my-etraxc.com/" style="font-weight: bold;">My-eTRaXC</a>, a free online training, lifestyle, and record keeping account.  
        THIS IS A COMPLETELY FREE SERVICE!  The only thing you MIGHT wish to spend your hard earned money on is a mobile app.
        We are currently working on the iPhone version of this service.
    </div>
    <br>
    <div class="bg-danger">
        Due to the major restructuring of GSE's My History utility, all original accounts have been deleted and will need to be re-created.  We chose this
        approach because the utility was not reasonably functional and it gave us the opportunity to start over and "do it right."  We thank you for your
        understanding.
    </div>

    <div class="col-sm-6">
	    <%If Not sHackMsg = vbNullString Then%>
		    <p class="text-danger"><%=sHackMsg%></p>
	    <%Else%>
		    <%If Not sErrMsg = vbNullString Then%>
			    <p class="text-danger"><%=sErrMsg%></p>
		    <%End If%>

            <h4 class="h4">Sign in</h4>

  			<form role="form" class="form"  name="my_hist_login" method="Post" action="my_hist_login.asp" onSubmit="return chkFlds();">
            <div class="form-group">
  			    <label for="user_name">User Name:</label>
			    <input type="text" class="form-control" name="user_name" id="user_name" size="12"  maxlength="12" value="<%=sUserName%>">
            </div>
            <div class="form-group">
			    <label for="password">Password:</label>
			    <input type="password" class="form-control" name="password" id="password" size="12"  maxlength="12" value="<%=sPassword%>">
            </div>
            <div class="form-group">
				<input type="hidden" name="submit_login" id="submit_login" value="submit_login">
				<input type="submit" class="form-control" name="submit1" id="submit1" value="Login">
            </div>
            </form>
			<div class="bg-success" style="text-align: center;">
				<a href="javascript:pop('forgot_signin.asp',600,550)">Forgot Sign In?</a>
                |
                <a href="create_accnt.asp?part_id=<%=lPartID%>">Create Account</a>
            </div>
	    <%End If%>
    </div>
	<div class="col-sm-6">
        <br>
		<a href="http://www.my-etraxc.com/" onclick="openThis2(this.href,1024,760);return false;">
		    <img src="/graphics/my-etraxc_ad.gif" alt="My-eTRaXC" class="img-responsive">
        </a>

		<a href="<%=sRandPic%>" onclick="openThis2(this.href,1024,768);return false;">
            <img src="<%=sRandPic%>" alt="<%=sRandPic%>" class="img-responsive">
        </a>
	</div>
</div>
<%
conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>
</body>
</html>
