<%@ Language=VBScript %>
<%
Option Explicit

Dim rs, sql, conn, conn2, rs2, sql2
Dim i
Dim sErrMsg, sUserName, sPassword, sRole, sSignOut, sLoginErr
Dim dTimeNow
Dim bNotFound

dTimeNow = Time()

Session.Contents.Remove("role")

sUserName = vbNullString
sPassword = vbNullString

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
				
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_login") = "submit_login" Then
	'see if this user has entered from the form correctly within the past 20 minutes
    Set rs = Server.CreateObject("ADODB.REcordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'login' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_login") = "y"
	Set rs = Nothing

	If Session("access_login") = "y" Then	'if they are an authorized user allow them to proceed
        sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'login'"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        Session.Contents.Remove("access_login")

		Dim sHackMsg, sMsgText
		
		sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
		If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))
		
		If sHackMsg = vbNullString Then
            sRole = Request.Form.Item("role")

            Select Case sRole
                Case "staff"
				    sql = "SELECT StaffID, FirstName, LastName, Email FROM Staff WHERE UserID = '" & sUserName & "' AND Password = '" & sPassword & "' "
                    sql = sql & " AND Active = 'y'"
				    Set rs = conn.Execute(sql)
                    If rs.BOF and rs.EOF Then
                        bNotFound = True
                    Else
                        Session("role") = "staff"
                        Session("staff_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                        bNotFound = False
                    End If
				    Set rs = Nothing

                    'log this login and then redirect
                    If bNotFound = False Then
                        sql = "INSERT INTO StaffLogin (StaffID, WhenVisit, IPAddress, Browser) VALUES (" & Session("staff_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn.Execute(sql)
                        Set rs=Nothing

                        Response.Redirect "/staff/profile.asp"
                    Else
                        sLoginErr = "I am sorry.  Those credentials were not found for that role."
                    End If
                Case "admin"
			        If sUserName = "bobbabuoy" And sPassword = "Zeroto@123" Then 
                        Session("role") = "admin"
                        Session("my_name") = "Bob Schneider"
                        Session("my_email") = "bob.schneider@gopherstateevents.com"
                        bNotFound = False
			        ElseIf sUserName = "solveig" And sPassword = "colaianni" Then
                        Session("role") = "admin"
                        Session("my_name") = "Solveig Colianni"
                        Session("my_email") = "solveigkc@gmail.com"
                        bNotFound = False
                    Else
                        bNotFound = True
			        End If

                    If bNotFound = False Then
                        sql = "INSERT INTO AdminLogin (AdminName, WhenVisit, IPAddress, Browser) VALUES ('" & Session("my_name") & "', '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn.Execute(sql)
                        Set rs=Nothing

                        Response.Redirect "/ccmeet_admin/meets.asp" 
                    Else
                        sLoginErr ="I am sorry.  Those credentials were not found for that role."
                    End If
                Case "coach"
                    Set rs = Server.CreateObject("ADODB.Recordset")
		            sql = "SELECT CoachesID, FirstName, LastName, Email FROM Coaches WHERE UserID = '" & sUserName & "' AND Password = '" & sPassword & "'"
				    rs.Open sql, conn2, 1, 2
                    If rs.RecordCount > 0 Then
			            Session("role") = "coach"
			            Session("my_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                        bNotFound = False
                    Else
                        bNotFound = True
                    End If
                    rs.Close
		            Set rs = Nothing

                    If bNotFound = False Then
                        sql = "INSERT INTO CoachLogin (CoachesID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn2.Execute(sql)
                        Set rs=Nothing
                       
                        Response.Redirect "/cc_meet/coach/meets/lineup_mgr.asp"
                    Else
                        'check for team staff login
		                sql = "SELECT TeamStaffID,  FirstName, LastName, Email, CoachesID FROM TeamStaff WHERE UserName = '" & sUserName 
                        sql = sql & "' AND Password = '" & sPassword & "'"
				        Set rs = conn2.Execute(sql)
                        If rs.BOF and rs.EOF Then
                            bNotFound = True
                        Else
			                Session("role") = "team_staff"
			                Session("my_id") = rs(0).Value
                            Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                            Session("my_email") = rs(3).Value
                            Session("team_coach_id") = rs(4).Value
                            bNotFound = False
                        End If
		                Set rs = Nothing

                        If bNotFound = False Then
                            sql = "INSERT INTO TeamStaffLogin (TeamStaffID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                            sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                            Set rs=conn2.Execute(sql)
                            Set rs=Nothing
                       
                            Response.Redirect "/cc_meet/coach/meets/lineup_mgr.asp"
                        End If
                    End If

                    If bNotFound = True Then sLoginErr ="I am sorry.  Those credentials were not found for that role."
                Case "meet_dir"
		            sql = "SELECT MeetDirID,  FirstName, LastName, Email FROM MeetDir WHERE UserID = '" & sUserName & "' AND Password = '" 
                    sql = sql & sPassword & "'"
                    Set rs = conn2.Execute(sql)
                    If rs.BOF and rs.EOF Then
                        bNotFound = True
                    Else
			            Session("role") = "meet_dir"
			            Session("my_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                    End If
		            Set rs = Nothing

                    If bNotFound = False Then
                        sql = "INSERT INTO MeetDirLogin (MeetDirID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn2.Execute(sql)
                        Set rs=Nothing

                        Response.Redirect "/cc_meet/meet_dir/meet_dir_home.asp"
                    Else
                        sLoginErr ="I am sorry.  Those credentials were not found for that role."
                    End If
                Case "event_dir"
		            sql = "SELECT EventDirID,  FirstName, LastName, Email FROM EventDir WHERE UserID = '" & sUserName & "' AND Password = '" & sPassword & "'"
                    Set rs = conn.Execute(sql)
                    If rs.BOF and rs.EOF Then
                        bNotFound = True
                    Else
			            Session("role") = "event_dir"
			            Session("my_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                    End If
		            Set rs = Nothing
 					
                    If bNotFound = False Then
                        sql = "INSERT INTO EventDirLogin (EventDirID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn.Execute(sql)
                        Set rs=Nothing
                       
                        Response.Redirect "/events/event_dir/event_dir_home.asp"
                    Else
                        sLoginErr ="I am sorry.  Those credentials were not found for that role."
                    End If
            End Select
		End If
	End If
End If

'log this user if they are just entering the site
If Session("access_login") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'login')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If
%>
<!--#include file = "../includes/clean_input.asp" -->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events User Login</title>
<meta name="description" content="Login to your Gopher State Events account">

<script>
    function chkFlds() {
        if (document.site_login.user_name.value == '' ||
        document.site_login.password.value == '' ||
        document.site_login.role.value == '') {
            alert('All fields are required!');
            return false;
        }
        else
            return true;
    }
</script>
</head>

<body>
<div class="container">
   	<!--#include file = "../includes/header.asp" -->

    <h3 class="h3">Gopher State Events Login</h3>
    <br>
    <%If Not sLoginErr = vbNullString Then%>
        <p class="bg-danger"><%=sLoginErr%></p>
    <%End If%>
 
    <div class="row">
        <div class="col-sm-2" style="text-align:center;">
            <h3><%=dTimeNow%></h3>
            <a href="login.asp"><img class="img-responsive" src="/graphics/get_time.png" alt="Refresh Time"></a>
        </div>
        <div class="col-sm-6">
            <form role="form" class="form-horizontal" name="site_login" method="Post" action="login.asp" 
            onsubmit="return chkFlds();">
            <div class="row form-group">
                <label class="control-label col-sm-4" for="user_name">User Name:</label>
                <div class="col-sm-8">
                    <input type="text" class="form-control form-control-sm" name="user_name" id="user_name" value="<%=sUserName%>" placeholder="form-control-sm">
                </div>
            </div>
            <div class="row form-group">
                <label class="control-label col-sm-4" for="password">Password:</label>
                <div class="col-sm-8">
                    <input type="password" class="form-control form-control-sm" name="password" id="password" maxlength="12" value="" placeholder="form-control-sm">
                </div>
            </div>
            <div class="row form-group">
                <label class="control-label col-sm-4" for="role">Role:</label>
                <div class="col-sm-8">
                    <select  class="form-control form-control-sm" name="role" id="role" placeholder="form-control-sm">
                        <option value="">&nbsp;</option>
                        <option value="admin">GSE Administrator</option>
                        <option value="coach">CC/Nordic Coach</option>
                        <option value="staff">GSE Staff</option>
                        <option value="meet_dir">CC/Nordic Meet Director</option>
                        <option value="event_dir">Fitness Event Director</option>
                    </select>
                </div>
            </div>
            <div class="row form-group">
                <input  class="form-control form-control-sm" type="hidden" name="submit_login" id="submit_login" value="submit_login">
                <input  class="form-control form-control-sm" type="submit" name="submit1" id="submit1" value="Login">
            </div>
            </form>

            <div class="row">
                <a href="javascript:pop('/misc/forgot_login.asp',600,750)">Forgot Sign In?</a>
            </div>
        </div>
        <div class="col-sm-4">
            <img src="/graphics/custom_bib_small.png" alt="Custom Bib Image" class="img-responsive">
        </div>
    </div>
</div>
<!--#include file = "../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
