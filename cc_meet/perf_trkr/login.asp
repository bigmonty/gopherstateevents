<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, conn3, rs, sql
Dim i
Dim lPartID
Dim sPartName, sMyGender, sErrMsg, sMsg, sUserName, sPassword
Dim cdoMessage, cdoConfig
Dim bNotFound

Session.Contents.RemoveAll()

lPartID = Request.QueryString("part_id")
If CStr(lPartID) & "" = "" Then lPartID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_login") = "submit_login" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'login' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn3, 1, 2
	If rs.RecordCount > 0 Then Session("access_login") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_login") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg
		
		sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
		If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))
		
		If sHackMsg = vbNullString Then
			sql = "SELECT p.PerfTrkrID, r.FirstName, r.LastName, p.Email, r.RosterID, r.Gender "
            sql = sql & "FROM PerfTrkr p INNER JOIN Roster r ON p.RosterID = r.RosterID "
            sql = sql & "WHERE p.UserName = '" & sUserName & "' AND p.Password = '" & sPassword & "'"
            Set rs = conn.Execute(sql)
            If rs.BOF and rs.EOF Then
                bNotFound = True
            Else
                Session("perf_trkr_id") = rs(0).value
                Session("role") = "perf_trkr"
                Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                Session("my_email") = rs(3).Value
                Session("roster_id") = rs(4).value
                Session("gender") = rs(5).value
            End If
			Set rs = Nothing

            'get team name
            If Not CStr(Session("roster_id")) = vbNullString Then
                sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN Roster r ON t.TeamsID = r.TeamsID "
                sql = sql & "WHERE r.RosterID = " & Session("roster_id")
                Set rs = conn.Execute(sql)
                Session("team_id") = rs(0).value
                Session("my_name") = Replace(rs(1).Value, "''", "'")
                Set rs = Nothing
            End If

			If Not CStr(Session("perf_trkr_id")) = vbNullString Then 
                sql = "INSERT INTO PerfTrkrLogin (PerfTrkrID, WhenVisit, IPAddress, Browser) VALUES (" & Session("perf_trkr_id") & ", '" & Now() 
                sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                Set rs=conn.Execute(sql)
                Set rs=Nothing

	            Session.Contents.Remove("access_login")

                Response.Redirect "/cc_meet/perf_trkr/perf_trkr.asp"
            Else
                sErrMsg ="I am sorry.  Those credentials were not found.  Please try again or create an account <a href='create_accnt.asp'>here</a>."
            End If
		End If
    End If
End If

'log this user if they are just entering the site
If Session("access_login") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'login')"
	Set rs = conn3.Execute(sql)
	Set rs = Nothing
End If
%>
<!--#include file = "../../includes/cdo_connect.asp" -->

<!--#include file = "../../includes/clean_input.asp" -->
<%
Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Performance Tracker Login</title>
<meta name="description" content="Performance Tracker login for a Gopher State Events (GSE).">
<script>
function chkFlds() {
if (document.login.user_name.value == '' || 
    document.login.password.value == '') 
{
 	alert('All fields are required!');
 	return false
 	}
else
 	return true;
}
</script>
</head>
<body onload="document.login.user_name.focus();">
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->

    <div class="row">
        <div class="col-md-4">
	        <%If Not sHackMsg = vbNullString Then%>
		        <p class="text-danger"><%=sHackMsg%></p>
	        <%Else%>
 			    <div class="bg-warning" style="text-align: center;">
				    <a style="color:#fff;" href="javascript:pop('forgot_signin.asp',600,550)">Forgot Sign In?</a>
                    |
                    <a style="color:#fff;" href="create_accnt.asp?part_id=<%=lPartID%>">Create Account</a>
                </div>

               <h4 class="h4">Log in</h4>

		        <%If Not sErrMsg = vbNullString Then%>
			        <p class="bg-success"><%=sErrMsg%></p>
		        <%End If%>

  			    <form role="form" class="form"  name="login" method="Post" action="login.asp" onSubmit="return chkFlds();">
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
	        <%End If%>
            <div>
                Performance Tracker is SCHOOL CROSS-COUNTRY/NORDIC SKI PARTICIPANTS ONLY utility that affords subscribers
                the ability to follow and compare the performances of themselves, teammates and competitors.  These competitors are followed via 
                "packs" (a pack can be a single participant or a group of participants that are of the same gender and of the same sport).  
                <br><br>
                Among other features, it allows you to have your <span style="font-weight: bold;">results emailed or texted to you, your parents, and
                friends within a few minutes of finishing your race</span> (NOTE:  ONLY if the event was timed by GSE).
            </div>

            <div class="bg-warning">
                This service carries a one-time fee of $5 to cover administrative and server fees.  You will receive a 7-day free trial which begins when you 
                create your account.
            </div>

            <div class="bg-success">
                This service does not divulge any information that is not already "public" 
                via online results lists.  What it does do is make those results more personal and formats them in a manner that is more informational.
                <br>
                <span style="font-weight: bold;">Please be patient with us.  This utility is a work in progress.</span>
            </div>
        </div>
        <div class="col-md-8">
            <div class="row">
                 <div class="col-md-6">
                    <h4 class="h4">What's The Point?</h4>
                    <ul class="list-group">
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Sends results texts/emails within minutes (GSE-timed events only) to you and whoever you wish to receive them (parents, siblings, etc.).</li>
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Create "Performance Packs" of competitors to follow.</li>
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Track yours and your competitors' performances.</li>
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Graphs of yours and your competitors' performances.</li>
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Access to <a href="http://www.my-etraxc.com">My-eTRaXC</a> training utility.</li>
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Enter performances in events not timed by GSE.</li>
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Compare performances graphically over time.</li>
                        <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Network with willing participants.</li>
                    </ul>
                     <h4 class="h4">About Performance Tracker</h4>
                    <iframe class="embed-responsive-item" src="https://www.youtube.com/embed/8pZQckevn70" frameborder="0" allowfullscreen></iframe>
                </div>
                <div class="col-md-6">
                    <img src="images/sample_cc.jpg" alt="Sample Picture" class="img-responsive">
                    <br>
		            <a href="http://www.etraxc.com/" onclick="openThis2(this.href,1024,760);return false;">
		                <img src="/graphics/banner_ads/etraxc_banner.png" alt="eTRaXC" class="img-responsive">
                    </a>
                    <hr>
                    <a href="http://www.my-etraxc.com/" onclick="openThis2(this.href,1024,760);return false;">
		                <img src="/graphics/my-etraxc_ad.gif" alt="My-eTRaXC" class="img-responsive">
                    </a>
                    <hr>
                    <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                    <!-- GSE Banner Ad -->
                    <ins class="adsbygoogle"
                            style="display:inline-block;width:350px;height:90px"
                            data-ad-client="ca-pub-1381996757332572"
                            data-ad-slot="1411231449"></ins>
                    <script>
                    (adsbygoogle = window.adsbygoogle || []).push({});
                    </script>
                </div>
           </div>
        </div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" --> 
<%
conn3.Close
Set conn3 = Nothing

conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>
</body>
</html>
