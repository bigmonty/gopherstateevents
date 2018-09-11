<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim sEmail, sPageStatus, sUserName, sPassword, sMsg
Dim cdoMessage, cdoConfig
	
Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("get_signin") = "get_signin" Then
	sEmail = Request.Form.Item("email")
	sPageStatus = "not_found"
			
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UserName, Password FROM MyHist WHERE Email = '" & sEmail & "'"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		sUserName = rs(0).value
		sPassword = rs(1).Value
		sPageStatus = "found"
	End If
    rs.Close
	Set rs = Nothing
	
	conn.Close
	Set conn = Nothing
	
	If sPageStatus = "found" Then
		sMsg = vbCrLf
		sMsg = sMsg & "You are receiving this email because a request for login information for My GSE History "
		sMsg = sMsg & "(www.gopherstateevents.com) was requested sent to this address.  If you did not make this request, please "
		sMsg = sMsg & "notify us immediately at 612.720.8427 or by sending an email to bob.schneider@gopherstateevents.com." & vbCrLf & vbCrLf
		
		sMsg = sMsg & "Here is your login information: " & vbCrLf
		sMsg = sMsg & "Your User Name is: " & sUserName & vbCrLf
		sMsg = sMsg & "Your Password is: " & sPassword & vbCrLf & vbCrLf
		
		sMsg = sMsg & "Sincerely~" & vbCrLf
		sMsg = sMsg & "Bob Schneider" & vbCrLf
		sMsg = sMsg & "GSE Administrator"
		
%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = sEmail
			.CC = "bob.schneider@gopherstateevents.com"
			.From = "bob.schneider@gopherstateevents.com"
			.Subject = "My GSE History Lost Login"
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
<!--#include file = "../includes/meta2.asp" -->
<title>Forgot My GSE History Signin</title>
<meta name="description" content="Forgot My GSE History signin information, a fitness event timing Service for road racing, nordic ski, showshoe, mountain bike, and cross-country meet timing.">
<!--#include file = "../includes/js.asp" --> 
</head>
<body onload="document.forgot_login.email.focus();">
<div class="container" style="margin: 10px;">
    <img src="/graphics/html_header.png" class="img-responsive" alt="My GSE History Portal">
    <h3 class="h3">My GSE History</h3>

    <div class="row">
        <h4 class="h4">Forgot My GSE History Sign-in</h4>
	
	    <%If sPageStatus = "found" Then%>
           <div class="text-danger">
		        <p>The login information has been sent to address requested.  Please notify us via 
		        <a href="mailto:bob.schneider@gopherstateevents.com">email</a> if you do not receive it.</p>
            </div>
	    <%ElseIf sPageStatus = "not_found" Then%>
           <div class="text-danger">
		        <p>We are sorry but there was no match in our database for this email address.  Please contact us via 
		        <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by phone at 612.720.8427 to resolve this 
		        discrepancy.</p>
            </div>
	    <%End If%>

        <%If Not sPageStatus = "found" Then%>
            <div class="bg-success">
		        <p>
                    Please enter your email address below and click the "Get Signin" button.  If the email address you entered matches one in our
                    database your login information will be sent to that address.  If it does not exist, you will receive a message indicating that.
                </p>  
            </div>
		    <form role="form" class="form" name="forgot_login" method="post" action="forgot_signin.asp">
			<div class="form-group">
				<label for="email">Email:</label>
				<input type="text" class="form-control" name="email" id="email" maxLength="50" size="25"value="<%=sEmail%>" tabindex="1"> 
			</div>

		    <input type="hidden" class="form-control" name="get_signin" id="get_signin" value="get_signin">
		    <input type="submit" class="form-control" name="submit1" id="submit1" value="Get Signin" tabindex="2">
		    </form>
        <%End If%>
    </div>
</div>
</body>
</html>
