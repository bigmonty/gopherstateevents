<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i
Dim sEmail, sDontSend
Dim cdoMessage, cdoConfig

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
												
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_email") = "submit_email" Then
	sEmail = Request.Form.Item("email")
    sDontSend = Request.Form.Item("dont_send")

    If ValidEmail(sEmail) = True Then
	    sql = "INSERT INTO DontSend(Email, WhenEntered, DontSend) VALUES ('" & sEmail & "', '" & Now() & "', '" & sDontSend & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        'notify me
%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = "bob.schneider@gopherstateevents.com"
			.From = "bob.schneider@gopherstateevents.com"
			.Subject = "New Entry in Dont Send List"
			.TextBody = sEmail & " has been added to the GSE Dont Send list. Dont send = " & sDontSend
			.Send
		End With
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing
    End If
End If

%>
<!--#include file = "../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Dont Send</title>

<script>
function chkFlds() {
if (document.enter_email.email.value == '') 
{
 	alert('You must submit an email!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

    <div class="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-sm-10">
			<h1 class="h1">Gopher State Events "Do Not Send"</h1>
	
            <p>To respect your privacy, once you enter your email address below and click the button, that address will 
                no longer receive any emails from Gopher State Events for the selected category.  If anyone else is using this email address on entry forms, 
                etc, they will also not receive any further emails from this address.</p>

            <div class="form-group">
	            <form class="form-inline" name="login" method="post" action="do_not_send.asp" onSubmit="return chkFlds();">
	            <label for="email">Email address:</label>
	            <input type="text" class="form-control" name="email" id="email" size="30">
                <select class="form-control"  name="dont_send" id="dont_send">
                    <option value="all">Don't Send Me Anything</option>
                    <option value="pre-race">Don't Send Me Pre-Race Emails</option>
                    <option value="promo">Don't Send Me Event Promotional Emails</option>
                    <option value="results">Don't Send Me My Results</option>
                </select>
	            <input class="form-control"  type="hidden" name="submit_email" id="submit_email" value="submit_email">
	            <input class="form-control"  type="submit" name="submit1" id="submit1" value="Submit Email">
	            </form>    
            </div>                        			
        </div>
    </div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>