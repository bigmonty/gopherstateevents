<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i
Dim sEmail
Dim cdoMessage, cdoConfig
Dim DontSend

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
												
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_email") = "submit_email" Then
	sEmail = Request.Form.Item("email")

    If ValidEmail(sEmail) = True Then
	    sql = "INSERT INTO DontSend(Email, WhenEntered) VALUES ('" & sEmail & "', '" & Now() & "')"
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
			.TextBody = sEmail & " has been added to the GSE Dont Send list."
			.Send
		End With
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing
    End If
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DontSendID, Email, WhenEntered, DontSend FROM DontSend ORDER BY WhenEntered DESC, Email"
rs.Open sql, conn, 1, 2
DontSend = rs.GetRows()
rs.Close
Set rs = Nothing

%>
<!--#include file = "../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE Dont Send</title>
<meta name="description" content="Do not send page for GSE (Gopher State Events).">

<!--#include file = "../includes/js.asp" -->

<style type="text/css">
<!--
p, input{
	font-size:0.85em;
	}
-->
</style>

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

	<div id="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">Gopher State Events "Do Not Send"</h4>

	        <form name="login" method="post" action="admin_dont_send.asp" onSubmit="return chkFlds();">
	        <span style="font-weight:bold;font-size:0.9em;">Email address:</span>
	        <input type="text" name="email" id="email" size="30">
	        <input type="hidden" name="submit_email" id="submit_email" value="submit_email">
	        <input type="submit" name="submit1" id="submit1" value="Submit Email">
	        </form>    

            <br>

            <h4 class="h4">"Do Not Send" List</h4>

            <h5>Num Records: <%=UBound(DontSend, 2) + 1%></h5>
            <table style="font-size: 0.9em;">
                <tr>
                    <th>No.</th><th>Email</th><th>When Submitted</th><th>Dont Send</th>
                </tr>
                <%For i = 0 To UBound(DontSend, 2)%>
                    <%If i mod 2 = 0 Then%>
                        <tr>
                            <td class="alt"><%=i + 1%>)</td>
                            <td class="alt"><%=DontSend(1, i)%></td>
                            <td class="alt"><%=DontSend(2, i)%></td>
                            <td class="alt"><%=DontSend(3, i)%></td>
                        </tr>
                    <%Else%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td><%=DontSend(1, i)%></td>
                            <td><%=DontSend(2, i)%></td>
                            <td><%=DontSend(3, i)%></td>
                        </tr>
                    <%End If%>
                <%Next%>
            </table>
        </div>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>