<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, m, j
Dim lMyHistID
Dim sSubject, sMsg, sErrMsg, sUserName, sPassword
Dim MyHist(), AttachArr(), EmailArr()
Dim cdoMessage, cdoConfig
Dim bRecipients

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2	
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	sSubject = Request.Form.Item("subject")
	sMsg = Request.Form.Item("message")
	bRecipients = False
	
	i = 0
	ReDim EmailArr(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MyHistID, FirstName, LastName, Email FROM PartData WHERE MyHistID IS NOT NULL ORDER BY LastName, FirstName"
	rs.Open sql, conn2, 1, 2
	Do While Not rs.EOF
		If Request.Form.Item("MyHist_" & rs(0).Value) = "on" Then
			EmailArr(0, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
			EmailArr(1, i) = rs(3).Value
			EmailArr(2, i) = rs(0).Value
			i = i + 1
			ReDim Preserve EmailArr(2, i)

			If bRecipients = False Then bRecipients = True
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing

	If bRecipients = False Then
		sErrMsg = "Please select at least one recipient."
	Else
%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
		For i = 0 to UBound(EmailArr, 2) - 1
			sSubject = Request.Form.Item("subject")
			
			sMsg = Request.Form.Item("message") & vbCrLf & vbCrLf
			
			If Request.Form.Item("send_login") = "on" Then
				Call GetLogin(EmailArr(2, i))
				sMsg = sMsg & "Login Information:" & vbCrLf
				sMsg = sMsg & "User Name: " & sUserName & vbCrLf
				sMsg = sMsg & "Password: " & sPassword & vbCrLf & vbCrLf
			End If
			
			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = EmailArr(1, i)
				.BCC = "bob.schneider@gopherstateevents.com"
				.From = "support@gopherstateevents.com"
				.Subject = sSubject
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
		Next
	
		Set cdoConfig = Nothing
	End If
End If

i = 0
ReDim MyHist(2, 0)
sql = "SELECT MyHistID, FirstName, LastName, Email FROM PartData WHERE MyHistID IS NOT NULL ORDER BY LastName, FirstName"
Set rs = conn2.Execute(sql)
Do While Not rs.EOF
    If Not CLng(rs(0).Value) = 0 Then
	    MyHist(0, i) = rs(0).Value
	    MyHist(1, i) = Replace(rs(1).Value, "''","'") & " " & Replace(rs(2).Value, "''", "'")
	    MyHist(2, i) = rs(3).Value
	    i = i + 1
	    ReDim Preserve MyHist(2, i)
    End If
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub GetLogin(lMyHistID)
	sql = "SELECT UserName, Password FROM MyHist WHERE MyHistID = " & lMyHistID
	Set rs = conn.Execute(sql)
	sUserName = rs(0).Value
	sPassword = rs(1).Value
	Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Contact My History Accounts</title>

<!--#include file = "../../includes/js.asp" -->

</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <!--#include file = "my_hist_nav.asp" -->

		    <h4 class="h4">Contact My History Accounts</h4>
			
			<%If Not sErrMsg = vbNullString then%>
				<p><%=sErrMsg%></p>
			<%End If%>
			
            <form name="send_email" method="post" action="contact.asp">
   			<table>
				<tr>
                    <td valign="top">
                        <ul style="width: 150px;list-style:none;margin-left: 0;padding-left: 0;">
						    <%For i = 0 to UBound(MyHist, 2) - 1%>
								<li>
									<input type="checkbox" name="MyHist_<%=MyHist(0, i)%>" id="MyHist_<%=MyHist(0, i)%>">
									<a href="mailto:<%=MyHist(2, i)%>"><%=MyHist(1, i)%></a>
								</li>
						    <%Next%>
                        </ul>
					</td>
					<td valign="top" style="background-color:#ececd8;padding: 5px;">
						<h4 class="h4">Create Message</h4>

						<table style="width:550px;font-size: 1.0em;">
							<tr>
								<td style="text-align:right" valign="top">Subject:</td>
								<td><input type="text" name="subject" id="subject" size="50" value="<%=sSubject%>"></td>
							</tr>
							<tr>
								<td style="text-align:right;white-space:nowrap;">Send Login Info:</td>
								<td style="text-align:left">
									<input type="checkbox" name="send_login" id="send_login">
									Send site login information with email.
								</td>
							</tr>
							<tr>
								<td style="text-align:right" valign="top">Message:</td>
								<td><textarea name="message" id="message" rows="15" cols="60" style="font-size:1.35em;"><%=sMsg%></textarea></td>
							</tr>
							<tr>
								<td style="text-align:center" colspan="2">
									<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
									<input type="submit" name="submit2" id="submit2" value="Send">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		    </form>
   		</div>
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
