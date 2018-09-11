<%@ Language=VBScript%>

<%
Option Explicit

Dim i
Dim ASPErr
Dim ErrArray(23)
Dim bDontSend

bDontSend = False

On Error Resume Next
'Response.Clear
Set ASPErr = Server.GetLastError

ErrArray(0) = "Domain: Gopher State Events"
ErrArray(1) = "Organization: " & vbNullString
ErrArray(2) = "Site User: " & vbNullString
ErrArray(3) = "Role: " & vbNullString
ErrArray(4) = "Email: " & vbNullString
ErrArray(5) = "When Occur: " & Now()
ErrArray(6) = "Web Page: " & "//" & Request.ServerVariables ("SERVER_NAME") & ASPErr.File
ErrArray(7) = "Line: " & ASPErr.Line
ErrArray(8) = "Column: " & ASPErr.Column
ErrArray(9) = "Source: " & ASPErr.Source
ErrArray(10) = "Description: " & ASPErr.Description
ErrArray(11) = "Err #: " & ASPErr.ASPCode
ErrArray(12) = "Com Err #: " & ASPErr.Number
ErrArray(13) = "Category: " & ASPErr.Category
ErrArray(14) = "URL: " & Request.ServerVariables("URL")
ErrArray(15) = "ASP Descr: " & ASPErr.ASPDescription
ErrArray(16) = "REQUEST_METHOD: " & Request.ServerVariables("REQUEST_METHOD")
ErrArray(17) = "SERVER_PORT: " & Request.ServerVariables("SERVER_PORT")
ErrArray(18) = "HTTPS: " & Request.ServerVariables("HTTPS")
ErrArray(19) = "LOCAL_ADDR: " & Request.ServerVariables("LOCAL_ADDR")
ErrArray(20) = "REMOTE_ADDR: " & Request.ServerVariables("REMOTE_ADDR")
ErrArray(21) = "HTTP_USER_AGENT: " & Request.ServerVariables("HTTP_USER_AGENT")
ErrArray(22) = "Query STring: " & Request.QueryString
ErrArray(23) = "Form: " & Request.Form

If Session("role") = "admin" Then bDontSend = True

If bDontSend = False Then
    If Not ErrArray(21) & "" = "" Then
 	    If Not (InStr(ErrArray(21), "bot") = 0 AND InStr(ErrArray(21), "spider") = 0 AND InStr(ErrArray(21), "Presto") = 0) Then bDontSend = True
     End If
End If

If bDontSend = False Then
    Dim cdoMessage, cdoConfig
    Dim sMsgText
		
	sMsgText = "A 500 error has been returned on the Gopher State Events domain.  The details are as follows: " & vbCrLf & vbCrLf
		
	For i = 0 To UBound(ErrArray)
		sMsgText = sMsgText & ErrArray(i) & vbCrLf
	Next
		
%>
<!--#include file = "includes/cdo_connect.asp" -->
<%
		
	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.com"
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "GSE Error Notice"
		.TextBody = sMsgText
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "includes/meta2.asp" -->
<title>Gopher State Events&reg; 500 Error Page</title>
<meta name="description" content="A 500 error has been returned on a GSE page.">
<!--#include file = "includes/js.asp" -->
</head>

<body>
<div class="container">
	<div id="header">
		<img src="/graphics/error_hdr.png" alt="GSE">
	</div>
	
	<div style="margin:10px 10px 25px 10px;padding:25px;background-color:#fff;width:600px;">
		<h2>Didn't see that coming...</h2>

		<p style="font-size:1.5em;margin-top:25px;">Oops!  Possibly a ton of people are looking for results right now or
            maybe something unexpected happened.  Either way, we have been notified!  <span style="color: red;">You might get what you want by refreshing 
            your browser.</span>  Otherwise, please try again later.  GSE (Gopher State Events) thanks you for your patience!</p>
		
		<%If Session("role") = "admin" Then%>
			<ul>
				<%For i = 5 To 23%>
					<li><%=ErrArray(i)%></li>
				<%Next%>
			</ul>
		<%End If%>
	</div>
</div>
</body>
</html>
