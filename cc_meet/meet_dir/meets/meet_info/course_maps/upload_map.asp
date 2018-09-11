<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lMeetID, sMeetName
Dim sSent
Dim cdoMessage, cdoConfig
Dim sMsg

If Not (Session("role") = "meet_dir" Or Session("role") = "admin") Then Response.Redirect "/default.asp?sign_out=y"

lMeetID = Request.QueryString("meet_id")
sSent = Request.QueryString("sent")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If CStr(sSent) = CStr("y") Then
	sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
	Set rs = conn.Execute(sql)
	sMeetName = Replace(rs(0).Value, "''", "'")
	Set rs = Nothing
	
	sMsg = vbCrLf
	sMsg = sMsg & "A course map has been uploaded for " & sMeetName & vbCrLf & vbCrLf
		
%>
<!--#include file = "../../../../../includes/cdo_connect.asp" -->
<%

	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.com"
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "GSE CCMeet Course Map Upload"
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
	
	Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../../includes/meta2.asp" -->
<title>GSE Upload Course Map</title>
<!--#include file = "../../../../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<h4 class="h4">Upload Course Map</h4>
	<form class="form-inline" name="upload" method="Post" action="receive_file.asp?meet_id=<%=lMeetID%>" enctype="multipart/form-data">
	<input type="FILE" class="form-control" name="File1" id="File1">
	<input type="submit" class="form-control" id="submit_1" name="submit_1" value="Upload!">
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
