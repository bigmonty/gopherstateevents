<%@ Language=VBScript%>

<%
Option Explicit

Session("event_id") = Request.QueryString("event_id")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" 
"http://www.w3.org/TR/html4/strict.dtd">

<html lang="en">
<head>
<title>Gopher State Events Upload Proofs</title>



</head>

<body>
<div style="font-size:0.75em;width:350px;text-align:center;margin:10px;background-color:#fff;">
	<h4 style="padding:5px;">Upload Proof Sheet</h4>
	
	<h5 style="margin:0;padding:0;">Guidelines for Uploading Images</h5>
	<form name="upload" method="Post" action="receive_proofs.asp?event_id=<%=Session("event_id")%>" enctype="multipart/form-data">
	<input type="file" name="file1" id="file1" size="50">
	<br>
	<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
	<input type="submit" id="submit1" name="submit1" value="Upload!">
	</form>
</div>
</body>
</html>