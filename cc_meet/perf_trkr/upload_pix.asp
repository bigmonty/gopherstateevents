<%@ Language=VBScript%>

<%
Option Explicit

Session("this_user") = Request.QueryString("this_user")

If Request.Form.Item("submit_this") = "submit_this" Then Response.Redirect "receive_pix.asp"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Upload GSE Performance Tracker Profile Picture</title>
<meta name="description" content="Forgot My GSE History signin information, a fitness event timing Service for road racing, nordic ski, showshoe, mountain bike, and cross-country meet timing.">
<!--#include file = "../../includes/js.asp" --> 
</head>

<body>
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->

    <h3 class="h3">Upload Profile Picture</h3>

	<form class="form-inline" name="upload" method="Post" action="receive_pix.asp" enctype="multipart/form-data">
	(Note-this will overwrite any picture already on file for you!)
	<br>
	<input class="form-control" type="file" name="file1" id="file1" size="50">
	<br>
	<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
	<input class="form-control" type="submit" id="submit1" name="submit1" value="Upload!">
	</form>
</div>
</body>
</html>