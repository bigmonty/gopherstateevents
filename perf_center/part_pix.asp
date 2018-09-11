<%@ Language=VBScript%>

<%
Option Explicit

%>
<<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>My GSE History Profile Picture Upload</title>
<meta name="description" content="Upload my profile picture for My GSE History account.">
<!--#include file = "../includes/js.asp" -->
</head>
</head>

<body>
<img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
<div class="container">		
    <h3 class="h3">My GSE History Picture Uploader</h3>

	<h5 class="h5">Select Picture:</h5>
	<form name="upload" method="Post" action="receive_pix.asp" enctype="multipart/form-data">
	<input type="FILE" name="File1" id="File1" size="50">
	<input type="submit" id="submit2" name="submit2" value="Upload!">
	</form>
</div>
</body>

</html>
