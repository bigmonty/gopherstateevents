<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lFeaturedEventID
Dim sEventName, sBannerImage
Dim dEventDate

If Not Session("role") = "admin" Then 
    If Not Session("role") = "event_dir" Then Response.Redirect "http://www.google.com"
End If

lFeaturedEventID = Request.QueryString("featured_event_id")
Session("featured_event_id") = lFeaturedEventID     'this is needed because I can't pass a query string to the file receive page

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate, BannerImage FROM FeaturedEvents WHERE FeaturedEventsID = " & lFeaturedEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sBannerImage = rs(2).Value
rs.Close
Set rs = Nothing
%>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Featured Event Banner Image Upload</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
<div class="container">
	<h3 class="h3">Featured Event Banner Image Upload:<br><%=sEventName%> on <%=dEventDate%></h3>

    <h4 class="h4">Existing Image</h4>	
    <%If sBannerImage & "" = "" Then%>
        <p>No image available.</p>
    <%Else%>
        <img src="/featured_events/images/<%=sBannerImage%>" alt="Banner Image" class="img-responsive">
    <%End If%>
	
    <br>

    <h4 class="h4">Upload New Image:</h4>
	<form name="upload" method="Post" action="receive_banner.asp" enctype="multipart/form-data">
	<input type="file" name="file1" id="file1" size="50">
	<br>
	<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
	<input type="submit" id="submit1" name="submit1" value="Upload!">
	</form>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>