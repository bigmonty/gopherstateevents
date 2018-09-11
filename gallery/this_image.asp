<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim lEventID, lImageID
Dim sEventName, sPixName, sCaption
Dim dEventDate

lEventID = Request.QueryString("event_id")
lImageID = Request.QueryString("image_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PixName, Caption FROM RacePix WHERE RacePixID = " & lImageID
rs.Open sql, conn, 1, 2
sPixName = rs(0).Value
If Not rs(1).Value & "" = "" Then sCaption = Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

If sCaption = vbNullString Then sCaption = "No caption available."
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xasp1/DTD/xhtm1-transitional.dtd">

<html lang="en">
<head>
<title><%=sEventName%> (<%=dEventDate%>) Image Viewer</title>
<!--#include file = "../includes/meta2.asp" -->



	
<script>
<!--
var i=0;
function resize() {
  if (navigator.appName == 'Netscape') i=40;
  if (document.images[0]) window.resizeTo(document.images[0].width +350, document.images[0].height+185-i);
  self.focus();
}
//-->
</script>
	
<script>
  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-20252412-1']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();
</script>
</head>

<body onload="resize();">
<h4 style="text-align:center;color:#fff;background-color:<%=sPrimColor%>;margin-bottom:10px;"><%=sEventName%> (<%=dEventDate%>) Image Viewer</h4>
<div style="width:200px;float:left;margin:0 0 0 10px;padding:50px 10px 0 0;position:absolute;color:<%=sPrimColor%>;">
	<h5 style="text-align:left;">Caption:</h5>
	<%=sCaption%>
</div>
<div style="margin:10px 10px 10px 250px;float:right;">
	<a href="javascript:window.close();"> 
	<img src="/gallery/<%=lEventID%>/<%=sPixName%>" alt="<%=sPixName%>" title="the title" width="640"></a> 
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>

