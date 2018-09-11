<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim Testim

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Testimonial, Author, Role FROM Testimonials ORDER BY NewID()"
rs.Open sql, conn, 1, 2
Testim = rs.GetRows()
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta name="viewport" content="width=device-width, initial-scale=1">

<link rel="alternate" href="http://gopherstateevents.com" hreflang="en-us" />
<link rel="shortcut icon" href="/assets/images/g-transparent2-351x345.png" type="image/x-icon">
<link rel="stylesheet" href="/assets/web/assets/mobirise-icons/mobirise-icons.css">
<link rel="stylesheet" href="/assets/tether/tether.min.css">
<link rel="stylesheet" href="/assets/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="/assets/bootstrap/css/bootstrap-grid.min.css">
<link rel="stylesheet" href="/assets/bootstrap/css/bootstrap-reboot.min.css">
<link rel="stylesheet" href="/assets/socicon/css/styles.css">
<link rel="stylesheet" href="/assets/dropdown/css/style.css">
<link rel="stylesheet" href="/assets/theme/css/style.css">
<link rel="stylesheet" href="/assets/mobirise/css/mbr-additional.css" type="text/css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.6.0/css/bootstrap-datepicker.css">

<script src="/assets/web/assets/jquery/jquery.min.js"></script>

<script src="/misc/scripts.js"></script>

<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-56760028-1', 'auto');
  ga('send', 'pageview');
</script>
<title>Gopher State Events Testimonials</title>
<meta name="description" content="Testimonials for Gopher State Events (GSE) chip/rfid event timing service in Minnetonka, MN.">
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div class="row">
		<div class="col-sm-10">
			<h4 class="h4">Gopher State Events Testimonials</h4>

			<%For i = 0 To UBound(Testim, 2)%>
                <%=Testim(0, i)%>
               <p><%=Testim(1, i) %><br><%=Testim(2, i)%></p>
                <hr>
            <%Next%>
		</div>
		<!--#include file = "../includes/cmng_evnts.asp" -->
	</div>
<!--#include file = "../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>