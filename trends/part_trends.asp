<%@ Language=VBScript %>
<%
Option Explicit

Dim sShowAvg

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

sShowAvg = Request.QueryString("show_avg")
If sShowAvg = vbNullString Then sShowAvg = "n"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Trends: Participant Trends</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-sm-10">
            <!--#include file = "trends_menu.asp" -->

            <%If sShowAvg = "y" Then%>
                <h3 class="h3">Participant Trends: Averages</h3>
            <%Else%>
                <h3 class="h3">Participant Trends: Totals</h3>
            <%End If%>

            <ul class="nav">
                <li class="nav-item"><a class="nav-link" href="part_trends.asp">Totals</a></li>
                <li class="nav-item"><a class="nav-link" href="part_trends.asp?show_avg=y">Averages</a></li>
                <li class="nav-item"><a class="nav-link" href="part_ytd.asp">Year to Date</a></li>
                <li class="nav-item"><a class="nav-link" href="part_ytd.asp?show_avg=<%=sShowAvg%>">Year to Date Averages</a></li>
            </ul>
            
            <div class="embed-responsive embed-responsive-16by9">
				<iframe name="part_graph" id="part_graph" frameborder="0" 
					src="part_graph.asp?show_avg=<%=sShowAvg%>" style="width:800px;height:400px;"></iframe>
            </div>
        </div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
