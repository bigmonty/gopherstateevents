<%@ Language=VBScript %>
<%
Option Explicit

Dim sShowAvg, sShowWhat

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

sShowAvg = Request.QueryString("show_avg")
If sShowAvg = vbNullString Then sShowAvg = "n"

sShowWhat = Request.QueryString("show_what")
If sShowWhat = vbNullString Then sShowWhat = "events"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Trends: Event Trends By Race Size</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-sm-10">
            <!--#include file = "trends_menu.asp" -->

            <%Select Case sShowWhat%>
                <%Case "events"%>
                    <h3 class="h3">Event Trends By Race Size (Number of Events)</h3>
                <%Case "parts"%>
                    <%If sShowAvg = "y" Then%>
                        <h3 class="h3">Event Trends By Race Size (Average Number of Finishers)</h3>
                    <%Else%>
                        <h3 class="h3">Event Trends By Race Size (Number of Finishers)</h3>
                    <%End If%>
           <%End Select%>

            <ul class="nav">
                <li class="nav-item"><a class="nav-link" href="size_trends.asp?show_what=events">Num Events</a></li>
                <li class="nav-item"><a class="nav-link" href="size_trends.asp?show_what=parts">Num Parts</a></li>
                <li class="nav-item"><a class="nav-link" href="size_trends.asp?show_avg=y&amp;show_what=parts">Avg Num Parts</a></li>
            </ul>

            <div class="embed-responsive embed-responsive-16by9">
				<iframe name="size_graph" id="size_graph" frameborder="0" 
					src="size_graph.asp?show_avg=<%=sShowAvg%>&amp;show_what=<%=sShowWhat%>" style="width:800px;height:400px;"></iframe>
            </div>
        </div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
