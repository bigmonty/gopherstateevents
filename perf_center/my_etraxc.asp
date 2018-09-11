<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

If CStr(Session("my_hist_id")) = vbNullString Then Response.Redirect "my_hist_login.asp"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>My GSE&copy; My-eTRaXC</title>
<meta name="description" content="My eTRaXC description for a Gopher State Events (GSE) timed event.">

<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap-theme.min.css">
<link rel="stylesheet" href="dist/css/bootstrap-submenu.min.css">

<script src="https://code.jquery.com/jquery-2.1.4.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>

<script src="/misc/scripts.js"></script>
<script>
$(document).ready(function() {
    // Now the document is ready, you can do your stuff!
    $('button').click(function() {
        //get the relevant content
        var text = $('#' + $(this).data('content')).html();
        $('#data_target').html(text);
    });
});
</script>

<style type="text/css">
    button{
        width: 150px;margin: 0 0 5px 0;
    }
</style>
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="My GSE History Portal">
    <h3 class="h3">My GSE History</h3>

    <!--#include file = "my_hist_nav.asp" -->

    <h4 class="h4">About My-eTRaXC</h4>

    <div class="bg-success">
        <a href="http://www.my-etraxc.com" style="font-weight:bold;" onclick="openThis2(this.href,1024,768);return false;">My-eTRaXC</a> is a completely free online training & fitness 
        management utility that interfaces well with My GSE history.  When you created your My GSE History account you also created your My-eTRaXC
        account.
    </div>
    
    <div class="col-sm-2">   
        <div style="margin-top: 10px;"> 
            <button data-content="content1" type="button" class="btn btn-primary">Training</button><br>
            <button data-content="content2" type="button" class="btn btn-info">Goal Setting</button><br>
            <button data-content="content3" type="button" class="btn btn-success">Race Results</button><br>
            <button data-content="content4" type="button" class="btn btn-danger">Nutrition</button><br>
            <button data-content="content5" type="button" class="btn btn-warning">Mobile-Friendly</button><br>
            <button data-content="content6" type="button" class="btn btn-danger">Media</button><br>
            <button data-content="content7" type="button" class="btn btn-primary">Statistics</button><br>
            <button data-content="content8" type="button" class="btn btn-info">Personal Logs</button><br>   
            <button data-content="content9" type="button" class="btn btn-success">Fitness Manager</button>
        </div>
    </div>
    <div class="col-sm-4">
        To view some of the functionality of each feature, click on the button for that feature.
        <div style="display:none;">
            <div id="content1">
                <h4 class="h4 bg-primary">Training Functionality</h4>
                <ul class="bg-primary">
                    <li>Training Log</li>
                    <li>Run Training</li>
                    <li>Cross-Training</li>
                    <li>Training Statistics</li>
                    <li>Virtual Training Groups</li>
                    <li>Workout Scheduler</li>
                </ul>
            </div>
            <div id="content2">
                <h4 class="h4 bg-info">Goal Setting Functionality</h4>
                <ul class="bg-info">
                    <li>Performance Goals</li>
                    <li>Training Goals</li>
                    <li>Fitness Goals</li>
                    <li>Nutrition Goals</li>
                </ul>
            </div>
            <div id="content3">
                <h4 class="h4 bg-success">Race Results Functionality</h4>
                <ul class="bg-success">
                    <li>Import GSE Race Results</li>
                    <li>Import External Race Results</li>
                </ul>
            </div>
            <div id="content4">
               <h4 class="h4 bg-danger">Nutrition Functionality</h4>
                <ul class="bg-danger">
                    <li>Nutrition Log</li>
                    <li>
                        Nutrition Management
                        <ul>
                            <li>Calorie Counter</li>
                            <li>Birds-Eye Overview</li>
                        </ul>
                    </li>
                    <li>Protein-Fat-Carb Breakdown Manager</li>
                </ul>
            </div>
            <div id="content5">
                <h4 class="h4 bg-warning">Mobile Functionality</h4>
                <ul class="bg-warning">
                    <li>Mobile-Responsive Design (in progress)</li>
                    <li>Mobile App (in progress)</li>
                </ul>
            </div>
            <div id="content6">
                <h4 class="h4 bg-danger">Nutrition Functionality</h4>
                <ul class="bg-danger">
                    <li>Link to your external picture galleries and video repositories.</li>
                </ul>
            </div>
            <div id="content7">
                <h4 class="h4 bg-primary">Training Stats Functionality</h4>
                <ul class="bg-primary">
                    <li>Track your training stats by day, week, month, season, and year</li>
                    <li>Compare yourself to other My eTRaXC users.</li>
                </ul>
            </div>
            <div id="content8">
                <h4 class="h4 bg-info">Personal Log Functionality</h4>
                <ul class="bg-info">
                    <li>Personal Log</li>
                    <li>Training Log</li>
                    <li>Fitness Log</li>
                    <li>Nutrition Log</li>
                    <li>Sleep Log</li>
                    <li>Shoe Log</li>
                </ul>
            </div>
            <div id="content9">
                <h4 class="h4 bg-success">Fitness Management Functionality</h4>
                <ul class="bg-success">
                    <p>Create a fitnes journey that encompasses training and nutrition, graphs, and fitness timeline (with images).</p>
                </ul>
            </div>
        </div>
        <div id="data_target"></div>
    </div>
    <div class="col-sm-6">
        <div class="bg-warning">
            <h4 class="h4">The My eTRaXC Mission</h4>

            <p>
                The mission of My-eTRaXC is to create an individual fitness management utility that is contemporary in its look and feel and includes 
                networking and social connectivity for the participants. It is an athletes-only partner site to eTRaXC for ease of use.  eTRaXC is a 
                team-based utility but use of My eTRaXC by individuals not connected to a team is encouraged.

		        <a href="http://www.my-etraxc.com/" onclick="openThis2(this.href,1024,760);return false;">
		            <img src="/graphics/my-etraxc_ad.gif" alt="My-eTRaXC" class="img-responsive">
                </a>
            </p>
        </div>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
