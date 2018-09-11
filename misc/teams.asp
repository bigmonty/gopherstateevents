<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>How GSE&copy; Team Scoring Works</title>
<meta name="description" content="How the Gopher State Events (GSE) series works.">
<!--#include file = "../includes/js.asp" -->

<style type="text/css">
    li{
        padding-top: 5px;
    }
</style>
</head>

<body>
<div class="container">
    <img class="img=responsive" src="/graphics/html_header.png" alt="Team Header">
	<div id="row">
        <h2 class="h2">How the GSE Team Utility Works</h2>

        <p>In an attempt to continue to stay abreast of our connected society and the way that it impacts fitness events, all of our events now have the
        option of addigng a team component.  This could be a competitive utility or simply a way to allow folks to share the entire experience. And best of
        all, this feature carries no additional charge for event directors or participants!</p>

        <p>The team component of a race can be scored or not scored and can add competitive or simply participatory awards, or none at all.  If the team
        component of an event is scored, the details below can add some clarity to the team scoring options.</p>
    
        <p>NOTE:  In events that have more than one race all members of a team must be entered into the same race.</p>

        <div style="padding-top: 10px;">
            <h4 style="text-align: left;background: none;border: none;">Scoring Parameters:</h4>
            <p>Our system is very flexible.  It allows you to set parameters for how you want your system set up and/or scored.</p>
            <ul>
                <li>MINIMUM PARTICIPANTS TO SCORE: What is the fewest number of participants that a team must have to score?</li>
                <li>MAXIMUM PARTICIPANTS TO SCORE: What is the most number of participants that a team can have to factor in their scoring?</li>
                <li>SCORING METHOD: You can score your teams using one of the below methods.</li>
            </ul>
            <br>
            <h4 style="text-align: left;background: none;border: none;">Scoring Legend:</h4>
            <ul>
                <li>"CUMULATIVE" represents the total combined time for the scoring members of a team.  This method is not useful when there is not a fixed
                    number of scoring members.</li>
                <li>"AVERAGE" represents the average time for each scoring member of a team.  It is useful when there is not a fixed number of scoring members.</li>
                <li>"SCORE" is derived by adding the points up for each scoring finisher in the team results (similar to cross country running).  In this case
                the lowest score wins.</li>
                <li>"POINTS" is derived by assigning a maximum number of points to the first team finisher and one less point for each 
                finisher from there.  In this case, the highest score wins.</li>
                <li>"DNF" (Did Not Finish) implies that a team did not have the minimum number of finishers to compile a score or that an individual participant
                    did not finish the race.</li>
            </ul>
        </div>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
